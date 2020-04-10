VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpRequisicoesNFOcx 
   ClientHeight    =   5730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8145
   ScaleHeight     =   5730
   ScaleWidth      =   8145
   Begin VB.Frame Frame1 
      Caption         =   "Requisitantes"
      Height          =   4020
      Index           =   2
      Left            =   435
      TabIndex        =   44
      Top             =   1350
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Frame FrameCodigo 
         Caption         =   "Código"
         Height          =   705
         Left            =   315
         TabIndex        =   51
         Top             =   450
         Width           =   3210
         Begin MSMask.MaskEdBox CodRequisitanteDe 
            Height          =   300
            Left            =   525
            TabIndex        =   12
            Top             =   240
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodRequisitanteAte 
            Height          =   300
            Left            =   2160
            TabIndex        =   13
            Top             =   240
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin VB.Label LabelCodRequisitanteAte 
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
            Left            =   1695
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   53
            Top             =   300
            Width           =   360
         End
         Begin VB.Label LabelCodRequisitanteDe 
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
            Left            =   180
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   52
            Top             =   300
            Width           =   315
         End
      End
      Begin VB.Frame FrameNome 
         Caption         =   "Nome"
         Height          =   675
         Left            =   315
         TabIndex        =   48
         Top             =   1605
         Width           =   5160
         Begin MSMask.MaskEdBox NomeReqDe 
            Height          =   300
            Left            =   525
            TabIndex        =   16
            Top             =   255
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NomeReqAte 
            Height          =   300
            Left            =   3060
            TabIndex        =   17
            Top             =   240
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin VB.Label LabelNomeReqDe 
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
            Left            =   165
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   50
            Top             =   315
            Width           =   315
         End
         Begin VB.Label LabelNomeReqAte 
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
            Left            =   2625
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   49
            Top             =   315
            Width           =   360
         End
      End
      Begin VB.Frame FrameCcl 
         Caption         =   "Centro de Custo"
         Height          =   705
         Left            =   330
         TabIndex        =   45
         Top             =   2925
         Visible         =   0   'False
         Width           =   3210
         Begin MSMask.MaskEdBox CclDe 
            Height          =   315
            Left            =   540
            TabIndex        =   14
            Top             =   240
            Visible         =   0   'False
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CclAte 
            Height          =   315
            Left            =   2145
            TabIndex        =   15
            Top             =   240
            Visible         =   0   'False
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin VB.Label LabelCclAte 
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
            Left            =   1710
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   47
            Top             =   300
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Label LabelCclDe 
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
            Left            =   165
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   46
            Top             =   300
            Visible         =   0   'False
            Width           =   315
         End
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpRequisicoesNFOcx.ctx":0000
      Left            =   945
      List            =   "RelOpRequisicoesNFOcx.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   90
      Width           =   2685
   End
   Begin VB.CommandButton BotaoExecutar 
      Caption         =   "Executar"
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
      Left            =   3870
      Picture         =   "RelOpRequisicoesNFOcx.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   135
      Width           =   1635
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5820
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpRequisicoesNFOcx.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpRequisicoesNFOcx.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpRequisicoesNFOcx.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpRequisicoesNFOcx.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOrdenacao 
      Height          =   315
      ItemData        =   "RelOpRequisicoesNFOcx.ctx":0A9A
      Left            =   1560
      List            =   "RelOpRequisicoesNFOcx.ctx":0AA7
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   465
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4200
      Index           =   1
      Left            =   345
      TabIndex        =   24
      Top             =   1305
      Width           =   6495
      Begin VB.Frame Frame3 
         Caption         =   "Requisições"
         Height          =   2895
         Left            =   90
         TabIndex        =   30
         Top             =   1170
         Width           =   6255
         Begin VB.Frame Frame5 
            Caption         =   "Código"
            Height          =   690
            Left            =   180
            TabIndex        =   41
            Top             =   1935
            Width           =   5775
            Begin MSMask.MaskEdBox CodRequisicaoDe 
               Height          =   300
               Left            =   645
               TabIndex        =   10
               Top             =   270
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CodRequisicaoAte 
               Height          =   300
               Left            =   3690
               TabIndex        =   11
               Top             =   270
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin VB.Label LabelCodRequisicaoAte 
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
               Left            =   3285
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   43
               Top             =   315
               Width           =   360
            End
            Begin VB.Label LabelCodRequisicaoDe 
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
               Left            =   270
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   42
               Top             =   330
               Width           =   315
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Data Limite"
            Height          =   690
            Left            =   180
            TabIndex        =   36
            Top             =   1125
            Width           =   5775
            Begin MSComCtl2.UpDown UpDownDataLimiteAte 
               Height          =   315
               Left            =   4860
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   270
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataLimiteAte 
               Height          =   315
               Left            =   3690
               TabIndex        =   9
               Top             =   270
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataLimiteDe 
               Height          =   315
               Left            =   1800
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   270
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataLimiteDe 
               Height          =   315
               Left            =   630
               TabIndex        =   8
               Top             =   270
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin VB.Label LabelDataLimiteAte 
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
               Left            =   3285
               TabIndex        =   40
               Top             =   315
               Width           =   360
            End
            Begin VB.Label LabelDataLimiteDe 
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
               Left            =   270
               TabIndex        =   39
               Top             =   315
               Width           =   315
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Data Envio"
            Height          =   690
            Left            =   180
            TabIndex        =   31
            Top             =   315
            Width           =   5775
            Begin MSComCtl2.UpDown UpDownDataEnvioAte 
               Height          =   315
               Left            =   4860
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   225
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataEnvioAte 
               Height          =   315
               Left            =   3690
               TabIndex        =   7
               Top             =   240
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataEnvioDe 
               Height          =   315
               Left            =   1800
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   225
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataEnvioDe 
               Height          =   315
               Left            =   615
               TabIndex        =   6
               Top             =   240
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin VB.Label LabelDataEnvioDe 
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
               Left            =   270
               TabIndex        =   35
               Top             =   315
               Width           =   315
            End
            Begin VB.Label LabelDataEnvioAte 
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
               Left            =   3285
               TabIndex        =   34
               Top             =   315
               Width           =   360
            End
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Filial Empresa"
         Height          =   1065
         Left            =   90
         TabIndex        =   25
         Top             =   45
         Width           =   6255
         Begin MSMask.MaskEdBox CodigoFilialDe 
            Height          =   300
            Left            =   1140
            TabIndex        =   2
            Top             =   285
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodigoFilialAte 
            Height          =   300
            Left            =   4230
            TabIndex        =   3
            Top             =   240
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NomeFilialAte 
            Height          =   300
            Left            =   4230
            TabIndex        =   5
            Top             =   615
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NomeFilialDe 
            Height          =   300
            Left            =   1140
            TabIndex        =   4
            Top             =   645
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin VB.Label LabelCodigoDe 
            AutoSize        =   -1  'True
            Caption         =   "Código De:"
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
            Left            =   165
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   29
            Top             =   323
            Width           =   960
         End
         Begin VB.Label LabelNomeDe 
            AutoSize        =   -1  'True
            Caption         =   "Nome De:"
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
            Left            =   270
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   28
            Top             =   705
            Width           =   855
         End
         Begin VB.Label LabelCodigoAte 
            AutoSize        =   -1  'True
            Caption         =   "Código Até:"
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
            Left            =   3195
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   27
            Top             =   315
            Width           =   1005
         End
         Begin VB.Label LabelNomeAte 
            AutoSize        =   -1  'True
            Caption         =   "Nome Até:"
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
            Left            =   3300
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   26
            Top             =   675
            Width           =   900
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4605
      Left            =   225
      TabIndex        =   54
      Top             =   990
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   8123
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Requisições"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Requisitante"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Left            =   270
      TabIndex        =   56
      Top             =   135
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Ordenados Por:"
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
      Left            =   240
      TabIndex        =   55
      Top             =   510
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "RelOpRequisicoesNFOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'RelOpRequisicoesNF
Const ORD_POR_CODIGO = 0
Const ORD_POR_LIMITE = 1
Const ORD_POR_ENVIO = 2


Private WithEvents objEventoCodRequisitanteDe As AdmEvento
Attribute objEventoCodRequisitanteDe.VB_VarHelpID = -1
Private WithEvents objEventoCodRequisitanteAte As AdmEvento
Attribute objEventoCodRequisitanteAte.VB_VarHelpID = -1
Private WithEvents objEventoCclDe As AdmEvento
Attribute objEventoCclDe.VB_VarHelpID = -1
Private WithEvents objEventoCclAte As AdmEvento
Attribute objEventoCclAte.VB_VarHelpID = -1
Private WithEvents objEventoNomeReqDe As AdmEvento
Attribute objEventoNomeReqDe.VB_VarHelpID = -1
Private WithEvents objEventoNomeReqAte As AdmEvento
Attribute objEventoNomeReqAte.VB_VarHelpID = -1
Private WithEvents objEventoNomeFilialDe As AdmEvento
Attribute objEventoNomeFilialDe.VB_VarHelpID = -1
Private WithEvents objEventoNomeFilialAte As AdmEvento
Attribute objEventoNomeFilialAte.VB_VarHelpID = -1
Private WithEvents objEventoCodFilialDe As AdmEvento
Attribute objEventoCodFilialDe.VB_VarHelpID = -1
Private WithEvents objEventoCodFilialAte As AdmEvento
Attribute objEventoCodFilialAte.VB_VarHelpID = -1
Private WithEvents objEventoCodRequisicaoDe As AdmEvento
Attribute objEventoCodRequisicaoDe.VB_VarHelpID = -1
Private WithEvents objEventoCodRequisicaoAte As AdmEvento
Attribute objEventoCodRequisicaoAte.VB_VarHelpID = -1

Dim iFrameAtual As Integer
Dim iAlterado As Integer
Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 73942
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 73943

    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 73942
        
        Case 73943
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172898)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()
    
    Unload Me
    
End Sub

Private Sub Limpa_Tela_Rel()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Rel
  
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 73944
    
    ComboOrdenacao.ListIndex = 0
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_Limpa_Tela_Rel:
    
    Select Case gErr
    
        Case 73944
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172899)

    End Select

    Exit Sub
   
End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela_Rel

End Sub


Public Sub Form_Load()

Dim lErro As Long
Dim sMascaraCcl As String

On Error GoTo Erro_Form_Load
    
    Set objEventoCodRequisitanteDe = New AdmEvento
    Set objEventoCodRequisitanteAte = New AdmEvento
        
    Set objEventoNomeReqDe = New AdmEvento
    Set objEventoNomeReqAte = New AdmEvento
        
    Set objEventoCclDe = New AdmEvento
    Set objEventoCclAte = New AdmEvento
        
    Set objEventoCodFilialDe = New AdmEvento
    Set objEventoCodFilialAte = New AdmEvento
    
    Set objEventoNomeFilialDe = New AdmEvento
    Set objEventoNomeFilialAte = New AdmEvento
        
    Set objEventoCodRequisicaoDe = New AdmEvento
    Set objEventoCodRequisicaoAte = New AdmEvento
        
    lErro = MascaraCcl(sMascaraCcl)
    If lErro <> SUCESSO Then gError 73945

    CclDe.Mask = sMascaraCcl
    CclAte.Mask = sMascaraCcl
    
    iFrameAtual = 1
    
    ComboOrdenacao.ListIndex = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 73945
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172900)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoCodRequisitanteDe = Nothing
    Set objEventoCodRequisitanteAte = Nothing
    
    Set objEventoNomeReqDe = Nothing
    Set objEventoNomeReqAte = Nothing
    
    Set objEventoCclDe = Nothing
    Set objEventoCclAte = Nothing
    
    Set objEventoCodFilialDe = Nothing
    Set objEventoCodFilialAte = Nothing
    
    Set objEventoNomeFilialDe = Nothing
    Set objEventoNomeFilialAte = Nothing
    
    Set objEventoCodRequisicaoDe = Nothing
    Set objEventoCodRequisicaoAte = Nothing
    
End Sub

Private Sub CodigoFilialAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoFilialAte, iAlterado)
    
End Sub

Private Sub CodigoFilialDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoFilialDe, iAlterado)
    
End Sub

Private Sub CodRequisicaoAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodRequisicaoAte, iAlterado)
    
End Sub

Private Sub CodRequisicaoDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodRequisicaoDe, iAlterado)
    
End Sub

Private Sub CodRequisitanteAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodRequisitanteAte, iAlterado)
    
End Sub

Private Sub CodRequisitanteDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodRequisitanteDe, iAlterado)
    
End Sub

Private Sub DataEnvioAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataEnvioAte, iAlterado)
    
End Sub

Private Sub DataEnvioDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataEnvioDe, iAlterado)
    
End Sub

Private Sub DataLimiteAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataLimiteAte, iAlterado)
    
End Sub

Private Sub DataLimiteDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataLimiteDe, iAlterado)
    
End Sub

Private Sub LabelCodRequisitanteAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objRequisitante As New ClassRequisitante

On Error GoTo Erro_LabelCodRequisitanteAte_Click

    If Len(Trim(CodRequisitanteAte.Text)) > 0 Then
        'Preenche com o requisitante da tela
        objRequisitante.lCodigo = StrParaLong(CodRequisitanteAte.Text)
    End If

    'Chama Tela RequisitanteLista
    Call Chama_Tela("RequisitanteLista", colSelecao, objRequisitante, objEventoCodRequisitanteAte)

   Exit Sub

Erro_LabelCodRequisitanteAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172901)

    End Select

    Exit Sub

End Sub

Private Sub DataEnvioDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEnvioDe_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataEnvioDe.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataEnvioDe.Text)
    If lErro <> SUCESSO Then gError 73946

    Exit Sub
                   
Erro_DataEnvioDe_Validate:

    Cancel = True

    Select Case gErr

        Case 73946
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172902)

    End Select

    Exit Sub

End Sub

Private Sub DataEnvioAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEnvioAte_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataEnvioAte.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataEnvioAte.Text)
    If lErro <> SUCESSO Then gError 73947

    Exit Sub
                   
Erro_DataEnvioAte_Validate:

    Cancel = True

    Select Case gErr

        Case 73947
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172903)

    End Select

    Exit Sub

End Sub

Private Sub DataLimiteAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataLimiteAte_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataLimiteAte.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataLimiteAte.Text)
    If lErro <> SUCESSO Then gError 73948

    Exit Sub
                   
Erro_DataLimiteAte_Validate:

    Cancel = True

    Select Case gErr

        Case 73948
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172904)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip1_Click()
    
    'Se frame atual corresponde ao tab selecionado, sai da rotina
    If TabStrip1.SelectedItem.Index = iFrameAtual Then Exit Sub

    'Torna Frame correspondente ao Tab selecionado visivel
    Frame1(TabStrip1.SelectedItem.Index).Visible = True

    'Torna Frame atual invisivel
    Frame1(iFrameAtual).Visible = False

    'Armazena novo valor de iFrameAtual
    iFrameAtual = TabStrip1.SelectedItem.Index

End Sub

Private Sub UpDownDataEnvioDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvioDe_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEnvioDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 73949

    Exit Sub

Erro_UpDownDataEnvioDe_DownClick:

    Select Case gErr

        Case 73949
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 172905)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEnvioDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvioDe_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEnvioDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 73950

    Exit Sub

Erro_UpDownDataEnvioDe_UpClick:

    Select Case gErr

        Case 73950
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 172906)

    End Select

    Exit Sub

End Sub
Private Sub UpDownDataEnvioAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvioAte_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEnvioAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 73951

    Exit Sub

Erro_UpDownDataEnvioAte_DownClick:

    Select Case gErr

        Case 73951
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 172907)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEnvioAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvioAte_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEnvioAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 73952

    Exit Sub

Erro_UpDownDataEnvioAte_UpClick:

    Select Case gErr

        Case 73952
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 172908)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataLimiteAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataLimiteAte_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataLimiteAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 73953

    Exit Sub

Erro_UpDownDataLimiteAte_DownClick:

    Select Case gErr

        Case 73953
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 172909)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataLimiteAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataLimiteAte_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataLimiteAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 73954

    Exit Sub

Erro_UpDownDataLimiteAte_UpClick:

    Select Case gErr

        Case 73954
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 172910)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataLimiteDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataLimiteDe_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataLimiteDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 73955

    Exit Sub

Erro_UpDownDataLimiteDe_DownClick:

    Select Case gErr

        Case 73955
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 172911)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataLimiteDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataLimiteDe_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataLimiteDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 73956

    Exit Sub

Erro_UpDownDataLimiteDe_UpClick:

    Select Case gErr

        Case 73956
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 172912)

    End Select

    Exit Sub

End Sub

Private Sub DataLimiteDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataLimiteDe_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataLimiteDe.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataLimiteDe.Text)
    If lErro <> SUCESSO Then gError 73957

    Exit Sub
                   
Erro_DataLimiteDe_Validate:

    Cancel = True

    Select Case gErr

        Case 73957
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172913)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigoDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelCodigoDe_Click

    If Len(Trim(CodigoFilialDe.Text)) > 0 Then
        'Preenche com a FilialEmpresa da tela
        objFilialEmpresa.iCodFilial = StrParaInt(CodigoFilialDe.Text)
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoCodFilialDe)

   Exit Sub

Erro_LabelCodigoDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172914)

    End Select

    Exit Sub

End Sub
Private Sub LabelCodigoAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelCodigoAte_Click

    If Len(Trim(CodigoFilialAte.Text)) > 0 Then
        'Preenche com a FilialEmpresa da tela
        objFilialEmpresa.iCodFilial = StrParaInt(CodigoFilialAte.Text)
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoCodFilialAte)

   Exit Sub

Erro_LabelCodigoAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172915)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodRequisicaoAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objReqCompras As New ClassRequisicaoCompras

On Error GoTo Erro_LabelCodRequisicaoAte_Click

    If Len(Trim(CodRequisicaoAte.Text)) > 0 Then
        'Preenche com o requisitante da tela
        objReqCompras.lCodigo = StrParaLong(CodRequisicaoAte.Text)
    End If

    'Chama Tela ReqComprasTodasLista
    Call Chama_Tela("ReqComprasTodasLista", colSelecao, objReqCompras, objEventoCodRequisicaoAte)

   Exit Sub

Erro_LabelCodRequisicaoAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172916)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodRequisicaoDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objReqCompras As New ClassRequisicaoCompras

On Error GoTo Erro_LabelCodRequisicaoDe_Click

    If Len(Trim(CodRequisicaoDe.Text)) > 0 Then
        'Preenche com o requisitante da tela
        objReqCompras.lCodigo = StrParaLong(CodRequisicaoDe.Text)
    End If

    'Chama Tela ReqComprasTodasLista
    Call Chama_Tela("ReqComprasTodasLista", colSelecao, objReqCompras, objEventoCodRequisicaoDe)

   Exit Sub

Erro_LabelCodRequisicaoDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172917)

    End Select

    Exit Sub

End Sub



Private Sub LabelCodRequisitanteDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objRequisitante As New ClassRequisitante

On Error GoTo Erro_LabelCodRequisitanteDe_Click

    If Len(Trim(CodRequisitanteDe.Text)) > 0 Then
        'Preenche com o requisitante da tela
        objRequisitante.lCodigo = StrParaLong(CodRequisitanteDe.Text)
    End If

    'Chama Tela RequisitanteLista
    Call Chama_Tela("RequisitanteLista", colSelecao, objRequisitante, objEventoCodRequisitanteDe)

   Exit Sub

Erro_LabelCodRequisitanteDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172918)

    End Select

    Exit Sub

End Sub

Private Sub LabelCclAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCcl As New ClassCcl
Dim sCclFormata As String
Dim iCclPreenchida As Integer

On Error GoTo Erro_LabelCclAte_Click

    If Len(Trim(CclAte.Text)) > 0 Then
        
        lErro = CF("Ccl_Formata", CclAte.Text, sCclFormata, iCclPreenchida)
        If lErro <> SUCESSO Then gError 73958
        
        'Preenche com o Ccl
        objCcl.sCcl = sCclFormata
        
    End If

    'Chama Tela Cclista
    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCclAte)

   Exit Sub

Erro_LabelCclAte_Click:

    Select Case gErr

        Case 73958
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172919)

    End Select

    Exit Sub

End Sub

Private Sub LabelCclDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCcl As New ClassCcl
Dim sCclFormata As String
Dim iCclPreenchida As Integer

On Error GoTo Erro_LabelCclDe_Click

    If Len(Trim(CclDe.Text)) > 0 Then
        
        lErro = CF("Ccl_Formata", CclDe.Text, sCclFormata, iCclPreenchida)
        If lErro <> SUCESSO Then gError 73959

        'Preenche com o Ccl
        objCcl.sCcl = sCclFormata
        
    End If

    'Chama Tela Cclista
    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCclDe)

   Exit Sub

Erro_LabelCclDe_Click:

    Select Case gErr

        Case 73959
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172920)

    End Select

    Exit Sub

End Sub

Private Sub LabelNomeDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelNomeDe_Click

    If Len(Trim(NomeFilialDe.Text)) > 0 Then
        'Preenche com o requisitante da tela
        objFilialEmpresa.sNome = NomeFilialDe.Text
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoNomeFilialDe)

   Exit Sub

Erro_LabelNomeDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172921)

    End Select

    Exit Sub

End Sub
Private Sub LabelNomeAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelNomeAte_Click

    If Len(Trim(NomeFilialAte.Text)) > 0 Then
        'Preenche com a FilialEmpresa da tela
        objFilialEmpresa.sNome = NomeFilialAte.Text
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoNomeFilialAte)

   Exit Sub

Erro_LabelNomeAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172922)

    End Select

    Exit Sub

End Sub

Private Sub LabelNomeReqDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objRequisitante As New ClassRequisitante

On Error GoTo Erro_LabelNomeReqDe_Click

    If Len(Trim(NomeReqDe.Text)) > 0 Then
        'Preenche com o requisitante da tela
        objRequisitante.sNomeReduzido = NomeReqDe.Text
    End If

    'Chama Tela RequisitanteLista
    Call Chama_Tela("RequisitanteLista", colSelecao, objRequisitante, objEventoNomeReqDe)

   Exit Sub

Erro_LabelNomeReqDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172923)

    End Select

    Exit Sub

End Sub

Private Sub LabelNomeReqAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objRequisitante As New ClassRequisitante

On Error GoTo Erro_LabelNomeReqAte_Click

    If Len(Trim(NomeReqAte.Text)) > 0 Then
        'Preenche com o requisitante da tela
        objRequisitante.sNomeReduzido = NomeReqAte.Text
    End If

    'Chama Tela RequisitanteLista
    Call Chama_Tela("RequisitanteLista", colSelecao, objRequisitante, objEventoNomeReqAte)

   Exit Sub

Erro_LabelNomeReqAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172924)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCodFilialAte_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    CodigoFilialAte.Text = CStr(objFilialEmpresa.iCodFilial)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoCodFilialDe_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    CodigoFilialDe.Text = CStr(objFilialEmpresa.iCodFilial)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoCodRequisicaoAte_evSelecao(obj1 As Object)

Dim objReqCompras As New ClassRequisicaoCompras

    Set objReqCompras = obj1

    CodRequisicaoAte.Text = CStr(objReqCompras.lCodigo)

    Me.Show

End Sub

Private Sub objEventoCodRequisicaoDe_evSelecao(obj1 As Object)

Dim objReqCompras As New ClassRequisicaoCompras

    Set objReqCompras = obj1

    CodRequisicaoDe.Text = CStr(objReqCompras.lCodigo)

    Me.Show

End Sub

Private Sub objEventoCodRequisitanteDe_evSelecao(obj1 As Object)

Dim objRequisitante As New ClassRequisitante

    Set objRequisitante = obj1

    CodRequisitanteDe.Text = CStr(objRequisitante.lCodigo)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoCodRequisitanteAte_evSelecao(obj1 As Object)

Dim objRequisitante As New ClassRequisitante

    Set objRequisitante = obj1

    CodRequisitanteAte.Text = CStr(objRequisitante.lCodigo)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNomeFilialAte_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    NomeFilialAte.Text = objFilialEmpresa.sNome

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNomeFilialDe_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    NomeFilialDe.Text = objFilialEmpresa.sNome

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNomeReqDe_evSelecao(obj1 As Object)

Dim objRequisitante As New ClassRequisitante

    Set objRequisitante = obj1

    NomeReqDe.Text = objRequisitante.sNomeReduzido

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNomeReqAte_evSelecao(obj1 As Object)

Dim objRequisitante As New ClassRequisitante

    Set objRequisitante = obj1

    NomeReqAte.Text = objRequisitante.sNomeReduzido

    Me.Show

    Exit Sub

End Sub
Private Sub objEventoCclDe_evSelecao(obj1 As Object)
'traz o ccl selecionado para a tela

Dim lErro As Long
Dim objCcl As ClassCcl
Dim sCclMascarado As String

On Error GoTo Erro_objEventoCclDe_evSelecao

    Set objCcl = obj1

    lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then gError 73960

    CclDe.PromptInclude = False
    CclDe.Text = sCclMascarado
    CclDe.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoCclDe_evSelecao:

    Select Case gErr

        Case 73960

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172925)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCclAte_evSelecao(obj1 As Object)
'traz o ccl selecionado para a tela

Dim lErro As Long
Dim objCcl As ClassCcl
Dim sCclMascarado As String

On Error GoTo Erro_objEventoCclAte_evSelecao

    Set objCcl = obj1

    lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then gError 73961

    CclAte.PromptInclude = False
    CclAte.Text = sCclMascarado
    CclAte.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoCclAte_evSelecao:

    Select Case gErr

        Case 73961

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172926)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 73962

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 73963

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 73964
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 73965
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 72962
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 73963 To 73965
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172927)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 73966

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 73967

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call Limpa_Tela_Rel
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 73966
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 73967

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172928)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 73968

    Select Case ComboOrdenacao.ListIndex

            Case ORD_POR_CODIGO
                
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "RequisicaoCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ItemReqCompra", 1)
                
            Case ORD_POR_LIMITE

                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "DataLimite", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "RequisicaoCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ItemReqCompra", 1)
                
            Case ORD_POR_ENVIO

                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "DataEnvio", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "RequisicaoCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ItemReqCompra", 1)
                
            Case Else
                gError 74971

    End Select

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 73968, 74971

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172929)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim sCodFilial_I As String
Dim sCodFilial_F As String
Dim sNomeFilial_I As String
Dim sNomeFilial_F As String
Dim sCodRequisitante_I As String
Dim sCodRequisitante_F As String
Dim sNomeReq_I As String
Dim sNomeReq_F As String
Dim sCcl_I As String
Dim sCcl_F As String
Dim sCodRequisicao_I As String
Dim sCodRequisicao_F As String
Dim sOrdenacaoPor As String
Dim iOrdenacao As Long
Dim sOrd As String
Dim sCheck As String

On Error GoTo Erro_PreencherRelOp
    
    lErro = Formata_E_Critica_Parametros(sCodRequisitante_I, sCodRequisitante_F, sNomeReq_I, sNomeReq_F, sCcl_I, sCcl_F, sCodFilial_I, sCodFilial_F, sNomeFilial_I, sNomeFilial_F, sCodRequisicao_I, sCodRequisicao_F)
    If lErro <> SUCESSO Then gError 73969

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 73970
         
    lErro = objRelOpcoes.IncluirParametro("NCODREQINIC", sCodRequisitante_I)
    If lErro <> AD_BOOL_TRUE Then gError 73971
         
    lErro = objRelOpcoes.IncluirParametro("TNOMEREQINIC", NomeReqDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 73972
    
    lErro = objRelOpcoes.IncluirParametro("TCCLINIC", sCcl_I)
    If lErro <> AD_BOOL_TRUE Then gError 73973
    
    lErro = objRelOpcoes.IncluirParametro("NCODFILIALINIC", sCodFilial_I)
    If lErro <> AD_BOOL_TRUE Then gError 73974
         
    lErro = objRelOpcoes.IncluirParametro("TNOMEFILIALINIC", NomeFilialDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 73975
    
    lErro = objRelOpcoes.IncluirParametro("NCODREQUISICAOINIC", sCodRequisicao_I)
    If lErro <> AD_BOOL_TRUE Then gError 73976
         
    lErro = objRelOpcoes.IncluirParametro("TNOMEREQUISICAOINIC", NomeReqDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 73977
    
    'Preenche dataenvio inicial
    If Trim(DataEnvioDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DENVINIC", DataEnvioDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DENVINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 73978
    
    'Preenche a data limite inicial
    If Trim(DataLimiteDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DLIMINIC", DataLimiteDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DLIMINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 73979
    
    lErro = objRelOpcoes.IncluirParametro("NCODREQFIM", sCodRequisitante_F)
    If lErro <> AD_BOOL_TRUE Then gError 73980
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEREQFIM", NomeReqAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 73981
        
    lErro = objRelOpcoes.IncluirParametro("TCCLFIM", sCcl_F)
    If lErro <> AD_BOOL_TRUE Then gError 73982
    
    lErro = objRelOpcoes.IncluirParametro("NCODFILIALFIM", sCodFilial_F)
    If lErro <> AD_BOOL_TRUE Then gError 73983
         
    lErro = objRelOpcoes.IncluirParametro("TNOMEFILIALFIM", NomeFilialAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 73984
    
    lErro = objRelOpcoes.IncluirParametro("NCODREQUISICAOFIM", sCodRequisicao_F)
    If lErro <> AD_BOOL_TRUE Then gError 73985
         
    lErro = objRelOpcoes.IncluirParametro("TNOMEREQUISICAOFIM", NomeReqAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 73986
    
    'Preenche data de envio Final
    If Trim(DataEnvioAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DENVFIM", DataEnvioAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DENVFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 73987
    
    'Preenche data limite final
    If Trim(DataLimiteAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DLIMFIM", DataLimiteAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DLIMFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 73988
    
    Select Case ComboOrdenacao.ListIndex
        
            Case ORD_POR_CODIGO
            
                sOrdenacaoPor = "CodReq"
                
            Case ORD_POR_ENVIO
                sOrdenacaoPor = "DataEnvio"
                
            Case ORD_POR_LIMITE
                sOrdenacaoPor = "DataLimite"
                
            Case Else
                gError 73989
                  
    End Select

    lErro = objRelOpcoes.IncluirParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> AD_BOOL_TRUE Then gError 73990
   
    sOrd = ComboOrdenacao.ListIndex
    lErro = objRelOpcoes.IncluirParametro("NORDENACAO", sOrd)
    If lErro <> AD_BOOL_TRUE Then gError 73991
   
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCodRequisitante_I, sCodRequisitante_F, sNomeReq_I, sNomeReq_F, sCcl_I, sCcl_F, sCodFilial_I, sCodFilial_F, sNomeFilial_I, sNomeFilial_F, sCodRequisicao_I, sCodRequisicao_F, sOrdenacaoPor, sOrd)
    If lErro <> SUCESSO Then gError 73992

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 73969 To 73992
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172930)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sCodRequisitante_I As String, sCodRequisitante_F As String, sNomeReq_I As String, sNomeReq_F As String, sCcl_I As String, sCcl_F As String, sCodFilial_I As String, sCodFilial_F As String, sNomeFilial_I As String, sNomeFilial_F As String, sCodRequisicao_I As String, sCodRequisicao_F As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long
Dim sCclFormata As String
Dim iCclPreenchida As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros
       
    'critica Codigo Inicial e Final
    If CodRequisitanteDe.Text <> "" Then
        sCodRequisitante_I = CStr(CodRequisitanteDe.Text)
    Else
        sCodRequisitante_I = ""
    End If
    
    If CodRequisitanteAte.Text <> "" Then
        sCodRequisitante_F = CStr(CodRequisitanteAte.Text)
    Else
        sCodRequisitante_F = ""
    End If
            
    If sCodRequisitante_I <> "" And sCodRequisitante_F <> "" Then
        
        If StrParaLong(sCodRequisitante_I) > StrParaLong(sCodRequisitante_F) Then gError 73993
        
    End If
    
    If NomeReqDe.Text <> "" Then
        sNomeReq_I = NomeReqDe.Text
    Else
        sNomeReq_I = ""
    End If
    
    If NomeReqAte.Text <> "" Then
        sNomeReq_F = NomeReqAte.Text
    Else
        sNomeReq_F = ""
    End If
    
    If sNomeReq_I <> "" And sNomeReq_F <> "" Then
        If sNomeReq_I > sNomeReq_F Then gError 73994
    End If
    
    'critica Ccl Inicial e Final
    If CclDe.ClipText <> "" Then
        lErro = CF("Ccl_Formata", CclDe.Text, sCclFormata, iCclPreenchida)
        If lErro <> SUCESSO Then gError 73995
        
        sCcl_I = sCclFormata
    Else
        sCcl_I = ""
    End If
    
    If CclAte.ClipText <> "" Then
        lErro = CF("Ccl_Formata", CclAte.Text, sCclFormata, iCclPreenchida)
        If lErro <> SUCESSO Then gError 73996
        
        sCcl_F = sCclFormata
    Else
        sCcl_F = ""
    End If
            
    If sCcl_I <> "" And sCcl_F <> "" Then
        
        If sCcl_I > sCcl_F Then gError 73997
        
    End If
    
    'critica CodigoFilial Inicial e Final
    If CodigoFilialDe.Text <> "" Then
        sCodFilial_I = CStr(CodigoFilialDe.Text)
    Else
        sCodFilial_I = ""
    End If

    If CodigoFilialAte.Text <> "" Then
        sCodFilial_F = CStr(CodigoFilialAte.Text)
    Else
        sCodFilial_F = ""
    End If

    If sCodFilial_I <> "" And sCodFilial_F <> "" Then

        If StrParaInt(sCodFilial_I) > StrParaInt(sCodFilial_F) Then gError 73998

    End If

    'critica Nome da Filial inicial e final
    If NomeFilialDe.Text <> "" Then
        sNomeFilial_I = NomeFilialDe.Text
    Else
        sNomeFilial_I = ""
    End If

    If NomeFilialAte.Text <> "" Then
        sNomeFilial_F = NomeFilialAte.Text
    Else
        sNomeFilial_F = ""
    End If

    If sNomeFilial_I <> "" And sNomeFilial_F <> "" Then
        If sNomeFilial_I > sNomeFilial_F Then gError 73999
    End If
    
    'critica Codigo Inicial e Final
    If CodRequisicaoDe.Text <> "" Then
        sCodRequisicao_I = CStr(CodRequisicaoDe.Text)
    Else
        sCodRequisicao_I = ""
    End If
    
    If CodRequisicaoAte.Text <> "" Then
        sCodRequisicao_F = CStr(CodRequisicaoAte.Text)
    Else
        sCodRequisicao_F = ""
    End If
            
    If sCodRequisicao_I <> "" And sCodRequisicao_F <> "" Then
        
        If StrParaLong(sCodRequisicao_I) > StrParaLong(sCodRequisicao_F) Then gError 74000
        
    End If
    
    'data de Envio inicial não pode ser maior que a final
    If Trim(DataEnvioDe.ClipText) <> "" And Trim(DataEnvioAte.ClipText) <> "" Then
    
         If CDate(DataEnvioDe.Text) > CDate(DataEnvioAte.Text) Then gError 74001
    
    End If
    
    
    'data Limite inicial não pode ser maior que a data limite final
    If Trim(DataLimiteDe.ClipText) <> "" And Trim(DataLimiteAte.ClipText) <> "" Then
    
         If CDate(DataLimiteDe.Text) > CDate(DataLimiteAte.Text) Then gError 74002
    
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                
        Case 73993
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_INICIAL_MAIOR", gErr)
            CodRequisitanteDe.SetFocus
                
        Case 73994
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_INICIAL_MAIOR", gErr)
            NomeReqDe.SetFocus
            
        Case 73995, 73996
        
        Case 73997
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_INICIAL_MAIOR", gErr)
            CclDe.SetFocus
                
        Case 73998
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            CodigoFilialDe.SetFocus
            
        Case 73999
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            NomeFilialDe.SetFocus
            
        Case 74000
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISICAO_INICIAL_MAIOR", gErr)
            CodRequisicaoDe.SetFocus
            
        Case 74001
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAENVIO_INICIAL_MAIOR", gErr)
            DataEnvioDe.SetFocus
            
        Case 74002
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATALIMITE_INICIAL_MAIOR", gErr)
            DataLimiteDe.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172931)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCodigo_I As String, sCodigo_F As String, sNome_I As String, sNome_F As String, sCcl_I As String, sCcl_F As String, sCodFilial_I As String, sCodFilial_F As String, sNomeFilial_I As String, sNomeFilial_F As String, sCodRequisicao_I As String, sCodRequisicao_F As String, sOrdenacaoPor As String, sOrd As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

   If sCodigo_I <> "" Then sExpressao = "ReqCod >= " & Forprint_ConvLong(StrParaLong(sCodigo_I))

   If sCodigo_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "ReqCod <= " & Forprint_ConvLong(StrParaLong(sCodigo_F))

    End If

   If sNome_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "ReqNomeInic"

    End If
    
    If sNome_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "ReqNomeFim"

    End If
   
'    If sCcl_I <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Ccl >= " & Forprint_ConvTexto((sCcl_I))
'
'    End If
'
'    If sCcl_F <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Ccl <= " & Forprint_ConvTexto((sCcl_F))
'
'    End If
   
    If sCodFilial_I <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilEmpCodInic"

    End If
    
    If sCodFilial_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilEmpCodFim"

    End If

    If sNomeFilial_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilEmpNomeInic"

    End If
    
    If sNomeFilial_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilEmpNomeFim"

    End If
   
    If sCodRequisicao_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Req >= " & Forprint_ConvLong(StrParaLong(sCodRequisicao_I))

    End If

    If sCodRequisicao_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Req <= " & Forprint_ConvLong(StrParaLong(sCodRequisicao_F))

    End If

   If Trim(DataEnvioDe.ClipText) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Envio >= " & Forprint_ConvData(CDate(DataEnvioDe.Text))
        
    End If
    
    If Trim(DataEnvioAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Envio <= " & Forprint_ConvData(CDate(DataEnvioAte.Text))

    End If
        
    If Trim(DataLimiteDe.ClipText) <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Limite >= " & Forprint_ConvData(CDate(DataLimiteDe.Text))

    End If
    
    If Trim(DataLimiteAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Limite <= " & Forprint_ConvData(CDate(DataLimiteAte.Text))

    End If
    
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172932)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim sOrdenacaoPor As String
Dim sCclMascarado As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 74003
   
    'pega Codigo inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODREQINIC", sParam)
    If lErro <> SUCESSO Then gError 74004
    
    CodRequisitanteDe.Text = sParam
    Call CodRequisitanteDe_Validate(bSGECancelDummy)
    
    'pega  Codigo final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODREQFIM", sParam)
    If lErro <> SUCESSO Then gError 74005
    
    CodRequisitanteAte.Text = sParam
    Call CodRequisitanteAte_Validate(bSGECancelDummy)
                
    'pega  Nome Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEREQINIC", sParam)
    If lErro <> SUCESSO Then gError 74006
                   
    NomeReqDe.Text = sParam
    Call NomeReqDe_Validate(bSGECancelDummy)
    
    'pega  Nome Final e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEREQFIM", sParam)
    If lErro <> SUCESSO Then gError 74007
                   
    NomeReqAte.Text = sParam
    Call NomeReqAte_Validate(bSGECancelDummy)
                        
    'pega  Ccl Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TCCLINIC", sParam)
    If lErro <> SUCESSO Then gError 74008
                   
    If Len(Trim(sParam)) > 0 Then
        lErro = Mascara_MascararCcl(sParam, sCclMascarado)
        If lErro <> SUCESSO Then gError 74009
        CclDe.PromptInclude = False
        CclDe.Text = sCclMascarado
        CclDe.PromptInclude = True
        
    End If
    Call CclDe_Validate(bSGECancelDummy)
                          
                          
    'pega  Ccl Final e exibe
    lErro = objRelOpcoes.ObterParametro("TCCLFIM", sParam)
    If lErro <> SUCESSO Then gError 74010
                   
    If Len(Trim(sParam)) > 0 Then
    
        lErro = Mascara_MascararCcl(sParam, sCclMascarado)
        If lErro <> SUCESSO Then gError 74011
        
        CclAte.PromptInclude = False
        CclAte.Text = sCclMascarado
        CclAte.PromptInclude = True
        
    End If
    Call CclAte_Validate(bSGECancelDummy)
                              
    'pega Codigo inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILIALINIC", sParam)
    If lErro <> SUCESSO Then gError 74012
    
    CodigoFilialDe.Text = sParam
    Call CodigoFilialDe_Validate(bSGECancelDummy)
    
    'pega  Codigo final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILIALFIM", sParam)
    If lErro <> SUCESSO Then gError 74013
    
    CodigoFilialAte.Text = sParam
    Call CodigoFilialAte_Validate(bSGECancelDummy)
                
    'pega  Nome Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFILIALINIC", sParam)
    If lErro <> SUCESSO Then gError 74014
                   
    NomeFilialDe.Text = sParam
    Call NomeFilialDe_Validate(bSGECancelDummy)
    
    'pega  Nome Final e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFILIALFIM", sParam)
    If lErro <> SUCESSO Then gError 74015
                   
    NomeFilialAte.Text = sParam
    Call NomeFilialAte_Validate(bSGECancelDummy)
                        
    'pega Codigo inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODREQUISICAOINIC", sParam)
    If lErro <> SUCESSO Then gError 74016
    
    CodRequisicaoDe.Text = sParam
    
    'pega  Codigo final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODREQUISICAOFIM", sParam)
    If lErro <> SUCESSO Then gError 74017
    
    CodRequisicaoAte.Text = sParam
                                   
    'pega DataEnvio inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DENVINIC", sParam)
    If lErro <> SUCESSO Then gError 74018
    
    Call DateParaMasked(DataEnvioDe, CDate(sParam))
    
    'pega data de envio final e exibe
    lErro = objRelOpcoes.ObterParametro("DENVFIM", sParam)
    If lErro <> SUCESSO Then gError 74019

    Call DateParaMasked(DataEnvioAte, CDate(sParam))

    'pega data limite inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DLIMINIC", sParam)
    If lErro <> SUCESSO Then gError 74020

    Call DateParaMasked(DataLimiteDe, CDate(sParam))
       
    'pega data limite final e exibe
    lErro = objRelOpcoes.ObterParametro("DLIMFIM", sParam)
    If lErro <> SUCESSO Then gError 74021

    Call DateParaMasked(DataLimiteAte, CDate(sParam))
              
    lErro = objRelOpcoes.ObterParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> SUCESSO Then gError 74022
    
    Select Case sOrdenacaoPor
        
            Case "CodReq"
            
                ComboOrdenacao.ListIndex = ORD_POR_CODIGO
            
            Case "DataEnvio"
                ComboOrdenacao.ListIndex = ORD_POR_ENVIO
                
            Case "DataLimite"
                ComboOrdenacao.ListIndex = ORD_POR_LIMITE
                
            Case Else
                gError 74023
                  
    End Select
        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 74003 To 74023
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172933)

    End Select

    Exit Function

End Function
Private Sub CodigoFilialDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_CodigoFilialDe_Validate

    If Len(Trim(CodigoFilialDe.Text)) > 0 Then

        objFilialEmpresa.iCodFilial = StrParaInt(CodigoFilialDe.Text)
        'Lê o código informado
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 74024
        
        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 74025

    End If

    Exit Sub

Erro_CodigoFilialDe_Validate:

    Cancel = True


    Select Case gErr

        Case 74024

        Case 74025
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172934)

    End Select

    Exit Sub

End Sub
Private Sub CodigoFilialAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_CodigoFilialAte_Validate

    If Len(Trim(CodigoFilialAte.Text)) > 0 Then

        objFilialEmpresa.iCodFilial = StrParaInt(CodigoFilialAte.Text)
        'Lê o código informado
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 74026
        
        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 74027

    End If

    Exit Sub

Erro_CodigoFilialAte_Validate:

    Cancel = True


    Select Case gErr

        Case 74026

        Case 74027
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172935)

    End Select

    Exit Sub

End Sub

Private Sub CodRequisitanteDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objRequisitante As New ClassRequisitante

On Error GoTo Erro_CodRequisitanteDe_Validate

    If Len(Trim(CodRequisitanteDe.Text)) > 0 Then

        objRequisitante.lCodigo = StrParaLong(CodRequisitanteDe.Text)
        'Lê o código informado
        lErro = CF("Requisitante_Le", objRequisitante)
        If lErro <> SUCESSO And lErro <> 49084 Then gError 74028

        'Se não encontrou o Requisitante ==> erro
        If lErro = 49084 Then gError 74029
        
    End If

    Exit Sub

Erro_CodRequisitanteDe_Validate:

    Cancel = True


    Select Case gErr

        Case 74028

        Case 74029
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_INEXISTENTE", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172936)

    End Select

    Exit Sub
    
End Sub


Private Sub CodRequisitanteAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objRequisitante As New ClassRequisitante

On Error GoTo Erro_CodRequisitanteAte_Validate

    If Len(Trim(CodRequisitanteAte.Text)) > 0 Then

        objRequisitante.lCodigo = StrParaLong(CodRequisitanteAte.Text)
        'Lê o código informado
        lErro = CF("Requisitante_Le", objRequisitante)
        If lErro <> SUCESSO And lErro <> 49084 Then gError 74030

        'Se não encontrou o Requisitante ==> erro
        If lErro = 49084 Then gError 74031
        
    End If

    Exit Sub

Erro_CodRequisitanteAte_Validate:

    Cancel = True


    Select Case gErr

        Case 74030

        Case 74031
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_INEXISTENTE", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172937)

    End Select

Exit Sub

End Sub

Private Sub NomeFilialDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim bAchou As Boolean
Dim colFiliais As New Collection

On Error GoTo Erro_NomeFilialDe_Validate

    bAchou = False
    
    If Len(Trim(NomeFilialDe.Text)) > 0 Then

        lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, colFiliais)
        If lErro <> SUCESSO Then gError 74032

        'Carrega a Filial com o Nome informado
        For Each objFilialEmpresa In colFiliais
            If objFilialEmpresa.sNome = NomeFilialDe.Text Then
                bAchou = True
                Exit For
            End If
        Next

        'Se não encontrou Filial com o Nome informado ==> erro
        If bAchou = False Then gError 74033
        
        NomeFilialDe.Text = objFilialEmpresa.sNome

    End If

    Exit Sub

Erro_NomeFilialDe_Validate:

    Cancel = True

    Select Case gErr

        Case 74032

        Case 74033
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, NomeFilialDe.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172938)

    End Select

Exit Sub

End Sub

Private Sub NomeFilialAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim bAchou As Boolean
Dim colFiliais As New Collection

On Error GoTo Erro_NomeFilialAte_Validate

    bAchou = False
    If Len(Trim(NomeFilialAte.Text)) > 0 Then

        lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, colFiliais)
        If lErro <> SUCESSO Then gError 73034

        'Carrega a Filial com o Nome informado
        For Each objFilialEmpresa In colFiliais
            If objFilialEmpresa.sNome = NomeFilialAte.Text Then
                bAchou = True
                Exit For
            End If
        Next

        'Se não encontrou Filial com o Nome informado ==> erro
        If bAchou = False Then gError 73035

        NomeFilialAte.Text = objFilialEmpresa.sNome

    End If

    Exit Sub

Erro_NomeFilialAte_Validate:

    Cancel = True


    Select Case gErr

        Case 73034

        Case 73035
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, NomeFilialAte.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172939)

    End Select

Exit Sub

End Sub

Private Sub NomeReqDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objRequisitante As New ClassRequisitante

On Error GoTo Erro_NomeReqDe_Validate

    If Len(Trim(NomeReqDe.Text)) > 0 Then

        objRequisitante.sNomeReduzido = NomeReqDe.Text
        'Lê o Requisitante informado
        lErro = CF("Requisitante_Le_NomeReduzido", objRequisitante)
        If lErro <> SUCESSO And lErro <> 51152 Then gError 73036

        'Se não encontrou o Requisitante ==> erro
        If lErro = 51152 Then gError 73037
        
        NomeReqDe.Text = objRequisitante.sNomeReduzido
        
    End If

    Exit Sub

Erro_NomeReqDe_Validate:

    Cancel = True


    Select Case gErr

        Case 73036

        Case 73037
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_INEXISTENTE", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172940)

    End Select

Exit Sub

End Sub

Private Sub NomeReqAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objRequisitante As New ClassRequisitante

On Error GoTo Erro_NomeReqAte_Validate

    If Len(Trim(NomeReqAte.Text)) > 0 Then

        objRequisitante.sNomeReduzido = NomeReqAte.Text
        'Lê o Requisitante informado
        lErro = CF("Requisitante_Le_NomeReduzido", objRequisitante)
        If lErro <> SUCESSO And lErro <> 51152 Then gError 73038

        'Se não encontrou o Requisitante ==> erro
        If lErro = 51152 Then gError 73039
        
        NomeReqAte.Text = objRequisitante.sNomeReduzido
        
    End If

    Exit Sub

Erro_NomeReqAte_Validate:

    Cancel = True


    Select Case gErr

        Case 73038

        Case 73039
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_INEXISTENTE", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172941)

    End Select

Exit Sub

End Sub

Private Sub CclDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclFormata As String
Dim iCclPreenchida As Integer

On Error GoTo Erro_CclDe_Validate

    If Len(Trim(CclDe.ClipText)) > 0 Then

        'Coloca Ccl no formato do BD
        lErro = CF("Ccl_Formata", CclDe.Text, sCclFormata, iCclPreenchida)
        If lErro <> SUCESSO Then gError 73040
        
        objCcl.sCcl = sCclFormata
        
        'Lê o Ccl informado
        lErro = CF("Ccl_Le", objCcl)
        If lErro <> SUCESSO And lErro <> 5599 Then gError 73041

        'Se não encontrou o Ccl ==> erro
        If lErro = 5599 Then gError 73042
            
    End If

    Exit Sub

Erro_CclDe_Validate:

    Cancel = True

    Select Case gErr

        Case 73040, 73041

        Case 73042
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_INEXISTENTE", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172942)

    End Select

Exit Sub

End Sub

Private Sub CclAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclFormata As String
Dim iCclPreenchida As Integer

On Error GoTo Erro_CclAte_Validate

    If Len(Trim(CclAte.ClipText)) > 0 Then

        'Coloca Ccl no formato do BD
        lErro = CF("Ccl_Formata", CclAte.Text, sCclFormata, iCclPreenchida)
        If lErro <> SUCESSO Then gError 73043
        
        objCcl.sCcl = sCclFormata
        
        'Lê o Ccl informado
        lErro = CF("Ccl_Le", objCcl)
        If lErro <> SUCESSO And lErro <> 5599 Then gError 73044

        'Se não encontrou o Ccl ==> erro
        If lErro = 5599 Then gError 73045
        
    End If

    Exit Sub

Erro_CclAte_Validate:

    Cancel = True

    Select Case gErr

        Case 73043, 73044

        Case 73045
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_INEXISTENTE", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172943)

    End Select

Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

''    Parent.HelpContextID = IDH_RELOP_REQ
    Set Form_Load_Ocx = Me
    Caption = "Requisições x Notas Fiscais"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpRequisicoesNF"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Controls
Public Property Get Controls() As Object
    Set Controls = UserControl.Controls
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Height() As Long
    Height = UserControl.Height
End Property

Public Property Get Width() As Long
    Width = UserControl.Width
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ActiveControl
Public Property Get ActiveControl() As Object
    Set ActiveControl = UserControl.ActiveControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Sub Unload(objme As Object)
    
   RaiseEvent Unload
    
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

'***** fim do trecho a ser copiado ******
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is CodRequisitanteDe Then
            Call LabelCodRequisitanteDe_Click
            
        ElseIf Me.ActiveControl Is CodRequisitanteAte Then
            Call LabelCodRequisitanteAte_Click
            
        ElseIf Me.ActiveControl Is NomeReqDe Then
            Call LabelNomeReqDe_Click
            
        ElseIf Me.ActiveControl Is NomeReqAte Then
            Call LabelNomeReqAte_Click
            
        ElseIf Me.ActiveControl Is CclDe Then
            Call LabelCclDe_Click
            
        ElseIf Me.ActiveControl Is CclAte Then
            Call LabelCclAte_Click
            
        ElseIf Me.ActiveControl Is CodigoFilialDe Then
            Call LabelCodigoDe_Click
        
        ElseIf Me.ActiveControl Is CodigoFilialAte Then
            Call LabelCodigoAte_Click
        
        ElseIf Me.ActiveControl Is NomeFilialDe Then
            Call LabelNomeDe_Click
        
        ElseIf Me.ActiveControl Is NomeFilialAte Then
            Call LabelNomeAte_Click
        
        ElseIf Me.ActiveControl Is CodRequisicaoDe Then
            Call LabelCodRequisicaoDe_Click
        
        ElseIf Me.ActiveControl Is CodRequisicaoAte Then
            Call LabelCodRequisicaoAte_Click
        
        End If
    
    End If

End Sub


Private Sub LabelCodigoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoDe, Source, X, Y)
End Sub

Private Sub LabelCodigoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoAte, Source, X, Y)
End Sub

Private Sub LabelCodigoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoAte, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub






Private Sub LabelNomeDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeDe, Source, X, Y)
End Sub

Private Sub LabelNomeDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeDe, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeAte, Source, X, Y)
End Sub

Private Sub LabelNomeAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodRequisitanteAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodRequisitanteAte, Source, X, Y)
End Sub

Private Sub LabelCodRequisitanteAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodRequisitanteAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodRequisitanteDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodRequisitanteDe, Source, X, Y)
End Sub

Private Sub LabelCodRequisitanteDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodRequisitanteDe, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeReqDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeReqDe, Source, X, Y)
End Sub

Private Sub LabelNomeReqDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeReqDe, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeReqAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeReqAte, Source, X, Y)
End Sub

Private Sub LabelNomeReqAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeReqAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCclAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCClAte, Source, X, Y)
End Sub

Private Sub LabelCclAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCClAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCclDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCclDE, Source, X, Y)
End Sub

Private Sub LabelCclDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCclDE, Button, Shift, X, Y)
End Sub

Private Sub LabelCodRequisicaoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodRequisicaoAte, Source, X, Y)
End Sub

Private Sub LabelCodRequisicaoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodRequisicaoAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodRequisicaoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodRequisicaoDe, Source, X, Y)
End Sub

Private Sub LabelCodRequisicaoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodRequisicaoDe, Button, Shift, X, Y)
End Sub

Private Sub LabelDataEnvioDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataEnvioDe, Source, X, Y)
End Sub

Private Sub LabelDataEnvioDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataEnvioDe, Button, Shift, X, Y)
End Sub

Private Sub LabelDataLimiteAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataLimiteAte, Source, X, Y)
End Sub

Private Sub LabelDataLimiteAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataLimiteAte, Button, Shift, X, Y)
End Sub

Private Sub LabelDataLimiteDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataLimiteDe, Source, X, Y)
End Sub

Private Sub LabelDataLimiteDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataLimiteDe, Button, Shift, X, Y)
End Sub

Private Sub LabelDataEnvioAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataEnvioAte, Source, X, Y)
End Sub

Private Sub LabelDataEnvioAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataEnvioAte, Button, Shift, X, Y)
End Sub

