VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpReqRequisicoesOcx 
   ClientHeight    =   5715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8025
   ScaleHeight     =   5715
   ScaleWidth      =   8025
   Begin VB.Frame Frame1 
      Caption         =   "Requisitantes"
      Height          =   4065
      Index           =   2
      Left            =   420
      TabIndex        =   45
      Top             =   1350
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Frame FrameCcl 
         Caption         =   "Centro de Custo"
         Height          =   705
         Left            =   330
         TabIndex        =   52
         Top             =   2925
         Visible         =   0   'False
         Width           =   3210
         Begin MSMask.MaskEdBox CclDe 
            Height          =   315
            Left            =   540
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
         Begin MSMask.MaskEdBox CclAte 
            Height          =   315
            Left            =   2145
            TabIndex        =   16
            Top             =   240
            Visible         =   0   'False
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            PromptChar      =   " "
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
            Left            =   120
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   54
            Top             =   300
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label LabelCclAte 
            AutoSize        =   -1  'True
            Caption         =   "At�:"
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
            TabIndex        =   53
            Top             =   300
            Visible         =   0   'False
            Width           =   360
         End
      End
      Begin VB.Frame FrameNome 
         Caption         =   "Nome"
         Height          =   675
         Left            =   315
         TabIndex        =   49
         Top             =   1605
         Width           =   5160
         Begin MSMask.MaskEdBox NomeReqDe 
            Height          =   300
            Left            =   525
            TabIndex        =   17
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
            TabIndex        =   18
            Top             =   240
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin VB.Label LabelNomeReqAte 
            AutoSize        =   -1  'True
            Caption         =   "At�:"
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
            TabIndex        =   51
            Top             =   315
            Width           =   360
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
      End
      Begin VB.Frame FrameCodigo 
         Caption         =   "C�digo"
         Height          =   705
         Left            =   315
         TabIndex        =   46
         Top             =   450
         Width           =   3210
         Begin MSMask.MaskEdBox CodRequisitanteDe 
            Height          =   300
            Left            =   525
            TabIndex        =   13
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
            TabIndex        =   14
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
            TabIndex        =   48
            Top             =   300
            Width           =   315
         End
         Begin VB.Label LabelCodRequisitanteAte 
            AutoSize        =   -1  'True
            Caption         =   "At�:"
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
            TabIndex        =   47
            Top             =   300
            Width           =   360
         End
      End
   End
   Begin VB.ComboBox ComboOrdenacao 
      Height          =   315
      ItemData        =   "RelOpReqRequisicoesOcx.ctx":0000
      Left            =   1590
      List            =   "RelOpReqRequisicoesOcx.ctx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   465
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5730
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpReqRequisicoesOcx.ctx":0043
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpReqRequisicoesOcx.ctx":019D
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpReqRequisicoesOcx.ctx":0327
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpReqRequisicoesOcx.ctx":0859
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
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
      Left            =   3840
      Picture         =   "RelOpReqRequisicoesOcx.ctx":09D7
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   135
      Width           =   1635
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpReqRequisicoesOcx.ctx":0AD9
      Left            =   855
      List            =   "RelOpReqRequisicoesOcx.ctx":0ADB
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   90
      Width           =   2805
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4200
      Index           =   1
      Left            =   330
      TabIndex        =   25
      Top             =   1305
      Width           =   6495
      Begin VB.Frame Frame2 
         Caption         =   "Filial Empresa"
         Height          =   1065
         Left            =   90
         TabIndex        =   40
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
         Begin VB.Label LabelNomeAte 
            AutoSize        =   -1  'True
            Caption         =   "Nome At�:"
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
            TabIndex        =   44
            Top             =   675
            Width           =   900
         End
         Begin VB.Label LabelCodigoAte 
            AutoSize        =   -1  'True
            Caption         =   "C�digo At�:"
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
            TabIndex        =   43
            Top             =   315
            Width           =   1005
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
            TabIndex        =   42
            Top             =   705
            Width           =   855
         End
         Begin VB.Label LabelCodigoDe 
            AutoSize        =   -1  'True
            Caption         =   "C�digo De:"
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
            TabIndex        =   41
            Top             =   323
            Width           =   960
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Requisi��es"
         Height          =   2985
         Left            =   90
         TabIndex        =   26
         Top             =   1125
         Width           =   6255
         Begin VB.CheckBox CheckRequisicoes 
            Caption         =   "Inclui Requisi��es Atendidas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   270
            TabIndex        =   12
            Top             =   2565
            Width           =   4080
         End
         Begin VB.Frame Frame8 
            Caption         =   "Data Envio"
            Height          =   690
            Left            =   180
            TabIndex        =   35
            Top             =   270
            Width           =   5775
            Begin MSComCtl2.UpDown UpDownDataEnvioAte 
               Height          =   315
               Left            =   4860
               TabIndex        =   36
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
               TabIndex        =   37
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
            Begin VB.Label LabelDataEnvioAte 
               AutoSize        =   -1  'True
               Caption         =   "At�:"
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
               TabIndex        =   39
               Top             =   315
               Width           =   360
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
               TabIndex        =   38
               Top             =   315
               Width           =   315
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Data Limite"
            Height          =   690
            Left            =   180
            TabIndex        =   30
            Top             =   1035
            Width           =   5775
            Begin MSComCtl2.UpDown UpDownDataLimiteAte 
               Height          =   315
               Left            =   4860
               TabIndex        =   31
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
               TabIndex        =   32
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
               TabIndex        =   34
               Top             =   315
               Width           =   315
            End
            Begin VB.Label LabelDataLimiteAte 
               AutoSize        =   -1  'True
               Caption         =   "At�:"
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
               TabIndex        =   33
               Top             =   315
               Width           =   360
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "C�digo"
            Height          =   690
            Left            =   180
            TabIndex        =   27
            Top             =   1800
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
               TabIndex        =   29
               Top             =   330
               Width           =   315
            End
            Begin VB.Label LabelCodRequisicaoAte 
               AutoSize        =   -1  'True
               Caption         =   "At�:"
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
               TabIndex        =   28
               Top             =   315
               Width           =   360
            End
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4605
      Left            =   240
      TabIndex        =   55
      Top             =   990
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   8123
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Requisi��es"
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
      Left            =   225
      TabIndex        =   57
      Top             =   510
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Op��o:"
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
      Left            =   210
      TabIndex        =   56
      Top             =   135
      Width           =   615
   End
End
Attribute VB_Name = "RelOpReqRequisicoesOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'RelOpRequisitantes
Const ORD_POR_CODIGO = 0
Const ORD_POR_NOME = 1
Const ORD_POR_CCL = 2
Const ORD_POR_LIMITE = 3
Const ORD_POR_ENVIO = 4


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

    If Not (gobjRelatorio Is Nothing) Then gError 68699
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 68700

    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 68700
        
        Case 68699
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172852)

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
    If lErro <> SUCESSO Then gError 68701
    
    ComboOrdenacao.ListIndex = 0
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    CheckRequisicoes.Value = vbUnchecked
    
    Exit Sub
    
Erro_Limpa_Tela_Rel:
    
    Select Case gErr
    
        Case 68701
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172853)

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
    If lErro <> SUCESSO Then gError 68703

    CclDe.Mask = sMascaraCcl
    CclAte.Mask = sMascaraCcl
    
    iFrameAtual = 1
    
    ComboOrdenacao.ListIndex = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 68703
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172854)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172855)

    End Select

    Exit Sub

End Sub

Private Sub DataEnvioDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEnvioDe_Validate

    'Verifica se a DataDe est� preenchida
    If Len(Trim(DataEnvioDe.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataEnvioDe.Text)
    If lErro <> SUCESSO Then gError 68794

    Exit Sub
                   
Erro_DataEnvioDe_Validate:

    Cancel = True

    Select Case gErr

        Case 68794
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172856)

    End Select

    Exit Sub

End Sub

Private Sub DataEnvioAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEnvioAte_Validate

    'Verifica se a DataDe est� preenchida
    If Len(Trim(DataEnvioAte.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataEnvioAte.Text)
    If lErro <> SUCESSO Then gError 68795

    Exit Sub
                   
Erro_DataEnvioAte_Validate:

    Cancel = True

    Select Case gErr

        Case 68795
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172857)

    End Select

    Exit Sub

End Sub

Private Sub DataLimiteAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataLimiteAte_Validate

    'Verifica se a DataDe est� preenchida
    If Len(Trim(DataLimiteAte.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataLimiteAte.Text)
    If lErro <> SUCESSO Then gError 68796

    Exit Sub
                   
Erro_DataLimiteAte_Validate:

    Cancel = True

    Select Case gErr

        Case 68796
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172858)

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
    If lErro <> SUCESSO Then gError 68799

    Exit Sub

Erro_UpDownDataEnvioDe_DownClick:

    Select Case gErr

        Case 68799
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 172859)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEnvioDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvioDe_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEnvioDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 68798

    Exit Sub

Erro_UpDownDataEnvioDe_UpClick:

    Select Case gErr

        Case 68798
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 172860)

    End Select

    Exit Sub

End Sub
Private Sub UpDownDataEnvioAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvioAte_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEnvioAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 68801

    Exit Sub

Erro_UpDownDataEnvioAte_DownClick:

    Select Case gErr

        Case 68801
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 172861)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEnvioAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvioAte_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEnvioAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 68800

    Exit Sub

Erro_UpDownDataEnvioAte_UpClick:

    Select Case gErr

        Case 68800
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 172862)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataLimiteAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataLimiteAte_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataLimiteAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 68803

    Exit Sub

Erro_UpDownDataLimiteAte_DownClick:

    Select Case gErr

        Case 68803
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 172863)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataLimiteAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataLimiteAte_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataLimiteAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 68802

    Exit Sub

Erro_UpDownDataLimiteAte_UpClick:

    Select Case gErr

        Case 68802
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 172864)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataLimiteDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataLimiteDe_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataLimiteDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 68805

    Exit Sub

Erro_UpDownDataLimiteDe_DownClick:

    Select Case gErr

        Case 68805
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 172865)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataLimiteDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataLimiteDe_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataLimiteDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 68804

    Exit Sub

Erro_UpDownDataLimiteDe_UpClick:

    Select Case gErr

        Case 68804
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 172866)

    End Select

    Exit Sub

End Sub

Private Sub DataLimiteDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataLimiteDe_Validate

    'Verifica se a DataDe est� preenchida
    If Len(Trim(DataLimiteDe.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataLimiteDe.Text)
    If lErro <> SUCESSO Then gError 68797

    Exit Sub
                   
Erro_DataLimiteDe_Validate:

    Cancel = True

    Select Case gErr

        Case 68797
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172867)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172868)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172869)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172870)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172871)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172872)

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
        If lErro <> SUCESSO Then gError 68552
        
        'Preenche com o Ccl
        objCcl.sCcl = sCclFormata
        
    End If

    'Chama Tela Cclista
    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCclAte)

   Exit Sub

Erro_LabelCclAte_Click:

    Select Case gErr

        Case 68552
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172873)

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
        If lErro <> SUCESSO Then gError 68553

        'Preenche com o Ccl
        objCcl.sCcl = sCclFormata
        
    End If

    'Chama Tela Cclista
    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCclDe)

   Exit Sub

Erro_LabelCclDe_Click:

    Select Case gErr

        Case 68553
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172874)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172875)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172876)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172877)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172878)

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
    If lErro <> SUCESSO Then gError 68704

    CclDe.PromptInclude = False
    CclDe.Text = sCclMascarado
    CclDe.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoCclDe_evSelecao:

    Select Case gErr

        Case 68704

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172879)

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
    If lErro <> SUCESSO Then gError 68705

    CclAte.PromptInclude = False
    CclAte.Text = sCclMascarado
    CclAte.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoCclAte_evSelecao:

    Select Case gErr

        Case 68705

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172880)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a op��o de relat�rio com os par�metros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da op��o de relat�rio n�o pode ser vazia
    If ComboOpcoes.Text = "" Then gError 68706

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 68707

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 68708
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 68709
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 68706
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 68707 To 68709
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172881)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 68710

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 68711

        'retira nome das op��es do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call Limpa_Tela_Rel
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 68710
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 68711

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172882)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 68712

    Select Case ComboOrdenacao.ListIndex

            Case ORD_POR_CODIGO
                
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "RequisitanteCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "RequisicaoCod", 1)
                
            Case ORD_POR_NOME

                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaNome", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "RequisitanteNome", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "RequisicaoCod", 1)
                
            Case ORD_POR_CCL

                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "Ccl", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "RequisitanteNome", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "RequisicaoCod", 1)
                
            Case ORD_POR_LIMITE

                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "RequisitanteCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "DataLimite", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "RequisicaoCod", 1)
                
            Case ORD_POR_ENVIO

                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "RequisitanteNome", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "DataEnvio", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "RequisicaoCod", 1)
                
            Case Else
                gError 74970

    End Select

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 68712, 74970

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172883)

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
    If lErro <> SUCESSO Then gError 68713

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 68714
         
    lErro = objRelOpcoes.IncluirParametro("NCODREQINIC", sCodRequisitante_I)
    If lErro <> AD_BOOL_TRUE Then gError 68715
         
    lErro = objRelOpcoes.IncluirParametro("TNOMEREQINIC", NomeReqDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 68716
    
    lErro = objRelOpcoes.IncluirParametro("TCCLINIC", sCcl_I)
    If lErro <> AD_BOOL_TRUE Then gError 68717
    
    lErro = objRelOpcoes.IncluirParametro("NCODFILIALINIC", sCodFilial_I)
    If lErro <> AD_BOOL_TRUE Then gError 68718
         
    lErro = objRelOpcoes.IncluirParametro("TNOMEFILIALINIC", NomeFilialDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 68719
    
    lErro = objRelOpcoes.IncluirParametro("NCODREQUISICAOINIC", sCodRequisicao_I)
    If lErro <> AD_BOOL_TRUE Then gError 68720
         
    lErro = objRelOpcoes.IncluirParametro("TNOMEREQUISICAOINIC", NomeReqDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 68721
    
    'Preenche dataenvio inicial
    If Trim(DataEnvioDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DENVINIC", DataEnvioDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DENVINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 68722
    
    'Preenche a data limite inicial
    If Trim(DataLimiteDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DLIMINIC", DataLimiteDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DLIMINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 68723
    
    lErro = objRelOpcoes.IncluirParametro("NCODREQFIM", sCodRequisitante_F)
    If lErro <> AD_BOOL_TRUE Then gError 68726
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEREQFIM", NomeReqAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 68727
        
    lErro = objRelOpcoes.IncluirParametro("TCCLFIM", sCcl_F)
    If lErro <> AD_BOOL_TRUE Then gError 68728
    
    lErro = objRelOpcoes.IncluirParametro("NCODFILIALFIM", sCodFilial_F)
    If lErro <> AD_BOOL_TRUE Then gError 68729
         
    lErro = objRelOpcoes.IncluirParametro("TNOMEFILIALFIM", NomeFilialAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 68730
    
    lErro = objRelOpcoes.IncluirParametro("NCODREQUISICAOFIM", sCodRequisicao_F)
    If lErro <> AD_BOOL_TRUE Then gError 68731
         
    lErro = objRelOpcoes.IncluirParametro("TNOMEREQUISICAOFIM", NomeReqAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 68732
    
    'Preenche data de envio Final
    If Trim(DataEnvioAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DENVFIM", DataEnvioAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DENVFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 68724
    
    'Preenche data limite final
    If Trim(DataLimiteAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DLIMFIM", DataLimiteAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DLIMFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 68725
    
    'Exibe Requisicoes Atendidas
    If CheckRequisicoes.Value Then
        sCheck = vbChecked
    Else
        sCheck = vbUnchecked
    End If

    lErro = objRelOpcoes.IncluirParametro("NREQATEND", sCheck)
    If lErro <> AD_BOOL_TRUE Then gError 72516

    Select Case ComboOrdenacao.ListIndex
        
            Case ORD_POR_CODIGO
            
                sOrdenacaoPor = "CodReq"
                
            Case ORD_POR_NOME
                
                sOrdenacaoPor = "NomeRequisicao"
                
            Case ORD_POR_CCL
                sOrdenacaoPor = "Ccl"
                
            Case ORD_POR_ENVIO
                sOrdenacaoPor = "DataEnvio"
                
            Case ORD_POR_LIMITE
                sOrdenacaoPor = "DataLimite"
                
            Case Else
                gError 68733
                  
    End Select

    lErro = objRelOpcoes.IncluirParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> AD_BOOL_TRUE Then gError 68734
   
    sOrd = ComboOrdenacao.ListIndex
    lErro = objRelOpcoes.IncluirParametro("NORDENACAO", sOrd)
    If lErro <> AD_BOOL_TRUE Then gError 68735
   
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCodRequisitante_I, sCodRequisitante_F, sNomeReq_I, sNomeReq_F, sCcl_I, sCcl_F, sCodFilial_I, sCodFilial_F, sNomeFilial_I, sNomeFilial_F, sCodRequisicao_I, sCodRequisicao_F, sOrdenacaoPor, sOrd)
    If lErro <> SUCESSO Then gError 68736

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 68713 To 68736, 72516
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172884)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sCodRequisitante_I As String, sCodRequisitante_F As String, sNomeReq_I As String, sNomeReq_F As String, sCcl_I As String, sCcl_F As String, sCodFilial_I As String, sCodFilial_F As String, sNomeFilial_I As String, sNomeFilial_F As String, sCodRequisicao_I As String, sCodRequisicao_F As String) As Long
'Verifica se os par�metros iniciais s�o maiores que os finais

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
        
        If StrParaLong(sCodRequisitante_I) > StrParaLong(sCodRequisitante_F) Then gError 68737
        
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
        If sNomeReq_I > sNomeReq_F Then gError 68738
    End If
    
    'critica Ccl Inicial e Final
    If CclDe.ClipText <> "" Then
        lErro = CF("Ccl_Formata", CclDe.Text, sCclFormata, iCclPreenchida)
        If lErro <> SUCESSO Then gError 68739
        
        sCcl_I = sCclFormata
    Else
        sCcl_I = ""
    End If
    
    If CclAte.ClipText <> "" Then
        lErro = CF("Ccl_Formata", CclAte.Text, sCclFormata, iCclPreenchida)
        If lErro <> SUCESSO Then gError 68740
        
        sCcl_F = sCclFormata
    Else
        sCcl_F = ""
    End If
            
    If sCcl_I <> "" And sCcl_F <> "" Then
        
        If sCcl_I > sCcl_F Then gError 68741
        
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

        If StrParaInt(sCodFilial_I) > StrParaInt(sCodFilial_F) Then gError 68742

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
        If sNomeFilial_I > sNomeFilial_F Then gError 68743
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
        
        If StrParaLong(sCodRequisicao_I) > StrParaLong(sCodRequisicao_F) Then gError 68744
        
    End If
    
    'data de Envio inicial n�o pode ser maior que a final
    If Trim(DataEnvioDe.ClipText) <> "" And Trim(DataEnvioAte.ClipText) <> "" Then
    
         If CDate(DataEnvioDe.Text) > CDate(DataEnvioAte.Text) Then gError 68745
    
    End If
    
    
    'data Limite inicial n�o pode ser maior que a data limite final
    If Trim(DataLimiteDe.ClipText) <> "" And Trim(DataLimiteAte.ClipText) <> "" Then
    
         If CDate(DataLimiteDe.Text) > CDate(DataLimiteAte.Text) Then gError 68746
    
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                
        Case 68737
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_INICIAL_MAIOR", gErr)
            CodRequisitanteDe.SetFocus
                
        Case 68738
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_INICIAL_MAIOR", gErr)
            NomeReqDe.SetFocus
            
        Case 68739, 68740
        
        Case 68741
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_INICIAL_MAIOR", gErr)
            CclDe.SetFocus
                
        Case 68742
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            CodigoFilialDe.SetFocus
            
        Case 68743
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            NomeFilialDe.SetFocus
            
        Case 68744
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISICAO_INICIAL_MAIOR", gErr)
            CodRequisicaoDe.SetFocus
            
        Case 68745
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAENVIO_INICIAL_MAIOR", gErr)
            DataEnvioDe.SetFocus
            
        Case 68746
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATALIMITE_INICIAL_MAIOR", gErr)
            DataLimiteDe.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172885)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCodigo_I As String, sCodigo_F As String, sNome_I As String, sNome_F As String, sCcl_I As String, sCcl_F As String, sCodFilial_I As String, sCodFilial_F As String, sNomeFilial_I As String, sNomeFilial_F As String, sCodRequisicao_I As String, sCodRequisicao_F As String, sOrdenacaoPor As String, sOrd As String) As Long
'monta a express�o de sele��o de relat�rio

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

   If sCodigo_I <> "" Then sExpressao = "CodReq >= " & Forprint_ConvLong(StrParaLong(sCodigo_I))

   If sCodigo_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CodReq <= " & Forprint_ConvLong(StrParaLong(sCodigo_F))

    End If

   If sNome_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "NomeReqInic"

    End If
    
    If sNome_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "NomeReqFim"

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172886)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'l� os par�metros armazenados no bd e exibe na tela

Dim lErro As Long, iTipoOrd As Integer, iAscendente As Integer
Dim sParam As String
Dim sTipoCliente As String, iTipo As Integer
Dim sOrdenacaoPor As String
Dim sCclMascarado As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 68747
   
    'pega Codigo inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODREQINIC", sParam)
    If lErro <> SUCESSO Then gError 68748
    
    CodRequisitanteDe.Text = sParam
    Call CodRequisitanteDe_Validate(bSGECancelDummy)
    
    'pega  Codigo final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODREQFIM", sParam)
    If lErro <> SUCESSO Then gError 68749
    
    CodRequisitanteAte.Text = sParam
    Call CodRequisitanteAte_Validate(bSGECancelDummy)
                
    'pega  Nome Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEREQINIC", sParam)
    If lErro <> SUCESSO Then gError 68750
                   
    NomeReqDe.Text = sParam
    Call NomeReqDe_Validate(bSGECancelDummy)
    
    'pega  Nome Final e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEREQFIM", sParam)
    If lErro <> SUCESSO Then gError 68751
                   
    NomeReqAte.Text = sParam
    Call NomeReqAte_Validate(bSGECancelDummy)
                        
    'pega  Ccl Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TCCLINIC", sParam)
    If lErro <> SUCESSO Then gError 68752
                   
    If Len(Trim(sParam)) > 0 Then
        lErro = Mascara_MascararCcl(sParam, sCclMascarado)
        If lErro <> SUCESSO Then gError 68753
        CclDe.PromptInclude = False
        CclDe.Text = sCclMascarado
        CclDe.PromptInclude = True
        
    End If
    Call CclDe_Validate(bSGECancelDummy)
                          
                          
    'pega  Ccl Final e exibe
    lErro = objRelOpcoes.ObterParametro("TCCLFIM", sParam)
    If lErro <> SUCESSO Then gError 68754
                   
    If Len(Trim(sParam)) > 0 Then
    
        lErro = Mascara_MascararCcl(sParam, sCclMascarado)
        If lErro <> SUCESSO Then gError 68755
        
        CclAte.PromptInclude = False
        CclAte.Text = sCclMascarado
        CclAte.PromptInclude = True
        
    End If
    Call CclAte_Validate(bSGECancelDummy)
                              
    'pega Codigo inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILIALINIC", sParam)
    If lErro <> SUCESSO Then gError 68756
    
    CodigoFilialDe.Text = sParam
    Call CodigoFilialDe_Validate(bSGECancelDummy)
    
    'pega  Codigo final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILIALFIM", sParam)
    If lErro <> SUCESSO Then gError 68757
    
    CodigoFilialAte.Text = sParam
    Call CodigoFilialAte_Validate(bSGECancelDummy)
                
    'pega  Nome Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFILIALINIC", sParam)
    If lErro <> SUCESSO Then gError 68758
                   
    NomeFilialDe.Text = sParam
    Call NomeFilialDe_Validate(bSGECancelDummy)
    
    'pega  Nome Final e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFILIALFIM", sParam)
    If lErro <> SUCESSO Then gError 68759
                   
    NomeFilialAte.Text = sParam
    Call NomeFilialAte_Validate(bSGECancelDummy)
                        
    'pega Codigo inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODREQUISICAOINIC", sParam)
    If lErro <> SUCESSO Then gError 68760
    
    CodRequisicaoDe.Text = sParam
    
    'pega  Codigo final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODREQUISICAOFIM", sParam)
    If lErro <> SUCESSO Then gError 68761
    
    CodRequisicaoAte.Text = sParam
                                   
    'pega DataEnvio inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DENVINIC", sParam)
    If lErro <> SUCESSO Then gError 68762
    
    Call DateParaMasked(DataEnvioDe, CDate(sParam))
    
    'pega data de envio final e exibe
    lErro = objRelOpcoes.ObterParametro("DENVFIM", sParam)
    If lErro <> SUCESSO Then gError 68763

    Call DateParaMasked(DataEnvioAte, CDate(sParam))

    'pega data limite inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DLIMINIC", sParam)
    If lErro <> SUCESSO Then gError 68764

    Call DateParaMasked(DataLimiteDe, CDate(sParam))
       
    'pega data limite final e exibe
    lErro = objRelOpcoes.ObterParametro("DLIMFIM", sParam)
    If lErro <> SUCESSO Then gError 68765

    Call DateParaMasked(DataLimiteAte, CDate(sParam))
       
    'pega 'Inclui Requisicoes Atendidas' e exibe
    lErro = objRelOpcoes.ObterParametro("NREQATEND", sParam)
    If lErro <> SUCESSO Then gError 72517

    If sParam = "1" Then
        CheckRequisicoes.Value = 1
    Else
        CheckRequisicoes.Value = 0
    End If
       
    lErro = objRelOpcoes.ObterParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> SUCESSO Then gError 68766
    
    Select Case sOrdenacaoPor
        
            Case "CodReq"
            
                ComboOrdenacao.ListIndex = ORD_POR_CODIGO
            
            Case "NomeReq"
            
                ComboOrdenacao.ListIndex = ORD_POR_NOME
                
            Case "Ccl"
            
                ComboOrdenacao.ListIndex = ORD_POR_CCL
                
            Case "DataEnvio"
                ComboOrdenacao.ListIndex = ORD_POR_ENVIO
                
            Case "DataLimite"
                ComboOrdenacao.ListIndex = ORD_POR_LIMITE
                
            Case Else
                gError 68767
                  
    End Select
        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 68747 To 68767, 72517
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172887)

    End Select

    Exit Function

End Function
Private Sub CodigoFilialDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_CodigoFilialDe_Validate

    If Len(Trim(CodigoFilialDe.Text)) > 0 Then

        objFilialEmpresa.iCodFilial = StrParaInt(CodigoFilialDe.Text)
        'L� o c�digo informado
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 68782
        
        'Se n�o encontrou a Filial ==> erro
        If lErro = 27378 Then gError 68783

    End If

    Exit Sub

Erro_CodigoFilialDe_Validate:

    Cancel = True


    Select Case gErr

        Case 68782

        Case 68783
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172888)

    End Select

    Exit Sub

End Sub
Private Sub CodigoFilialAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_CodigoFilialAte_Validate

    If Len(Trim(CodigoFilialAte.Text)) > 0 Then

        objFilialEmpresa.iCodFilial = StrParaInt(CodigoFilialAte.Text)
        'L� o c�digo informado
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 68784
        
        'Se n�o encontrou a Filial ==> erro
        If lErro = 27378 Then gError 68785

    End If

    Exit Sub

Erro_CodigoFilialAte_Validate:

    Cancel = True


    Select Case gErr

        Case 68784

        Case 68785
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172889)

    End Select

    Exit Sub

End Sub

Private Sub CodRequisitanteDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objRequisitante As New ClassRequisitante

On Error GoTo Erro_CodRequisitanteDe_Validate

    If Len(Trim(CodRequisitanteDe.Text)) > 0 Then

        objRequisitante.lCodigo = StrParaLong(CodRequisitanteDe.Text)
        'L� o c�digo informado
        lErro = CF("Requisitante_Le", objRequisitante)
        If lErro <> SUCESSO And lErro <> 49084 Then gError 68768

        'Se n�o encontrou o Requisitante ==> erro
        If lErro = 49084 Then gError 68769
        
    End If

    Exit Sub

Erro_CodRequisitanteDe_Validate:

    Cancel = True


    Select Case gErr

        Case 68768

        Case 68769
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_INEXISTENTE", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172890)

    End Select

    Exit Sub
    
End Sub


Private Sub CodRequisitanteAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objRequisitante As New ClassRequisitante

On Error GoTo Erro_CodRequisitanteAte_Validate

    If Len(Trim(CodRequisitanteAte.Text)) > 0 Then

        objRequisitante.lCodigo = StrParaLong(CodRequisitanteAte.Text)
        'L� o c�digo informado
        lErro = CF("Requisitante_Le", objRequisitante)
        If lErro <> SUCESSO And lErro <> 49084 Then gError 68770

        'Se n�o encontrou o Requisitante ==> erro
        If lErro = 49084 Then gError 68771
        
    End If

    Exit Sub

Erro_CodRequisitanteAte_Validate:

    Cancel = True


    Select Case gErr

        Case 68770

        Case 68771
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_INEXISTENTE", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172891)

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
        If lErro <> SUCESSO Then gError 68786

        'Carrega a Filial com o Nome informado
        For Each objFilialEmpresa In colFiliais
            If objFilialEmpresa.sNome = NomeFilialDe.Text Then
                bAchou = True
                Exit For
            End If
        Next

        'Se n�o encontrou Filial com o Nome informado ==> erro
        If bAchou = False Then gError 68787
        
        NomeFilialDe.Text = objFilialEmpresa.sNome

    End If

    Exit Sub

Erro_NomeFilialDe_Validate:

    Cancel = True

    Select Case gErr

        Case 68786

        Case 68787
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, NomeFilialDe.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172892)

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
        If lErro <> SUCESSO Then gError 68788

        'Carrega a Filial com o Nome informado
        For Each objFilialEmpresa In colFiliais
            If objFilialEmpresa.sNome = NomeFilialAte.Text Then
                bAchou = True
                Exit For
            End If
        Next

        'Se n�o encontrou Filial com o Nome informado ==> erro
        If bAchou = False Then gError 68789

        NomeFilialAte.Text = objFilialEmpresa.sNome

    End If

    Exit Sub

Erro_NomeFilialAte_Validate:

    Cancel = True


    Select Case gErr

        Case 68788

        Case 68789
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, NomeFilialAte.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172893)

    End Select

Exit Sub

End Sub

Private Sub NomeReqDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objRequisitante As New ClassRequisitante

On Error GoTo Erro_NomeReqDe_Validate

    If Len(Trim(NomeReqDe.Text)) > 0 Then

        objRequisitante.sNomeReduzido = NomeReqDe.Text
        'L� o Requisitante informado
        lErro = CF("Requisitante_Le_NomeReduzido", objRequisitante)
        If lErro <> SUCESSO And lErro <> 51152 Then gError 68772

        'Se n�o encontrou o Requisitante ==> erro
        If lErro = 51152 Then gError 68773
        
        NomeReqDe.Text = objRequisitante.sNomeReduzido
        
    End If

    Exit Sub

Erro_NomeReqDe_Validate:

    Cancel = True


    Select Case gErr

        Case 68772

        Case 68773
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_INEXISTENTE", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172894)

    End Select

Exit Sub

End Sub

Private Sub NomeReqAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objRequisitante As New ClassRequisitante

On Error GoTo Erro_NomeReqAte_Validate

    If Len(Trim(NomeReqAte.Text)) > 0 Then

        objRequisitante.sNomeReduzido = NomeReqAte.Text
        'L� o Requisitante informado
        lErro = CF("Requisitante_Le_NomeReduzido", objRequisitante)
        If lErro <> SUCESSO And lErro <> 51152 Then gError 68774

        'Se n�o encontrou o Requisitante ==> erro
        If lErro = 51152 Then gError 68775
        
        NomeReqAte.Text = objRequisitante.sNomeReduzido
        
    End If

    Exit Sub

Erro_NomeReqAte_Validate:

    Cancel = True


    Select Case gErr

        Case 68774

        Case 68775
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_INEXISTENTE", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172895)

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
        If lErro <> SUCESSO Then gError 68776
        
        objCcl.sCcl = sCclFormata
        
        'L� o Ccl informado
        lErro = CF("Ccl_Le", objCcl)
        If lErro <> SUCESSO And lErro <> 5599 Then gError 68777

        'Se n�o encontrou o Ccl ==> erro
        If lErro = 5599 Then gError 68778
            
    End If

    Exit Sub

Erro_CclDe_Validate:

    Cancel = True

    Select Case gErr

        Case 68775, 68776

        Case 68778
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_INEXISTENTE", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172896)

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
        If lErro <> SUCESSO Then gError 68779
        
        objCcl.sCcl = sCclFormata
        
        'L� o Ccl informado
        lErro = CF("Ccl_Le", objCcl)
        If lErro <> SUCESSO And lErro <> 5599 Then gError 68780

        'Se n�o encontrou o Ccl ==> erro
        If lErro = 5599 Then gError 68781
        
    End If

    Exit Sub

Erro_CclAte_Validate:

    Cancel = True

    Select Case gErr

        Case 68779, 68780

        Case 68781
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_INEXISTENTE", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172897)

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
    Caption = "Requisitantes x Requisi��es"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpReqRequisicoes"
    
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




Private Sub LabelDataEnvioAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataEnvioAte, Source, X, Y)
End Sub

Private Sub LabelDataEnvioAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataEnvioAte, Button, Shift, X, Y)
End Sub

Private Sub LabelDataLimiteDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataLimiteDe, Source, X, Y)
End Sub

Private Sub LabelDataLimiteDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataLimiteDe, Button, Shift, X, Y)
End Sub

Private Sub LabelDataLimiteAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataLimiteAte, Source, X, Y)
End Sub

Private Sub LabelDataLimiteAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataLimiteAte, Button, Shift, X, Y)
End Sub

Private Sub LabelDataEnvioDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataEnvioDe, Source, X, Y)
End Sub

Private Sub LabelDataEnvioDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataEnvioDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCodRequisicaoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodRequisicaoDe, Source, X, Y)
End Sub

Private Sub LabelCodRequisicaoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodRequisicaoDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCodRequisicaoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodRequisicaoAte, Source, X, Y)
End Sub

Private Sub LabelCodRequisicaoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodRequisicaoAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCclDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCclDE, Source, X, Y)
End Sub

Private Sub LabelCclDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCclDE, Button, Shift, X, Y)
End Sub

Private Sub LabelCclAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCClAte, Source, X, Y)
End Sub

Private Sub LabelCclAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCClAte, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeReqAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeReqAte, Source, X, Y)
End Sub

Private Sub LabelNomeReqAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeReqAte, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeReqDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeReqDe, Source, X, Y)
End Sub

Private Sub LabelNomeReqDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeReqDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCodRequisitanteDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodRequisitanteDe, Source, X, Y)
End Sub

Private Sub LabelCodRequisitanteDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodRequisitanteDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCodRequisitanteAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodRequisitanteAte, Source, X, Y)
End Sub

Private Sub LabelCodRequisitanteAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodRequisitanteAte, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeAte, Source, X, Y)
End Sub

Private Sub LabelNomeAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeAte, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeDe, Source, X, Y)
End Sub

Private Sub LabelNomeDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeDe, Button, Shift, X, Y)
End Sub

