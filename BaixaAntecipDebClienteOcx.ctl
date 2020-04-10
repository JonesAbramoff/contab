VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl BaixaAntecipDebClienteOcx 
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9390
   ScaleHeight     =   5790
   ScaleMode       =   0  'User
   ScaleWidth      =   9390
   Begin VB.Frame FramePrincipal 
      BorderStyle     =   0  'None
      Height          =   4815
      Index           =   3
      Left            =   150
      TabIndex        =   64
      Top             =   780
      Visible         =   0   'False
      Width           =   9120
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4560
         TabIndex        =   107
         Tag             =   "1"
         Top             =   1560
         Width           =   870
      End
      Begin VB.CommandButton CTBBotaoImprimir 
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7725
         TabIndex        =   77
         Top             =   60
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   6330
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   900
         Width           =   2700
      End
      Begin VB.CommandButton CTBBotaoLimparGrid 
         Caption         =   "Limpar Grid"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6270
         TabIndex        =   75
         Top             =   60
         Width           =   1245
      End
      Begin VB.CommandButton CTBBotaoModeloPadrao 
         Caption         =   "Modelo Padrão"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6270
         TabIndex        =   74
         Top             =   375
         Width           =   2700
      End
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2850
         IntegralHeight  =   0   'False
         Left            =   6330
         TabIndex        =   73
         Top             =   1530
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.CheckBox CTBLancAutomatico 
         Caption         =   "Recalcula Automaticamente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3450
         TabIndex        =   72
         Top             =   945
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin VB.CheckBox CTBAglutina 
         Height          =   210
         Left            =   4470
         TabIndex        =   71
         Top             =   2565
         Width           =   870
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   70
         Top             =   2190
         Width           =   1770
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   65
         Top             =   3330
         Width           =   5895
         Begin VB.Label CTBCclLabel 
            AutoSize        =   -1  'True
            Caption         =   "Centro de Custo:"
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
            Left            =   240
            TabIndex        =   69
            Top             =   660
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.Label CTBLabel7 
            AutoSize        =   -1  'True
            Caption         =   "Conta:"
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
            Left            =   1125
            TabIndex        =   68
            Top             =   300
            Width           =   570
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   67
            Top             =   285
            Width           =   3720
         End
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   66
            Top             =   645
            Visible         =   0   'False
            Width           =   3720
         End
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4680
         TabIndex        =   78
         Top             =   1920
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBConta 
         Height          =   225
         Left            =   525
         TabIndex        =   79
         Top             =   1860
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBDebito 
         Height          =   225
         Left            =   3435
         TabIndex        =   80
         Top             =   1890
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBCredito 
         Height          =   225
         Left            =   2280
         TabIndex        =   81
         Top             =   1830
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBCcl 
         Height          =   225
         Left            =   1545
         TabIndex        =   82
         Top             =   1875
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown CTBUpDown 
         Height          =   300
         Left            =   1650
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   585
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox CTBDataContabil 
         Height          =   300
         Left            =   570
         TabIndex        =   84
         Top             =   570
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBLote 
         Height          =   300
         Left            =   5580
         TabIndex        =   85
         Top             =   150
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBDocumento 
         Height          =   300
         Left            =   3795
         TabIndex        =   86
         Top             =   135
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid CTBGridContabil 
         Height          =   1770
         Left            =   15
         TabIndex        =   87
         Top             =   1230
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   3122
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin MSComctlLib.TreeView CTBTvwContas 
         Height          =   2850
         Left            =   6330
         TabIndex        =   88
         Top             =   1530
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   5027
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView CTBTvwCcls 
         Height          =   2850
         Left            =   6330
         TabIndex        =   89
         Top             =   1530
         Visible         =   0   'False
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   5027
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin VB.Label CTBLabel1 
         AutoSize        =   -1  'True
         Caption         =   "Modelo:"
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
         Left            =   6330
         TabIndex        =   106
         Top             =   690
         Width           =   690
      End
      Begin VB.Label CTBLabel21 
         Caption         =   "Origem:"
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
         Left            =   30
         TabIndex        =   105
         Top             =   180
         Width           =   720
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   104
         Top             =   135
         Width           =   1530
      End
      Begin VB.Label CTBLabel14 
         Caption         =   "Período:"
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
         Left            =   4230
         TabIndex        =   103
         Top             =   645
         Width           =   735
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5010
         TabIndex        =   102
         Top             =   615
         Width           =   1185
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   101
         Top             =   600
         Width           =   1185
      End
      Begin VB.Label CTBLabel13 
         Caption         =   "Exercício:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1995
         TabIndex        =   100
         Top             =   630
         Width           =   870
      End
      Begin VB.Label CTBLabel5 
         AutoSize        =   -1  'True
         Caption         =   "Lançamentos"
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
         Left            =   45
         TabIndex        =   99
         Top             =   990
         Width           =   1140
      End
      Begin VB.Label CTBLabelHistoricos 
         Caption         =   "Históricos"
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
         Left            =   6345
         TabIndex        =   98
         Top             =   1275
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label CTBLabelContas 
         Caption         =   "Plano de Contas"
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
         Left            =   6345
         TabIndex        =   97
         Top             =   1275
         Width           =   2340
      End
      Begin VB.Label CTBLabelCcl 
         Caption         =   "Centros de Custo / Lucro"
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
         Left            =   6345
         TabIndex        =   96
         Top             =   1275
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.Label CTBLabelTotais 
         Caption         =   "Totais:"
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
         Height          =   345
         Left            =   1980
         TabIndex        =   95
         Top             =   3030
         Width           =   615
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3885
         TabIndex        =   94
         Top             =   3015
         Width           =   1155
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2640
         TabIndex        =   93
         Top             =   3015
         Width           =   1155
      End
      Begin VB.Label CTBLabel8 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
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
         Left            =   45
         TabIndex        =   92
         Top             =   600
         Width           =   480
      End
      Begin VB.Label CTBLabelDoc 
         AutoSize        =   -1  'True
         Caption         =   "Documento:"
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
         Left            =   2700
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   91
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label CTBLabelLote 
         AutoSize        =   -1  'True
         Caption         =   "Lote:"
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
         Left            =   5100
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   90
         Top             =   180
         Width           =   450
      End
   End
   Begin VB.Frame FramePrincipal 
      BorderStyle     =   0  'None
      Height          =   4815
      Index           =   1
      Left            =   150
      TabIndex        =   5
      Top             =   780
      Width           =   9120
      Begin VB.Frame FrameFiltrosAdiantamento 
         Caption         =   "Conta Corrente"
         Height          =   1575
         Left            =   3030
         TabIndex        =   16
         Top             =   2400
         Width           =   5535
         Begin VB.OptionButton CtaCorrenteApenas 
            Caption         =   "Apenas:"
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
            Left            =   360
            TabIndex        =   19
            Top             =   960
            Width           =   1095
         End
         Begin VB.OptionButton CtaCorrenteTodas 
            Caption         =   "Todas"
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
            Left            =   360
            TabIndex        =   18
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.ComboBox ContaCorrenteSeleciona 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "BaixaAntecipDebClienteOcx.ctx":0000
            Left            =   1560
            List            =   "BaixaAntecipDebClienteOcx.ctx":0002
            TabIndex        =   17
            Top             =   945
            Width           =   2055
         End
      End
      Begin VB.Frame FrameTipoBaixas 
         Caption         =   "Selecione os documento que serão baixados"
         Height          =   615
         Left            =   510
         TabIndex        =   32
         Top             =   360
         Width           =   8055
         Begin VB.OptionButton TipoBaixaAdiantamento 
            Caption         =   "Adiantamentos"
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
            Left            =   1320
            TabIndex        =   34
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton TipoBaixaDebitos 
            Caption         =   "Débitos / Devoluções"
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
            TabIndex        =   33
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame FrameCliente 
         Caption         =   "Cliente"
         Height          =   885
         Left            =   510
         TabIndex        =   27
         Top             =   1200
         Width           =   8055
         Begin MSMask.MaskEdBox Cliente 
            Height          =   300
            Left            =   1080
            TabIndex        =   29
            Top             =   390
            Width           =   3345
            _ExtentX        =   5900
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   "_"
         End
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   5490
            TabIndex        =   28
            Top             =   390
            Width           =   1815
         End
         Begin VB.Label LabelFil 
            Caption         =   "Filial:"
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
            Height          =   255
            Left            =   4875
            TabIndex        =   31
            Top             =   435
            Width           =   615
         End
         Begin VB.Label LabelCliente 
            Caption         =   "Cliente:"
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
            Height          =   255
            Left            =   210
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   30
            Top             =   435
            Width           =   675
         End
      End
      Begin VB.Frame FrameDataEmissao 
         Caption         =   "Data da Emissão"
         Height          =   1575
         Left            =   510
         TabIndex        =   20
         Top             =   2400
         Width           =   2175
         Begin MSComCtl2.UpDown UpDownEmissaoInic 
            Height          =   300
            Left            =   1710
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   435
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox EmissaoInic 
            Height          =   300
            Left            =   630
            TabIndex        =   22
            Top             =   450
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownEmissaoFim 
            Height          =   300
            Left            =   1725
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   960
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox EmissaoFim 
            Height          =   300
            Left            =   630
            TabIndex        =   24
            Top             =   960
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label LabelDataEmissaoAte 
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
            Height          =   255
            Left            =   195
            TabIndex        =   26
            Top             =   990
            Width           =   375
         End
         Begin VB.Label LabelDataEmissaoDe 
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
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   480
            Width           =   375
         End
      End
      Begin VB.Frame FrameFiltrosDebito 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   2760
         TabIndex        =   6
         Top             =   2400
         Visible         =   0   'False
         Width           =   5895
         Begin VB.Frame FrameNumDebito 
            Caption         =   "Nº Débito / Devolução"
            Height          =   735
            Left            =   270
            TabIndex        =   7
            Top             =   0
            Width           =   5535
            Begin MSMask.MaskEdBox NumDebitoInic 
               Height          =   300
               Left            =   1320
               TabIndex        =   8
               Top             =   285
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   6
               Mask            =   "999999"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox NumDebitoFim 
               Height          =   300
               Left            =   3615
               TabIndex        =   9
               Top             =   285
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   6
               Mask            =   "999999"
               PromptChar      =   " "
            End
            Begin VB.Label LabelNumDebitoDe 
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
               Height          =   255
               Left            =   960
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   11
               Top             =   315
               Width           =   375
            End
            Begin VB.Label LabelNumDebitoAte 
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
               Height          =   255
               Left            =   3195
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   10
               Top             =   315
               Width           =   375
            End
         End
         Begin VB.Frame FrameTipoDebito 
            Caption         =   "Tipo de Débito"
            Height          =   735
            Left            =   270
            TabIndex        =   12
            Top             =   840
            Width           =   5535
            Begin VB.OptionButton TipoDebitoTodos 
               Caption         =   "Todos"
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
               Left            =   960
               TabIndex        =   15
               Top             =   310
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton TipoDebitoApenas 
               Caption         =   "Apenas:"
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
               TabIndex        =   14
               Top             =   310
               Width           =   1095
            End
            Begin VB.ComboBox TipoDebitoSeleciona 
               Height          =   315
               Left            =   3360
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   280
               Width           =   1815
            End
         End
      End
   End
   Begin VB.Frame FramePrincipal 
      BorderStyle     =   0  'None
      Height          =   4815
      Index           =   2
      Left            =   150
      TabIndex        =   35
      Top             =   780
      Width           =   9120
      Begin VB.Frame FrameDebitos 
         Caption         =   "Débitos / Devoluções"
         Height          =   2925
         Left            =   390
         TabIndex        =   55
         Top             =   60
         Width           =   8295
         Begin MSMask.MaskEdBox FilialDebito 
            Height          =   225
            Left            =   480
            TabIndex        =   63
            Top             =   1440
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "9999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NumDebito 
            Height          =   225
            Left            =   3840
            TabIndex        =   61
            Top             =   1440
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "999999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataDebito 
            Height          =   240
            Left            =   1800
            TabIndex        =   58
            Top             =   1440
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox SaldoDebito 
            Height          =   225
            Left            =   5880
            TabIndex        =   59
            Top             =   1440
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDebito 
            Height          =   225
            Left            =   4680
            TabIndex        =   60
            Top             =   1440
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TipoDebito 
            Height          =   225
            Left            =   3120
            TabIndex        =   62
            Top             =   1440
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   50
            PromptChar      =   "_"
         End
         Begin VB.CheckBox DebitoSeleciona 
            Height          =   255
            Left            =   480
            TabIndex        =   56
            Top             =   960
            Width           =   615
         End
         Begin MSFlexGridLib.MSFlexGrid GridDebitos 
            Height          =   1935
            Left            =   240
            TabIndex        =   57
            Top             =   240
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   3413
            _Version        =   393216
         End
      End
      Begin VB.Frame FrameAdiantamentos 
         Caption         =   "Adiantamentos"
         Height          =   2925
         Left            =   390
         TabIndex        =   36
         Top             =   60
         Width           =   8295
         Begin MSMask.MaskEdBox FilialAdiantamento 
            Height          =   225
            Left            =   720
            TabIndex        =   54
            Top             =   1440
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "9999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataAdiantamento 
            Height          =   225
            Left            =   1560
            TabIndex        =   51
            Top             =   1440
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox CtaCorrenteAdiantamento 
            Height          =   225
            Left            =   2880
            TabIndex        =   49
            Top             =   1440
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   2
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MeioPagtoAdiantamento 
            Height          =   225
            Left            =   4200
            TabIndex        =   50
            Top             =   1440
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   2
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox ValorAdiantamento 
            Height          =   225
            Left            =   5640
            TabIndex        =   52
            Top             =   1440
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox SaldoAdiantamento 
            Height          =   225
            Left            =   6720
            TabIndex        =   53
            Top             =   1440
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.CheckBox AdiantamentoSeleciona 
            Height          =   255
            Left            =   240
            TabIndex        =   48
            Top             =   1080
            Width           =   615
         End
         Begin MSFlexGridLib.MSFlexGrid GridAdiantamentos 
            Height          =   1935
            Left            =   240
            TabIndex        =   37
            Top             =   240
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   3413
            _Version        =   393216
            Rows            =   8
            Cols            =   7
         End
      End
      Begin VB.Frame FrameBaixa 
         Caption         =   "Baixa"
         Height          =   1575
         Left            =   390
         TabIndex        =   38
         Top             =   3060
         Width           =   8295
         Begin VB.ComboBox ContaCorrente 
            Height          =   315
            Left            =   1920
            TabIndex        =   45
            Top             =   700
            Width           =   1815
         End
         Begin VB.TextBox Historico 
            Height          =   320
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   43
            Top             =   1175
            Width           =   5655
         End
         Begin MSComCtl2.UpDown UpDownDataBaixa 
            Height          =   300
            Left            =   3015
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   250
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataBaixa 
            Height          =   300
            Left            =   1920
            TabIndex        =   40
            Top             =   250
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label LabelValorBaixa2 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   6360
            TabIndex        =   47
            Top             =   250
            Width           =   1215
         End
         Begin VB.Label LabelValorBaixa 
            AutoSize        =   -1  'True
            Caption         =   "Valor da Baixa:"
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
            Left            =   4920
            TabIndex        =   46
            Top             =   303
            Width           =   1305
         End
         Begin VB.Label LabelCtaCorrente 
            AutoSize        =   -1  'True
            Caption         =   "Conta Corrente:"
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
            Left            =   480
            TabIndex        =   44
            Top             =   760
            Width           =   1350
         End
         Begin VB.Label LabelHistorico 
            AutoSize        =   -1  'True
            Caption         =   "Histórico:"
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
            Left            =   960
            TabIndex        =   42
            Top             =   1235
            Width           =   825
         End
         Begin VB.Label LabelDataBaixa 
            AutoSize        =   -1  'True
            Caption         =   "Data da Baixa:"
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
            Left            =   600
            TabIndex        =   41
            Top             =   303
            Width           =   1275
         End
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   555
      Left            =   7440
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1695
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "BaixaAntecipDebClienteOcx.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   607
         Picture         =   "BaixaAntecipDebClienteOcx.ctx":015E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1125
         Picture         =   "BaixaAntecipDebClienteOcx.ctx":0690
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5235
      Left            =   90
      TabIndex        =   4
      Top             =   450
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   9234
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Filtros"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Adiantamentos / Débitos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Contabilização"
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
End
Attribute VB_Name = "BaixaAntecipDebClienteOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'contabilidade
Dim objGrid1 As AdmGrid
Dim objContabil As New ClassContabil

Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1

'mnemonicos contabilidade
Private Const CLIENTE_COD As String = "Cliente_Codigo"
Private Const CLIENTE_NOME As String = "Cliente_Nome"
Private Const FILIAL_COD As String = "FilialCli_Codigo"
Private Const FILIAL_NOME_RED As String = "FilialCli_Nome"
Private Const FILIAL_CONTA As String = "FilialCli_Conta_Ctb"
Private Const FILIAL_CGC_CPF As String = "FilialCli_CGC_CPF"
Private Const CONTA_COD As String = "Conta_Codigo"
Private Const CONTA_CONTABIL_CONTA As String = "Conta_Contabil_Conta"
Private Const VALOR1 As String = "ValorBaixa"
Private Const HISTORICO1 As String = "Historico"
Private Const DATA1 As String = "DataBaixa"

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iGridAtual As Integer

'Obj´s associados a eventos para chamada de browser
Private WithEvents objEventoClientes As AdmEvento
Attribute objEventoClientes.VB_VarHelpID = -1
Private WithEvents objEventoNumDebitoInic As AdmEvento
Attribute objEventoNumDebitoInic.VB_VarHelpID = -1
Private WithEvents objEventoNumDebitoFim As AdmEvento
Attribute objEventoNumDebitoFim.VB_VarHelpID = -1

'Obj global a tela que armazena todos os dados relevantes
Dim gobjBaixaAntecipDebCliente As ClassBaixaAntecipDebCliente

'Grid Adiantamentos:
Dim objGridAdiantamentos As AdmGrid
Dim iGrid_AdiantamentoSeleciona_Col As Integer
Dim iGrid_FilialAdiantamento_Col As Integer
Dim iGrid_DataAdiantamento_Col As Integer
Dim iGrid_CtaCorrenteAdiantamento_Col As Integer
Dim iGrid_MeioPagtoAdiantamento_Col As Integer
Dim iGrid_ValorAdiantamento_Col As Integer
Dim iGrid_SaldoAdiantamento_Col As Integer

'Grid Débitos:
Dim objGridDebitos As AdmGrid
Dim iGrid_DebitoSeleciona_Col As Integer
Dim iGrid_FilialDebito_Col As Integer
Dim iGrid_DataDebito_Col As Integer
Dim iGrid_TipoDebito_Col As Integer
Dim iGrid_NumDebito_Col As Integer
Dim iGrid_ValorDebito_Col As Integer
Dim iGrid_SaldoDebito_Col As Integer

'Variáveis para controle de tela
Public iAlterado As Integer
Dim iFramePrincipalAtual As Integer
Dim iFramePrincipalAlterado As Integer
Dim iUltTipoBaixa As Integer
Dim iClienteAlterado As Integer

'Definições dos TAB´s
Const TAB_FILTROS = 1
Const TAB_ADIANTAMENTOSDEBITOS = 2
Const TAB_Contabilizacao = 3

'Constantes para controle de tela
Const TIPOBAIXA_ADIANTAMENTO = 1
Const TIPOBAIXA_DEBITOS = 2
Const CONTACORRENTE_TODAS = 1
Const CONTACORRENTE_APENAS = 2
Const TIPODEBITO_TODOS = 1
Const TIPODEBITO_APENAS = 2
Const TELA_TODA = 1



'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? criar IDH Parent.HelpContextID = IDH_BAIXA_PARCELAS_RECEBER_TITULOS
    Set Form_Load_Ocx = Me
    Caption = "Baixa de Adiantamentos / Débitos Clientes"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "BaixaAntecipDebCliente"

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

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    'Grava a Baixa
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 94709

    Call Limpa_Tela_BaixaAntecipDebCliente
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 94709

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143217)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objBaixaAntecipDebCliente As New ClassBaixaAntecipDebCliente

On Error GoTo Erro_Gravar_Registro

    lErro = Valida_Gravacao()
    If lErro <> SUCESSO Then gError 94715
    
    lErro = Move_Tela_Memoria(objBaixaAntecipDebCliente)
    If lErro <> SUCESSO Then gError 92784
    
    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(DataBaixa.Text))
    If lErro <> SUCESSO Then gError 92792
    
    Select Case gobjBaixaAntecipDebCliente.iTipoBaixa
    
        Case TIPOBAIXA_ADIANTAMENTO
            
            lErro = CF("BaixaAntecipDebCli_Grava_Adiantamentos", objBaixaAntecipDebCliente, objContabil)
            If lErro <> SUCESSO Then gError 94716
        
        Case TIPOBAIXA_DEBITOS
    
            lErro = CF("BaixaAntecipDebCli_Grava_Debitos", objBaixaAntecipDebCliente, objContabil)
            If lErro <> SUCESSO Then gError 94717
    
    End Select
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:
    
    Gravar_Registro = gErr
    
    Select Case gErr
    
        Case 92784, 92792, 94715 To 94717
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143218)
            
    End Select
    
    Exit Function
    
End Function

Private Function Move_Tela_Memoria(objBaixaAntecipDebCliente As ClassBaixaAntecipDebCliente) As Long

Dim iLinha As Integer
Dim objCliente As New ClassCliente
Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    'Se o Cliente estiver preenchido lê o código no BD
    If Len(Trim(Cliente.Text)) > 0 Then
        
        objCliente.sNomeReduzido = Cliente.Text
        
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 92785
        
        'Não encontrou p Cliente --> erro
        If lErro = 12348 Then gError 92786

        objBaixaAntecipDebCliente.lCliente = objCliente.lCodigo

    End If

    If Len(Trim(Filial.Text)) > 0 Then objBaixaAntecipDebCliente.iFilialCliente = Codigo_Extrai(Filial.Text)

    If Len(Trim(LabelValorBaixa2.Caption)) > 0 Then objBaixaAntecipDebCliente.dValorBaixa = CDbl(LabelValorBaixa2.Caption)
    
    If Len(Trim(DataBaixa.ClipText)) > 0 Then
        objBaixaAntecipDebCliente.dtDataBaixa = CDate(DataBaixa.Text)
    Else
        objBaixaAntecipDebCliente.dtDataBaixa = DATA_NULA
    End If

    objBaixaAntecipDebCliente.iCtaCorrenteBaixa = Codigo_Extrai(ContaCorrente.Text)
    
    objBaixaAntecipDebCliente.sHistoricoBaixa = Historico.Text

    Set objBaixaAntecipDebCliente.colAntecipDebCliente = New Collection

    'se for adiantamento
    If TipoBaixaAdiantamento.Value = True Then
    
        objBaixaAntecipDebCliente.iTipoBaixa = TIPOBAIXA_ADIANTAMENTO
    
        For iLinha = 1 To objGridAdiantamentos.iLinhasExistentes
        
            If GridAdiantamentos.TextMatrix(iLinha, iGrid_AdiantamentoSeleciona_Col) = MARCADO Then
            
                objBaixaAntecipDebCliente.colAntecipDebCliente.Add gobjBaixaAntecipDebCliente.colAntecipDebCliente.Item(iLinha)
                
            End If
            
        Next
            
    Else
    
        objBaixaAntecipDebCliente.iTipoBaixa = TIPOBAIXA_DEBITOS
    
        For iLinha = 1 To objGridDebitos.iLinhasExistentes
        
            If GridDebitos.TextMatrix(iLinha, iGrid_DebitoSeleciona_Col) = MARCADO Then
            
                objBaixaAntecipDebCliente.colAntecipDebCliente.Add gobjBaixaAntecipDebCliente.colAntecipDebCliente.Item(iLinha)
                
            End If
            
        Next
    
    End If
    
    Move_Tela_Memoria = SUCESSO
    
    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr
    
        Case 92785
    
        Case 92786
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143219)
            
    End Select
    
    Exit Function

End Function

Private Function Valida_Gravacao() As Long

Dim iLinha As Integer

On Error GoTo Erro_Valida_Gravacao

    'Se o cliente não está preenchido => erro
    If Len(Trim(Cliente.Text)) = 0 Then gError 94710
    
    'Se a filial do cliente não está preenchida => erro
    If Len(Trim(Filial.Text)) = 0 Then gError 94711
    
    'Se o valor da baixa não foi informado = > erro
    If StrParaDbl(LabelValorBaixa2.Caption) = 0 Then gError 94713
    
    'Se a conta corrente não foi informada => erro
    If Len(Trim(ContaCorrente.Text)) = 0 Then gError 94714
    
    'se for adiantamento
    If TipoBaixaAdiantamento.Value = True Then
    
        For iLinha = 1 To objGridAdiantamentos.iLinhasExistentes
        
            If GridAdiantamentos.TextMatrix(iLinha, iGrid_AdiantamentoSeleciona_Col) = MARCADO Then
            
                If CDate(GridAdiantamentos.TextMatrix(iLinha, iGrid_DataAdiantamento_Col)) > CDate(DataBaixa.Text) Then gError 94760
                
            End If
            
        Next
            
    Else
    
        For iLinha = 1 To objGridDebitos.iLinhasExistentes
        
            If GridDebitos.TextMatrix(iLinha, iGrid_DebitoSeleciona_Col) = MARCADO Then
            
                If Len(GridDebitos.TextMatrix(iLinha, iGrid_DataDebito_Col)) > 0 Then
            
                    If CDate(GridDebitos.TextMatrix(iLinha, iGrid_DataDebito_Col)) > CDate(DataBaixa.Text) Then gError 92783
                
                End If
                
            End If
            
        Next
    
    End If
    
    Valida_Gravacao = SUCESSO
    
    Exit Function

Erro_Valida_Gravacao:

    Valida_Gravacao = gErr
    
    Select Case gErr
    
        Case 92783
            Call Rotina_Erro(vbOKOnly, "ERRO_DATADOCUMENTO_MAIOR_DATABAIXA", gErr, GridDebitos.TextMatrix(iLinha, iGrid_DataDebito_Col), iLinha, DataBaixa.Text)
            
        Case 94710
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
            
        Case 94711
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
            
        Case 94713
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORBAIXAR_ANTECIPDEBCLI_NAO_INFORMADO", gErr)

        Case 94714
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_ANTECIPDEBCLI_NAO_INFORMADA", gErr)
        
        Case 94760
            Call Rotina_Erro(vbOKOnly, "ERRO_DATADOCUMENTO_MAIOR_DATABAIXA", gErr, GridAdiantamentos.TextMatrix(iLinha, iGrid_DataAdiantamento_Col), iLinha, DataBaixa.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143220)
            
    End Select
    
    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Verifica se houve alterações e confirma se deseja salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 94665

    'Limpa a tela
    Call Limpa_Tela_BaixaAntecipDebCliente

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 94665

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143221)

    End Select

    Exit Sub
   
End Sub

Private Sub ContaCorrenteSeleciona_Click()
    If ContaCorrenteSeleciona.ListIndex <> -1 Then
        iFramePrincipalAlterado = REGISTRO_ALTERADO
        iAlterado = REGISTRO_ALTERADO
    End If
End Sub

Private Sub GridAdiantamentos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Verifica se foi o botao direito do mouse que foi pressionado
    If Button = vbRightButton Then
    
        iGridAtual = GRID_ADIANTAMENTOS
        
        'Seta objTela como a Tela de Baixas a Receber
        Set PopUpMenuGridMD.objTela = Me
        
        'Chama o Menu PopUp
        PopUpMenuGridMD.PopupMenu PopUpMenuGridMD.mnuGrid, vbPopupMenuRightButton
        
        'Limpa o objTela
        Set PopUpMenuGridMD.objTela = Nothing
        
    End If
    
End Sub

Private Sub GridDebitos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Verifica se foi o botao direito do mouse que foi pressionado
    If Button = vbRightButton Then
    
        iGridAtual = GRID_DEVOLUCOES
        
        'Seta objTela como a Tela de Baixas a Receber
        Set PopUpMenuGridMD.objTela = Me
        
        'Chama o Menu PopUp
        PopUpMenuGridMD.PopupMenu PopUpMenuGridMD.mnuGrid, vbPopupMenuRightButton
        
        'Limpa o objTela
        Set PopUpMenuGridMD.objTela = Nothing
        
    End If
    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Private Sub Unload(objme As Object)

   RaiseEvent Unload

End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    'm_Caption = New_Caption
End Property

'***** fim do trecho a ser copiado ******

Private Sub GridAdiantamentos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridAdiantamentos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAdiantamentos, iAlterado)
    End If

End Sub

Private Sub GridAdiantamentos_GotFocus()
    Call Grid_Recebe_Foco(objGridAdiantamentos)
End Sub

Private Sub GridAdiantamentos_EnterCell()
    Call Grid_Entrada_Celula(objGridAdiantamentos, iAlterado)
End Sub

Private Sub GridAdiantamentos_LeaveCell()
    Call Saida_Celula(objGridAdiantamentos)
End Sub

Private Sub GridAdiantamentos_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridAdiantamentos)
End Sub

Private Sub GridAdiantamentos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridAdiantamentos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAdiantamentos, iAlterado)
    End If

End Sub

Private Sub GridAdiantamentos_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridAdiantamentos)
End Sub

Private Sub GridAdiantamentos_RowColChange()
    Call Grid_RowColChange(objGridAdiantamentos)
End Sub

Private Sub GridAdiantamentos_Scroll()
    Call Grid_Scroll(objGridAdiantamentos)
End Sub

Private Sub GridDebitos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridDebitos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDebitos, iAlterado)
    End If

End Sub

Private Sub GridDebitos_GotFocus()
    Call Grid_Recebe_Foco(objGridDebitos)
End Sub

Private Sub GridDebitos_EnterCell()
    Call Grid_Entrada_Celula(objGridDebitos, iAlterado)
End Sub

Private Sub GridDebitos_LeaveCell()
    Call Saida_Celula(objGridDebitos)
End Sub

Private Sub GridDebitos_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridDebitos)
End Sub

Private Sub GridDebitos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridDebitos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDebitos, iAlterado)
    End If

End Sub

Private Sub GridDebitos_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridDebitos)
End Sub

Private Sub GridDebitos_RowColChange()
    Call Grid_RowColChange(objGridDebitos)
End Sub

Private Sub GridDebitos_Scroll()
    Call Grid_Scroll(objGridDebitos)
End Sub

Private Sub AdiantamentoSeleciona_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAdiantamentos)
End Sub

Private Sub AdiantamentoSeleciona_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdiantamentos)
End Sub

Private Sub AdiantamentoSeleciona_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridAdiantamentos.objControle = AdiantamentoSeleciona
    lErro = Grid_Campo_Libera_Foco(objGridAdiantamentos)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub FilialAdiantamento_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAdiantamentos)
End Sub

Private Sub FilialAdiantamento_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdiantamentos)
End Sub

Private Sub FilialAdiantamento_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridAdiantamentos.objControle = FilialAdiantamento
    lErro = Grid_Campo_Libera_Foco(objGridAdiantamentos)
    If lErro <> SUCESSO Then Cancel = True
End Sub
    
Private Sub DataAdiantamento_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAdiantamentos)
End Sub

Private Sub DataAdiantamento_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdiantamentos)
End Sub

Private Sub DataAdiantamento_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridAdiantamentos.objControle = DataAdiantamento
    lErro = Grid_Campo_Libera_Foco(objGridAdiantamentos)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub CtaCorrenteAdiantamento_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAdiantamentos)
End Sub

Private Sub CtaCorrenteAdiantamento_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdiantamentos)
End Sub

Private Sub CtaCorrenteAdiantamento_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridAdiantamentos.objControle = CtaCorrenteAdiantamento
    lErro = Grid_Campo_Libera_Foco(objGridAdiantamentos)
    If lErro <> SUCESSO Then Cancel = True
End Sub
    
Private Sub MeioPagtoAdiantamento_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAdiantamentos)
End Sub

Private Sub MeioPagtoAdiantamento_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdiantamentos)
End Sub

Private Sub MeioPagtoAdiantamento_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridAdiantamentos.objControle = MeioPagtoAdiantamento
    lErro = Grid_Campo_Libera_Foco(objGridAdiantamentos)
    If lErro <> SUCESSO Then Cancel = True
End Sub
    
Private Sub ValorAdiantamento_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAdiantamentos)
End Sub

Private Sub ValorAdiantamento_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdiantamentos)
End Sub

Private Sub ValorAdiantamento_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridAdiantamentos.objControle = ValorAdiantamento
    lErro = Grid_Campo_Libera_Foco(objGridAdiantamentos)
    If lErro <> SUCESSO Then Cancel = True
End Sub
    
Private Sub SaldoAdiantamento_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAdiantamentos)
End Sub

Private Sub SaldoAdiantamento_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdiantamentos)
End Sub

Private Sub SaldoAdiantamento_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridAdiantamentos.objControle = SaldoAdiantamento
    lErro = Grid_Campo_Libera_Foco(objGridAdiantamentos)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub DebitoSeleciona_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridDebitos)
End Sub

Private Sub DebitoSeleciona_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDebitos)
End Sub

Private Sub DebitoSeleciona_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridDebitos.objControle = DebitoSeleciona
    lErro = Grid_Campo_Libera_Foco(objGridDebitos)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub FilialDebito_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridDebitos)
End Sub

Private Sub FilialDebito_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDebitos)
End Sub

Private Sub FilialDebito_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridDebitos.objControle = FilialDebito
    lErro = Grid_Campo_Libera_Foco(objGridDebitos)
    If lErro <> SUCESSO Then Cancel = True
End Sub
    
Private Sub DataDebito_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridDebitos)
End Sub

Private Sub DataDebito_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDebitos)
End Sub

Private Sub DataDebito_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridDebitos.objControle = DataDebito
    lErro = Grid_Campo_Libera_Foco(objGridDebitos)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub ValorDebito_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridDebitos)
End Sub

Private Sub ValorDebito_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDebitos)
End Sub

Private Sub ValorDebito_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridDebitos.objControle = ValorDebito
    lErro = Grid_Campo_Libera_Foco(objGridDebitos)
    If lErro <> SUCESSO Then Cancel = True
End Sub
    
Private Sub SaldoDebito_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridDebitos)
End Sub

Private Sub SaldoDebito_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDebitos)
End Sub

Private Sub SaldoDebito_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridDebitos.objControle = SaldoDebito
    lErro = Grid_Campo_Libera_Foco(objGridDebitos)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da ceélula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
        
        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 92789
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 94659

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 92789

        Case 94659
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143222)

    End Select

    Exit Function

End Function

Private Sub AdiantamentoSeleciona_Click()
 
Dim lErro As Long

    'Se o usuário selecionou a parcela
    If GridAdiantamentos.TextMatrix(GridAdiantamentos.Row, iGrid_AdiantamentoSeleciona_Col) = True Then
        
        'Guarda no obj a informação de que a parcela está selecionada
        gobjBaixaAntecipDebCliente.colAntecipDebCliente(GridAdiantamentos.Row).iSelecionado = MARCADO
    
    'Senão
    Else
        
        'Guarda no obj global a informação de que a parcela não está selecionada
        gobjBaixaAntecipDebCliente.colAntecipDebCliente(GridAdiantamentos.Row).iSelecionado = DESMARCADO
    
    End If
    
    'Calcula o valor total de adiantamentos a serem baixados
    Call Calcula_Total_Baixar(TIPOBAIXA_ADIANTAMENTO)
    
    'Atualiza os CheckBox na tela
    Call Grid_Refresh_Checkbox(objGridAdiantamentos)

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DebitoSeleciona_Click()
 
Dim lErro As Long

    'Se o usuário selecionou a parcela
    If GridDebitos.TextMatrix(GridDebitos.Row, iGrid_DebitoSeleciona_Col) = True Then
        
        'Guarda no obj a informação de que a parcela está selecionada
        gobjBaixaAntecipDebCliente.colAntecipDebCliente(GridDebitos.Row).iSelecionado = MARCADO
    
    'Senão
    Else
        
        'Guarda no obj global a informação de que a parcela não está selecionada
        gobjBaixaAntecipDebCliente.colAntecipDebCliente(GridDebitos.Row).iSelecionado = DESMARCADO
    
    End If
    
    Call Calcula_Total_Baixar(TIPOBAIXA_DEBITOS)

    'Atualiza os CheckBox na tela
    Call Grid_Refresh_Checkbox(objGridDebitos)
    
    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Calcula_Total_Baixar(iTipoBaixa As Integer)

Dim iIndice As Integer
Dim dTotalSelecionado As Double
Dim objGrid As AdmGrid
Dim ControleGrid As Object
Dim iColunaSeleciona As Integer
Dim iColunaValor As Integer

    dTotalSelecionado = 0
    
    'Verifica o tipo de baixa passado como parâmetro
    Select Case iTipoBaixa
    
        'Se for baixa de adiantamentos => o grid utilizado para o cálculo deve ser o de adiantamentos
        Case TIPOBAIXA_ADIANTAMENTO
        
            'Instancia o obj e controle com os seus correspondentes para o grid Adiantamentos
            Set objGrid = objGridAdiantamentos
            Set ControleGrid = GridAdiantamentos
            
            'Indica a coluna que deverá ser testada para verificar se a linha está selecionada
            iColunaSeleciona = iGrid_AdiantamentoSeleciona_Col
        
            'Indica a coluna que contém o valor que será acumulado
            iColunaValor = iGrid_SaldoAdiantamento_Col
        
        'Se for baixa de débitos => o grid utilizado para o cálculo deve ser o de débitos
        Case TIPOBAIXA_DEBITOS
            
            'Instancia o obj e controle com os seus correspondentes para o grid Débitos
            Set objGrid = objGridDebitos
            Set ControleGrid = GridDebitos
        
            'Indica a coluna que deverá ser testada para verificar se a linha está selecionada
            iColunaSeleciona = iGrid_DebitoSeleciona_Col
        
            'Indica a coluna que contém o valor que será acumulado
            iColunaValor = iGrid_SaldoDebito_Col
        
    End Select
    
    'Para cada linha do grid
    For iIndice = 1 To objGrid.iLinhasExistentes
        
        'Se a linha está selecionada
        If StrParaInt(ControleGrid.TextMatrix(iIndice, iColunaSeleciona)) <> 0 Then
            
            'Adiciona o valor da linha ao valor total das linhas selecionadas
            dTotalSelecionado = dTotalSelecionado + StrParaDbl(ControleGrid.TextMatrix(iIndice, iColunaValor))
        End If
        
    Next
    
    'Exibe na tela o valor encontrado acima
    LabelValorBaixa2.Caption = Format(dTotalSelecionado, "Standard")
    
    'Guarda no obj global o valor encontrado
    gobjBaixaAntecipDebCliente.dValorBaixa = dTotalSelecionado
    
    Exit Sub
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

Dim lErro As Long

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_UnLoad(Cancel As Integer)

    Set objEventoClientes = Nothing
    Set objEventoNumDebitoInic = Nothing
    Set objEventoNumDebitoFim = Nothing
    Set objGridAdiantamentos = Nothing
    Set objGridDebitos = Nothing
    Set gobjBaixaAntecipDebCliente = Nothing
    
    'Eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing
    
    Set objGrid1 = Nothing
    Set objContabil = Nothing
    
End Sub

Function Trata_Parametros() As Long
    Trata_Parametros = SUCESSO
End Function
Public Sub Form_Load()

Dim lErro As Long
Dim iSubTipo As Long

On Error GoTo Erro_Form_Load

    If giTipoVersao = VERSAO_LIGHT Then
        
        TabStrip1.Tabs.Remove (TAB_Contabilizacao)
    
    End If
    
    iFramePrincipalAtual = 1
    iFramePrincipalAlterado = 0

    FramePrincipal(1).Visible = True
    FramePrincipal(2).Visible = False
    FramePrincipal(3).Visible = False

    Call Inicializa_gobjBaixaAntecipDebCliente(TELA_TODA)
    
    'Inicializa os obj's associados aos grid's
    Set objGridAdiantamentos = New AdmGrid
    Set objGridDebitos = New AdmGrid

    'Inicializa os obj's associados a eventos
    Set objEventoClientes = New AdmEvento
    Set objEventoNumDebitoInic = New AdmEvento
    Set objEventoNumDebitoFim = New AdmEvento

    'Preenche o campo data da baixa com a data atual
    DataBaixa.PromptInclude = False
    DataBaixa.Text = Format(gdtDataHoje, "dd/mm/yy")
    DataBaixa.PromptInclude = True
    
    'Carrega list de ComboBox ContaCorrente utilizada para filtro
    lErro = Carrega_Combos_ContaCorrente(ContaCorrenteSeleciona)
    If lErro <> SUCESSO Then gError 94601

    'Carrega list de ComboBox ContaCorrente utilizada para selecionar onde será feito o movimento de baixa
    lErro = Carrega_Combos_ContaCorrente(ContaCorrente)
    If lErro <> SUCESSO Then gError 94626

    'Carrega a list de tipos de Débitos / Devoluções
    lErro = TipoDebitoSeleciona_Carrega(TipoDebitoSeleciona)
    If lErro <> SUCESSO Then gError 94602

    'Executa a Inicialização do grid Parcelas
    lErro = Inicializa_GridAdiantamentos(objGridAdiantamentos)
    If lErro <> SUCESSO Then gError 94603

    'Executa a Inicialização do grid Debitos
    lErro = Inicializa_GridDebitos(objGridDebitos)
    If lErro <> SUCESSO Then gError 94604

    'inicializacao da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_CONTASARECEBER)
    If lErro <> SUCESSO Then gError 92788

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 92788, 94601 To 94604, 94626

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143223)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Private Sub Inicializa_gobjBaixaAntecipDebCliente(Optional iTelaToda As Integer = 0)

    If iTelaToda = TELA_TODA Then
    
        'Inicializa o obj que serve de "espelho" à tela
        Set gobjBaixaAntecipDebCliente = New ClassBaixaAntecipDebCliente
    
        'Guarda no obj global a informação de que não existem cliente e filial selecionados
        gobjBaixaAntecipDebCliente.lCliente = 0
        gobjBaixaAntecipDebCliente.iFilialCliente = -1
    
        'Guarda no obj global a informação do tipo de baixa selecionado
        gobjBaixaAntecipDebCliente.iTipoBaixa = TIPOBAIXA_ADIANTAMENTO
    
        'Guarda na variável a informação do último de baixa que foi selecionado ao se clicar no tabstrip
        iUltTipoBaixa = TIPOBAIXA_DEBITOS
    
        'Guarda no obj global a informação que todas as contas correntes estão selecionadas
        gobjBaixaAntecipDebCliente.iCtaCorrenteTodas = CONTACORRENTE_TODAS
    
        'Guarda no obj global a informação que não existe uma conta específica selecionada
        gobjBaixaAntecipDebCliente.iCtaCorrenteSelecionada = COMBO_INDICE
    
        'Guarda no obj global a informação de que não foram informadas datas de emissão
        gobjBaixaAntecipDebCliente.dtDataEmissaoAte = DATA_NULA
        gobjBaixaAntecipDebCliente.dtDataEmissaoDe = DATA_NULA
    
        'Guarda no obj global a informação de que não foram informados números de débitos
        gobjBaixaAntecipDebCliente.lNumDebitoDe = 0
        gobjBaixaAntecipDebCliente.lNumDebitoAte = 0
        
        'Guarda no obj global a informação de que todos os tipos de débito estão selecionados
        gobjBaixaAntecipDebCliente.iTipoDebitoTodos = TIPODEBITO_TODOS
        
        'Guarda no obj global a informação que não existe um tipo específico selecionado
        gobjBaixaAntecipDebCliente.sTipoDebito = ""
    
    End If
    
    'Guarda no obj global a data da baixa
    gobjBaixaAntecipDebCliente.dtDataBaixa = gdtDataHoje
    
    'Guarda no obj global a informação de que o valor das parcelas selecionadas é zero
    gobjBaixaAntecipDebCliente.dValorBaixa = 0
    
    'Guarda no obj global a informação de que ainda não existe conta selecionada para o movimento de baixa
    gobjBaixaAntecipDebCliente.iCtaCorrenteBaixa = COMBO_INDICE
    
    'Guarda no obj global a informação de que não há histórico informado
    gobjBaixaAntecipDebCliente.sHistoricoBaixa = ""

End Sub

Private Function Carrega_Combos_ContaCorrente(objComboBox As ComboBox) As Long

Dim lErro As Long
Dim colCodigoNomeRed As New AdmColCodigoNome
Dim objCodigoNomeRed As AdmCodigoNome

On Error GoTo Erro_Carrega_Combos_ContaCorrente

    'Lê Codigos, NomesReduzidos de ContasCorrentes
    lErro = CF("ContasCorrentesInternas_Le_CodigosNomesRed", colCodigoNomeRed)
    If lErro <> SUCESSO Then gError 94599

    'Preeche list de ComboBox
    For Each objCodigoNomeRed In colCodigoNomeRed
        objComboBox.AddItem CStr(objCodigoNomeRed.iCodigo) & SEPARADOR & objCodigoNomeRed.sNome
        objComboBox.ItemData(objComboBox.NewIndex) = objCodigoNomeRed.iCodigo
    Next

    Carrega_Combos_ContaCorrente = SUCESSO

    Exit Function

Erro_Carrega_Combos_ContaCorrente:

    Carrega_Combos_ContaCorrente = gErr

    Select Case gErr

        Case 94599

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143224)

    End Select

    Exit Function

End Function

Private Function TipoDebitoSeleciona_Carrega(objComboBox As ComboBox) As Long

Dim lErro As Long
Dim colTipoDebito As New colTipoDocumento
Dim objTipoDebito As ClassTipoDocumento

On Error GoTo Erro_TipoDebitoSeleciona_Carrega

    'Lê os Tipos de Débito / Devoluções no BD
    lErro = CF("TiposDocumento_Le_DebReceber", colTipoDebito)
    If lErro <> SUCESSO Then gError 94600

    'Inclui na combo passada como parâmetro os tipos lidos do BD
    For Each objTipoDebito In colTipoDebito
        objComboBox.AddItem objTipoDebito.sSigla & SEPARADOR & objTipoDebito.sDescricaoReduzida
    Next

    TipoDebitoSeleciona_Carrega = SUCESSO

    Exit Function

Erro_TipoDebitoSeleciona_Carrega:

    TipoDebitoSeleciona_Carrega = gErr

    Select Case gErr

        Case 94600

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143225)

    End Select

    Exit Function

End Function

Private Function Inicializa_GridAdiantamentos(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Recebimentos Antecipados

    'tela em questão
    Set objGridAdiantamentos.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Filial Emp")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Conta Corrente")
    objGridInt.colColuna.Add ("Meio Pagto")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Saldo")

   'campos de edição do grid
    objGridInt.colCampo.Add (AdiantamentoSeleciona.Name)
    objGridInt.colCampo.Add (FilialAdiantamento.Name)
    objGridInt.colCampo.Add (DataAdiantamento.Name)
    objGridInt.colCampo.Add (CtaCorrenteAdiantamento.Name)
    objGridInt.colCampo.Add (MeioPagtoAdiantamento.Name)
    objGridInt.colCampo.Add (ValorAdiantamento.Name)
    objGridInt.colCampo.Add (SaldoAdiantamento.Name)

    iGrid_AdiantamentoSeleciona_Col = 1
    iGrid_FilialAdiantamento_Col = 2
    iGrid_DataAdiantamento_Col = 3
    iGrid_CtaCorrenteAdiantamento_Col = 4
    iGrid_MeioPagtoAdiantamento_Col = 5
    iGrid_ValorAdiantamento_Col = 6
    iGrid_SaldoAdiantamento_Col = 7

    objGridInt.objGrid = GridAdiantamentos

    'todas as linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 7

    'largura da primeira coluna
    GridAdiantamentos.ColWidth(0) = 0

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Não permite incluir novas linhas nem excluir as existentes
    objGridInt.iProibidoIncluir = 1
    objGridInt.iProibidoExcluir = 1

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridAdiantamentos = SUCESSO

    Exit Function

End Function

Private Function Inicializa_GridDebitos(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Debitos

    'tela em questão
    Set objGridDebitos.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Filial Emp")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Saldo")


    'campos de edição do grid
    objGridInt.colCampo.Add (DebitoSeleciona.Name)
    objGridInt.colCampo.Add (FilialDebito.Name)
    objGridInt.colCampo.Add (DataDebito.Name)
    objGridInt.colCampo.Add (TipoDebito.Name)
    objGridInt.colCampo.Add (NumDebito.Name)
    objGridInt.colCampo.Add (ValorDebito.Name)
    objGridInt.colCampo.Add (SaldoDebito.Name)

    iGrid_DebitoSeleciona_Col = 1
    iGrid_FilialDebito_Col = 2
    iGrid_DataDebito_Col = 3
    iGrid_TipoDebito_Col = 4
    iGrid_NumDebito_Col = 5
    iGrid_ValorDebito_Col = 6
    iGrid_SaldoDebito_Col = 7

    objGridInt.objGrid = GridDebitos

    'todas as linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 7

    'largura da primeira coluna
    GridDebitos.ColWidth(0) = 0

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Não permite incluir novas linhas nem excluir as existentes
    objGridInt.iProibidoIncluir = 1
    objGridInt.iProibidoExcluir = 1

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridDebitos = SUCESSO

    Exit Function

End Function

Private Sub TabStrip1_BeforeClick(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_TabStrip1_BeforeClick

    'Se estava no tab de filtros e está passando para outro tab
    If iFramePrincipalAtual = TAB_FILTROS Then
    
        'Valida os filtros
        lErro = Valida_Filtros()
        If lErro <> SUCESSO Then gError 94627
    
    End If

    Exit Sub

Erro_TabStrip1_BeforeClick:

    Cancel = True

    Select Case gErr

        Case 94627

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143226)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click

    'Se Frame atual não corresponde ao Tab clicado
    If TabStrip1.SelectedItem.Index <> iFramePrincipalAtual Then

        If TabStrip_PodeTrocarTab(iFramePrincipalAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna o frame clicado visível
        FramePrincipal(TabStrip1.SelectedItem.Index).Visible = True

        'Torna o frame atual invisível
        FramePrincipal(iFramePrincipalAtual).Visible = False

        'Armazena novo valor de iFrameAtual
        iFramePrincipalAtual = TabStrip1.SelectedItem.Index

        'Se Frame selecionado foi o de Títulos
        If TabStrip1.SelectedItem.Index = TAB_FILTROS Then

            iFramePrincipalAlterado = 0

        'Se Frame selecionado foi o de Parcelas
        ElseIf TabStrip1.SelectedItem.Index = TAB_ADIANTAMENTOSDEBITOS And iUltTipoBaixa <> gobjBaixaAntecipDebCliente.iTipoBaixa And iFramePrincipalAlterado <> 0 Then

'            'Exibe e oculta os controles específicos para o tipo de baixa selecionado
'            Select Case gobjBaixaAntecipDebCliente.iTipoBaixa
'
'                Case TIPOBAIXA_ADIANTAMENTO
'
'                    'Torna visível o frame que contém o grid que exibirá os adiantamentos
'                    FrameAdiantamentos.Visible = True
'
'                    'Torna invisível o frame que contém o grid que exibirá os débitos / devoluções
'                    FrameDebitos.Visible = False
'
'                    'Exibe os controles que são exclusivos para a baixa de adiantamentos
'                    ContaCorrente.Visible = True
'                    LabelCtaCorrente.Visible = True
'
'                    'Reposiciona os controles na tela
'                    LabelHistorico.Top = 1235
'                    Historico.Top = 1175
'
'                Case TIPOBAIXA_DEBITOS
'
'                    'Torna visível o frame que contém o grid que exibirá os débitos / devoluções
'                    FrameDebitos.Visible = True
'
'                    'Torna invisível o frame que contém o grid que exibirá os adiantamentos
'                    FrameAdiantamentos.Visible = False
'
'                    'Oculta os controles que são exclusivos para a baixa de adiantamentos
'                    ContaCorrente.Visible = False
'                    LabelCtaCorrente.Visible = False
'
'                    'Reposiciona os controles na tela
'                    LabelHistorico.Top = 975
'                    Historico.Top = 925
'
'            End Select

            lErro = Carrega_Tab_AdiantamentosDebitos()
            If lErro <> SUCESSO Then gError 94605

            iFramePrincipalAlterado = 0

        End If

        'se estiver selecionando o tabstrip de contabilidade e o usuário não alterou a contabilidade ==> carrega o modelo padrao
        If TabStrip1.SelectedItem.Caption = TITULO_TAB_CONTABILIDADE Then Call objContabil.Contabil_Carga_Modelo_Padrao

        Select Case iFramePrincipalAtual

            Case TAB_FILTROS
                '??? criar idh Parent.HelpContextID = IDH_BAIXA_PARCELAS_RECEBER_TITULOS

            Case TAB_ADIANTAMENTOSDEBITOS
                '??? criar idh Parent.HelpContextID = IDH_BAIXA_PARCELAS_RECEBER_PARCELAS

'            Case TAB_Contabilizacao
'                '??? criar idh Parent.HelpContextID = IDH_BAIXA_PARCELAS_RECEBER_CONTABILIZACAO

        End Select

    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case 94605

        Case 94627

            'Call TabStrip1_Click

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143227)

    End Select

    Exit Sub

End Sub

Function Carrega_Tab_AdiantamentosDebitos() As Long

Dim lErro As Long

On Error GoTo Erro_Carrega_Tab_AdiantamentosDebitos

    Set gobjBaixaAntecipDebCliente.colAntecipDebCliente = Nothing
    
    Call Limpa_Tab_AdiantamentosDebitos
    
    Select Case gobjBaixaAntecipDebCliente.iTipoBaixa

        Case TIPOBAIXA_ADIANTAMENTO
        
            lErro = CF("BaixaAntecipDebCliente_Le_Adiantamentos", gobjBaixaAntecipDebCliente)
            If lErro <> SUCESSO And lErro <> 94664 Then gError 94628
            
            If lErro = 94664 Then gError 94660
            
            'Torna o Frame com o grid de Adiantamentos visível
            FrameAdiantamentos.Visible = True
            
            'Torna o Frame com o grid de Débitos invisível
            FrameDebitos.Visible = False
            
            Call Traz_Adiantamentos_Tela
            
        Case TIPOBAIXA_DEBITOS

            lErro = CF("BaixaAntecipDebCliente_Le_Debitos", gobjBaixaAntecipDebCliente)
            If lErro <> SUCESSO And lErro <> 94675 Then gError 94707
            
            If lErro = 94675 Then gError 94708
            
            'Torna o Frame com o grid de Débitos visível
            FrameDebitos.Visible = True
            
            'Torna o Frame com o grid de Adiantamentos invisível
            FrameAdiantamentos.Visible = False
            
            Call Traz_Debitos_Tela

    End Select

    Carrega_Tab_AdiantamentosDebitos = SUCESSO
    
    Exit Function

Erro_Carrega_Tab_AdiantamentosDebitos:

    Carrega_Tab_AdiantamentosDebitos = gErr

    Select Case gErr

        Case 94707, 94628
        
        Case 94660
            Call Rotina_Erro(vbOKOnly, "ERRO_ANTECIPREC_INEXISTENTE2", gErr)
            
        Case 94708
            Call Rotina_Erro(vbOKOnly, "ERRO_DEBITORECCLI_INEXISTENTE", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143228)
            
    End Select

    Exit Function

End Function

Private Sub TipoBaixaAdiantamento_Click()
    'Torna visível o frame que contém filtros exclusivos para baixar adiantamentos
    FrameFiltrosAdiantamento.Visible = True

    'Torna invisível o frame que contém filtros exclusivos para baixar débitos
    FrameFiltrosDebito.Visible = False

    'Guarda a informação do tipo de baixa que estava selecionado
    iUltTipoBaixa = gobjBaixaAntecipDebCliente.iTipoBaixa

    'Guarda a informação do tipo de baixa que está selecionado
    gobjBaixaAntecipDebCliente.iTipoBaixa = TIPOBAIXA_ADIANTAMENTO

    'Registra uma alteração na tela
    iAlterado = REGISTRO_ALTERADO

    'Registra uma alteração no frame principal da tela
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoBaixaDebitos_Click()
    'Torna visível o frame que contém filtros exclusivos para baixar débitos
    FrameFiltrosDebito.Visible = True

    'Torna invisível o frame que contém filtros exclusivos para baixar adiantamentos
    FrameFiltrosAdiantamento.Visible = False

    'Guarda a informação do tipo de baixa que estava selecionado
    iUltTipoBaixa = gobjBaixaAntecipDebCliente.iTipoBaixa

    'Guarda a informação do tipo de baixa que está selecionado
    gobjBaixaAntecipDebCliente.iTipoBaixa = TIPOBAIXA_DEBITOS

    'Registra uma alteração na tela
    iAlterado = REGISTRO_ALTERADO

    'Registra uma alteração no frame principal da tela
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub EmissaoInic_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub EmissaoInic_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_EmissaoInic_Validate

    'Se a data EmissaoInic está preenchida
    If Len(EmissaoInic.ClipText) > 0 Then

        'Verifica se a data EmissaoInic é válida
        lErro = Data_Critica(EmissaoInic.Text)
        If lErro <> SUCESSO Then gError 94622

        'Guarda no obj a data De que deve ser passada como parâmetro
        gobjBaixaAntecipDebCliente.dtDataEmissaoDe = StrParaDate(EmissaoInic.Text)

    Else

        'Guarda no obj a data De que deve ser passada como parâmetro
        gobjBaixaAntecipDebCliente.dtDataEmissaoDe = DATA_NULA

    End If

    Exit Sub

Erro_EmissaoInic_Validate:

    Cancel = True

    Select Case gErr

        Case 94622

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143229)

    End Select

    Exit Sub

End Sub

Private Sub EmissaoFim_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub EmissaoFim_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_EmissaoFim_Validate

    'Se a data EmissaoInic está preenchida
    If Len(EmissaoFim.ClipText) > 0 Then

        'Verifica se a data EmissaoInic é válida
        lErro = Data_Critica(EmissaoFim.Text)
        If lErro <> SUCESSO Then gError 94623

        'Guarda no obj a data Até que deve ser passada como parâmetro
        gobjBaixaAntecipDebCliente.dtDataEmissaoAte = StrParaDate(EmissaoFim.Text)

    Else

        'Guarda no obj a data Até que deve ser passada como parâmetro
        gobjBaixaAntecipDebCliente.dtDataEmissaoAte = DATA_NULA

    End If


    Exit Sub

Erro_EmissaoFim_Validate:

    Cancel = True

    Select Case gErr

        Case 94623

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143230)

    End Select

    Exit Sub

End Sub

Private Sub NumDebitoInic_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO

    'Guarda no obj global a tela o novo código de filtro para número de débito
    gobjBaixaAntecipDebCliente.lNumDebitoDe = StrParaLong(NumDebitoInic.Text)

End Sub

Private Sub NumDebitoFim_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO

    'Guarda no obj global a tela o novo código de filtro para número de débito
    gobjBaixaAntecipDebCliente.lNumDebitoAte = StrParaLong(NumDebitoFim.Text)

End Sub

Private Sub CtaCorrenteApenas_Click()

    'Habilita a combo para a seleção da conta corrente
    ContaCorrenteSeleciona.Enabled = True

    'Guarda no obj a informação que apenas uma conta será selecionada para filtro
    gobjBaixaAntecipDebCliente.iCtaCorrenteTodas = CONTACORRENTE_APENAS

    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CtaCorrenteTodas_Click()

    'Desabilita a combo para a seleção da conta corrente
    ContaCorrenteSeleciona.Enabled = False

    'Guarda no obj a informação que todas as contas estão selecionadas
    gobjBaixaAntecipDebCliente.iCtaCorrenteTodas = CONTACORRENTE_TODAS

    'Guarda no obj a informação de que não há uma conta corrente específica selecionada
    gobjBaixaAntecipDebCliente.iCtaCorrenteSelecionada = COMBO_INDICE

    'Limpa a combo de seleção de conta corrente
    ContaCorrenteSeleciona.ListIndex = COMBO_INDICE

    iFramePrincipalAlterado = REGISTRO_ALTERADO

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaCorrenteSeleciona_Change()

    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaCorrenteSeleciona_Validate(Cancel As Boolean)

Dim sContaCorrenteSeleciona As String
Dim iCodigo As Integer
Dim objContaCorrente As New ClassContasCorrentesInternas
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ContaCorrenteSeleciona_Validate

    'Indica que não existe conta corrente selecionada (só existirá uma conta selecionada depois da execução do validate)
    gobjBaixaAntecipDebCliente.iCtaCorrenteSelecionada = COMBO_INDICE

    'Se a ComboBox ContaCorrenteSeleciona não está preenchida => sai da função
    If Len(Trim(ContaCorrenteSeleciona.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox ContaCorrenteSeleciona
    If ContaCorrenteSeleciona.Text = ContaCorrenteSeleciona.List(ContaCorrenteSeleciona.ListIndex) Then
        gobjBaixaAntecipDebCliente.iCtaCorrenteSelecionada = ContaCorrenteSeleciona.ItemData(ContaCorrenteSeleciona.ListIndex)
        Exit Sub
    End If

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(ContaCorrenteSeleciona, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 94618

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        'Lê os dados da conta corrente a partir do código passado como parâmetro
        lErro = CF("ContaCorrenteInt_Le", iCodigo, objContaCorrente)
        If lErro <> SUCESSO And lErro <> 11807 Then gError 94619

        'Se não encontrou a conta => o número de erro recebe tratamento para perguntar se deseja criar a conta
        If lErro = 11807 Then gError 94620

        'Coloca os dados da conta na combo
        ContaCorrenteSeleciona.Text = CStr(objContaCorrente.iCodigo) & SEPARADOR & objContaCorrente.sNomeReduzido

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then

        'Lê os dados da conta corrente a partir do nome red. passado como parâmetro
        objContaCorrente.sNomeReduzido = ContaCorrente.Text
        lErro = CF("ContaCorrenteInt_Le_NomeReduzido", objContaCorrente)
        If lErro <> SUCESSO And lErro <> 15981 Then gError 94747

        'Se não encontrou a conta => Erro
        If lErro = 15981 Then gError 94621
        
        'Coloca os dados da conta na combo
        ContaCorrenteSeleciona.Text = CStr(objContaCorrente.iCodigo) & SEPARADOR & objContaCorrente.sNomeReduzido
    
    End If

    'Guarda no obj global a tela o código da conta corrente selecionada
    gobjBaixaAntecipDebCliente.iCtaCorrenteSelecionada = objContaCorrente.iCodigo

    Exit Sub

Erro_ContaCorrenteSeleciona_Validate:

    Cancel = True

    Select Case gErr

        Case 94618, 94619

        Case 94620
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONTACORRENTE", ContaCorrenteSeleciona.Text)

            If vbMsgRes = vbYes Then
                'Chama a tela de Filiais
                Call Chama_Tela("CtaCorrenteInt", objContaCorrente)
            Else
                'Segura o foco
            End If

        Case 94621
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", gErr, objContaCorrente.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143231)

    End Select

    Exit Sub

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate

    If iClienteAlterado = 0 Then Exit Sub

    'Se Cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código ou CPF ou CGC)
        lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
        If lErro <> SUCESSO Then gError 94606

        'Lê coleção de códigos, nomes de Filiais do Cliente
        lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
        If lErro <> SUCESSO Then gError 94607

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", Filial, iCodFilial)

    'Se Cliente não está preenchido
    ElseIf Len(Trim(Cliente.Text)) = 0 Then

        'Limpa a Combo de Filiais
        Filial.Clear

    End If

    'Guarda no obj global a tela o codigo atual do Cliente e da Filial
    gobjBaixaAntecipDebCliente.lCliente = objCliente.lCodigo
    gobjBaixaAntecipDebCliente.iFilialCliente = iCodFilial

    iClienteAlterado = 0

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 94606, 94607

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143232)

    End Select

    Exit Sub

End Sub

Private Sub LabelCliente_Click()
'Chamada do Browse de Clientes

Dim colSelecao As Collection
Dim objCliente As New ClassCliente

    'Passa o Cliente que está na tela para o Obj
    objCliente.sNomeReduzido = Trim(Cliente.Text)

    'Chama a tela com a lista de Clientes
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClientes)

    Exit Sub

End Sub

Private Sub objEventoClientes_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Me.Show

    'Preenche Cliente na tela com NomeReduzido
    Set objCliente = obj1

    Cliente.Text = CStr(objCliente.sNomeReduzido)

    'Chama Validate de Cliente
    Call Cliente_Validate(bCancel)

    Exit Sub

End Sub

Private Sub Cliente_Change()
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
    iClienteAlterado = REGISTRO_ALTERADO
    
    Call Cliente_Preenche
    
End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objFilialCliente As New ClassFilialCliente
Dim iCodigo As Integer
Dim sNomeRed As String

On Error GoTo Erro_Filial_Validate

    'Indica que não foi selecionada nenhuma filial
    gobjBaixaAntecipDebCliente.iFilialCliente = COMBO_INDICE

    'Verifica se foi preenchida a ComboBox Filial
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Filial
    If Filial.Text = Filial.List(Filial.ListIndex) Then
        gobjBaixaAntecipDebCliente.iFilialCliente = Filial.ItemData(Filial.ListIndex)
        Exit Sub
    End If

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 94608

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        'Verifica se foi preenchido o Cliente
        If Len(Trim(Cliente.Text)) = 0 Then gError 94609

        'Lê o Cliente que está na tela
        sNomeRed = Trim(Cliente.Text)

        'Passa o Código da Filial que está na tela para o Obj
        objFilialCliente.iCodFilial = iCodigo

        'Lê Filial no BD a partir do NomeReduzido do Cliente e Código da Filial
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sNomeRed, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 94610

        'Se não existe a Filial
        If lErro = 17660 Then gError 94611

        'Encontrou Filial no BD, coloca no Text da Combo
        Filial.Text = CStr(objFilialCliente.iCodFilial) & SEPARADOR & objFilialCliente.sNome

    ElseIf lErro = 6731 Then gError 94612

        '??? alterar codigo para ler pelo nome reduzido

    End If

    'Guarda no obj global a tela o código da filial selecionada
    gobjBaixaAntecipDebCliente.iFilialCliente = objFilialCliente.iCodFilial

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 94608, 94610

        Case 94609
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 94611
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE1", Filial.Text)

            If vbMsgRes = vbYes Then
                'Chama a tela de Filiais
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            Else
                'Segura o foco
            End If

        Case 94612
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143233)

    End Select

    Exit Sub

End Sub

Private Sub Filial_Click()
    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Filial_Change()
    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO
End Sub

Private Sub LabelNumDebitoDe_Click()

Dim objDebitoRecCli As New ClassDebitoRecCli
Dim colSelecao As New Collection

On Error GoTo Erro_LabelNumDebitoDe_Click

    objDebitoRecCli.lCliente = gobjBaixaAntecipDebCliente.lCliente
    objDebitoRecCli.iFilial = gobjBaixaAntecipDebCliente.iFilialCliente

    Call Chama_Tela("DebitosRecebLista_Consulta", colSelecao, objDebitoRecCli, objEventoNumDebitoInic)

    Exit Sub

Erro_LabelNumDebitoDe_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143234)

    End Select

    Exit Sub

End Sub

Private Sub LabelNumDebitoAte_Click()

Dim objDebitoRecCli As New ClassDebitoRecCli
Dim colSelecao As New Collection

On Error GoTo Erro_LabelNumDebitoAte_Click

    objDebitoRecCli.lCliente = gobjBaixaAntecipDebCliente.lCliente
    objDebitoRecCli.iFilial = gobjBaixaAntecipDebCliente.iFilialCliente

    Call Chama_Tela("DebitosRecebLista_Consulta", colSelecao, objDebitoRecCli, objEventoNumDebitoFim)

    Exit Sub

Erro_LabelNumDebitoAte_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143235)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNumDebitoInic_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objDebitoRecCli As ClassDebitoRecCli

On Error GoTo Erro_objEventoNumDebitoInic_evSelecao

    Me.Show

    Set objDebitoRecCli = obj1

    NumDebitoInic.PromptInclude = False
    NumDebitoInic.Text = objDebitoRecCli.lNumTitulo
    NumDebitoInic.PromptInclude = True

    Exit Sub

Erro_objEventoNumDebitoInic_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143236)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNumDebitoFim_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objDebitoRecCli As ClassDebitoRecCli

On Error GoTo Erro_objEventoNumDebitoFim_evSelecao

    Me.Show

    Set objDebitoRecCli = obj1

    NumDebitoFim.PromptInclude = False
    NumDebitoFim.Text = objDebitoRecCli.lNumTitulo
    NumDebitoFim.PromptInclude = True

    Exit Sub

Erro_objEventoNumDebitoFim_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143237)

    End Select

    Exit Sub

End Sub

Private Sub TipoDebitoTodos_Click()

    'Habilita a combo para a seleção do tipo de débito
    TipoDebitoSeleciona.Enabled = False

    'Guarda no obj a informação que todos os tipos de débito serão selecionados para filtro
    gobjBaixaAntecipDebCliente.iTipoDebitoTodos = TIPODEBITO_TODOS

    'Guarda no obj a informação de que não há um tipo de débito específico selecionado
    gobjBaixaAntecipDebCliente.sTipoDebito = ""
    
    'Limpa a combo de seleção de tipo de débito
    TipoDebitoSeleciona.ListIndex = COMBO_INDICE
    
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoDebitoApenas_Click()

    'Desabilita a combo para a seleção do tipo de débito
    TipoDebitoSeleciona.Enabled = True

    'Guarda no obj a informação que todas as contas estão selecionadas
    gobjBaixaAntecipDebCliente.iTipoDebitoTodos = TIPODEBITO_APENAS

    iFramePrincipalAlterado = REGISTRO_ALTERADO

    iAlterado = REGISTRO_ALTERADO

End Sub


Private Sub TipoDebitoSeleciona_Click()

Dim sTipoDebito As String
Dim lErro As Long

On Error GoTo Erro_TipoDebitoSeleciona_Click

    'Guarda na variável o conteúdo que será passado para a função
    'que obterá a sigla do Tipo
    sTipoDebito = TipoDebitoSeleciona.Text

    'Se foi selecionado um tipo
    If Len(Trim(sTipoDebito)) > 0 Then

        'Extrai a sigla do Tipo de Débito selecionado
         sTipoDebito = SCodigo_Extrai(sTipoDebito)

        'Guarda no obj a sigla do tipo de débito selecionado
        gobjBaixaAntecipDebCliente.sTipoDebito = sTipoDebito

    'Senão
    Else

        'Indica que nenhum tipo de débito foi selecionado
        gobjBaixaAntecipDebCliente.sTipoDebito = " "

    End If

    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_TipoDebitoSeleciona_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143238)

    End Select

End Sub

Private Sub UpDownDataBaixa_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataBaixa_DownClick

    'Diminui a DataBaixa em 1 dia
    lErro = Data_Up_Down_Click(DataBaixa, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 94662

    Exit Sub

Erro_UpDownDataBaixa_DownClick:

    Select Case gErr

        Case 94662

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143239)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataBaixa_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataBaixa_UpClick

    'Aumenta a DataBaixa em 1 dia
    lErro = Data_Up_Down_Click(DataBaixa, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 94663

    Exit Sub

Erro_UpDownDataBaixa_UpClick:

    Select Case gErr

        Case 94663

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143240)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoInic_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoInic_DownClick

    'Diminui a EmissaoInic em 1 dia
    lErro = Data_Up_Down_Click(EmissaoInic, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 94666

    Exit Sub

Erro_UpDownEmissaoInic_DownClick:

    Select Case gErr

        Case 94666

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143241)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoInic_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoInic_UpClick

    'Aumenta a EmissaoInic em 1 dia
    lErro = Data_Up_Down_Click(EmissaoInic, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 94667

    Exit Sub

Erro_UpDownEmissaoInic_UpClick:

    Select Case gErr

        Case 94667

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143242)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoFim_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoFim_DownClick

    'Diminui a EmissaoFim em 1 dia
    lErro = Data_Up_Down_Click(EmissaoFim, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 94668

    Exit Sub

Erro_UpDownEmissaoFim_DownClick:

    Select Case gErr

        Case 94668

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143243)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoFim_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoFim_UpClick

    'Aumenta a EmissaoFim em 1 dia
    lErro = Data_Up_Down_Click(EmissaoFim, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 94669

    Exit Sub

Erro_UpDownEmissaoFim_UpClick:

    Select Case gErr

        Case 94669

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143244)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    'Fecha a tela
    Unload Me
End Sub

Private Sub ContaCorrente_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ContaCorrente_Validate(Cancel As Boolean)

Dim sContaCorrenteSeleciona As String
Dim iCodigo As Integer
Dim objContaCorrente As New ClassContasCorrentesInternas
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ContaCorrente_Validate

    'Indica que não existe conta corrente selecionada (só existirá uma conta selecionada depois da execução do validate)
    gobjBaixaAntecipDebCliente.iCtaCorrenteBaixa = COMBO_INDICE

    'Se a ComboBox ContaCorrente não está preenchida => sai da função
    If Len(Trim(ContaCorrente.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox ContaCorrente
    If ContaCorrente.Text = ContaCorrente.List(ContaCorrente.ListIndex) Then
        gobjBaixaAntecipDebCliente.iCtaCorrenteBaixa = ContaCorrente.ItemData(ContaCorrente.ListIndex)
        Exit Sub
    End If

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(ContaCorrente, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 94744

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        'Lê os dados da conta corrente a partir do código passado como parâmetro
        objContaCorrente = iCodigo
        lErro = CF("ContaCorrenteInt_Le", iCodigo, objContaCorrente)
        If lErro <> SUCESSO And lErro <> 11807 Then gError 94745

        'Se não encontrou a conta => o número de erro recebe tratamento para perguntar se deseja criar a conta
        If lErro = 11807 Then gError 94746

        'Coloca os dados da conta na combo
        ContaCorrente.Text = CStr(objContaCorrente.iCodigo) & SEPARADOR & objContaCorrente.sNomeReduzido

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then
    
        'Lê os dados da conta corrente a partir do nome red. passado como parâmetro
        objContaCorrente.sNomeReduzido = ContaCorrente.Text
        lErro = CF("ContaCorrenteInt_Le_NomeReduzido", objContaCorrente)
        If lErro <> SUCESSO And lErro <> 15981 Then gError 94747

        'Se não encontrou a conta => Erro
        If lErro = 15981 Then gError 94748
        
        'Coloca os dados da conta na combo
        ContaCorrente.Text = CStr(objContaCorrente.iCodigo) & SEPARADOR & objContaCorrente.sNomeReduzido
        
    End If
        

    'Guarda no obj global a tela o código da conta corrente selecionada
    gobjBaixaAntecipDebCliente.iCtaCorrenteBaixa = objContaCorrente.iCodigo

    Exit Sub

Erro_ContaCorrente_Validate:

    Cancel = True

    Select Case gErr

        Case 94744, 94745, 94747

        Case 94746
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONTACORRENTE", ContaCorrenteSeleciona.Text)

            If vbMsgRes = vbYes Then
                'Chama a tela de Filiais
                Call Chama_Tela("CtaCorrenteInt", objContaCorrente)
            Else
                'Segura o foco
            End If

        Case 94748
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", gErr, objContaCorrente.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143245)

    End Select

    Exit Sub

End Sub

Private Sub DataBaixa_Change()
    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataBaixa_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataBaixa_Validate

    'Se a data DataBaixa está preenchida
    If Len(DataBaixa.ClipText) > 0 Then

        'Verifica se a data EmissaoInic é válida
        lErro = Data_Critica(DataBaixa.Text)
        If lErro <> SUCESSO Then gError 94749

        'Guarda no obj a data De que deve ser passada como parâmetro
        gobjBaixaAntecipDebCliente.dtDataBaixa = StrParaDate(DataBaixa.Text)

    Else

        'Guarda no obj a data De que deve ser passada como parâmetro
        gobjBaixaAntecipDebCliente.dtDataBaixa = DATA_NULA

    End If

    Exit Sub

Erro_DataBaixa_Validate:

    Cancel = True

    Select Case gErr

        Case 94749

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143246)

    End Select

    Exit Sub

End Sub

Private Sub Historico_Change()
    gobjBaixaAntecipDebCliente.sHistoricoBaixa = Historico.Text
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Function Valida_Filtros() As Long

On Error GoTo Erro_Valida_Filtros

    'Se o cliente não foi preenchido => erro
    If gobjBaixaAntecipDebCliente.lCliente = 0 Then gError 94613

    'Se a filial do cliente não foi preenchida = > erro
    If gobjBaixaAntecipDebCliente.iFilialCliente = 0 Then gError 94614

    'Se as datas de emissão (De e Até) foram preenchidas => verifica se data Até não é menor do que a data De
    If gobjBaixaAntecipDebCliente.dtDataEmissaoAte <> DATA_NULA And gobjBaixaAntecipDebCliente.dtDataEmissaoDe <> DATA_NULA Then

        'Se a data Até for menor do que a Data De => erro
        If gobjBaixaAntecipDebCliente.dtDataEmissaoAte < gobjBaixaAntecipDebCliente.dtDataEmissaoDe Then gError 94616

    End If

    Select Case gobjBaixaAntecipDebCliente.iTipoBaixa

        Case TIPOBAIXA_ADIANTAMENTO

            'Se o filtro de seleção por apenas uma conta corrente foi selecionado => Verifica se foi selecionada uma conta
            If gobjBaixaAntecipDebCliente.iCtaCorrenteTodas = CONTACORRENTE_APENAS Then

                'Se a Conta Corrente que será usada para filtro não foi informado => erro
                If gobjBaixaAntecipDebCliente.iCtaCorrenteSelecionada = COMBO_INDICE Then gError 94617

            End If

        Case TIPOBAIXA_DEBITOS

        'Se os números de débitos(De e Até) foram preenchidos => verifica se o número Até não é menor do que o número De
        If gobjBaixaAntecipDebCliente.lNumDebitoDe <> 0 And gobjBaixaAntecipDebCliente.lNumDebitoAte <> 0 Then

            'Se a data Até for menor do que a Data De => erro
            If gobjBaixaAntecipDebCliente.lNumDebitoAte < gobjBaixaAntecipDebCliente.lNumDebitoDe Then gError 94618

        End If
        
        'Se o filtro de seleção por apenas um tipo de débito foi selecionado => verifica se foi selecionado um tipo
        If gobjBaixaAntecipDebCliente.iTipoDebitoTodos = TIPODEBITO_APENAS Then
        
            'Se o tipo débito que será usado para filtro não foi informado => erro
            If Len(Trim(gobjBaixaAntecipDebCliente.sTipoDebito)) = 0 Then gError 94670
        
        End If

    End Select

    Valida_Filtros = SUCESSO

    Exit Function

Erro_Valida_Filtros:

    Valida_Filtros = gErr

    Select Case gErr

        Case 94613
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 94614
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_INFORMADA", gErr)

        Case 94616
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_EMISSAO_INICIAL_MAIOR", gErr)

        Case 94617
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_INFORMADA", gErr)

        Case 94618
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_ATE_MENOR_NUMERO_DE", gErr)
        
        Case 94670
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODEBITO_NAO_SELECIONADO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143247)

    End Select

    Exit Function

End Function

Sub Traz_Adiantamentos_Tela()

Dim objAntecipDebCliente As New ClassInfoBaixaAntecipDebCli
Dim iLinha As Integer

On Error GoTo Erro_Traz_Adiantamentos_Tela

    If gobjBaixaAntecipDebCliente.colAntecipDebCliente.Count >= objGridAdiantamentos.objGrid.Rows Then
        Call Refaz_Grid(objGridAdiantamentos, gobjBaixaAntecipDebCliente.colAntecipDebCliente.Count)
    End If
    
    For Each objAntecipDebCliente In gobjBaixaAntecipDebCliente.colAntecipDebCliente
    
        'Incrementa a variável de controle da linha atual
        iLinha = iLinha + 1
        
        With objAntecipDebCliente
        
            'Exibe a filial do adiantamento no grid
            GridAdiantamentos.TextMatrix(iLinha, iGrid_FilialAdiantamento_Col) = .iFilial
            
            'Exibe a data do adiantamento
            GridAdiantamentos.TextMatrix(iLinha, iGrid_DataAdiantamento_Col) = Format(.dtDataEmissao, "dd/mm/yy")
            
            'Exibe a conta corrente do adiantamento
            GridAdiantamentos.TextMatrix(iLinha, iGrid_CtaCorrenteAdiantamento_Col) = .iCodConta & " - " & .sNomeReduzidoConta
            
            'Verifica o tipo meio pagto
            Select Case .iTipoMeioPagto
                
                'Se for dinheiro
                Case DINHEIRO
                    
                    'Exibe a expressão "Dinheiro"
                    GridAdiantamentos.TextMatrix(iLinha, iGrid_MeioPagtoAdiantamento_Col) = "Dinheiro"
                
                'Se for cheque
                Case Cheque
                    
                    'Exibe a expressão "Cheque"
                    GridAdiantamentos.TextMatrix(iLinha, iGrid_MeioPagtoAdiantamento_Col) = "Cheque"
                    
                'Se for borderô
                Case BORDERO
                
                    'Exibe a expressão "Borderô"
                    GridAdiantamentos.TextMatrix(iLinha, iGrid_MeioPagtoAdiantamento_Col) = "Borderô"
            
            End Select
            
            'Exibe o valor total do adiantamento no formato moeda
            GridAdiantamentos.TextMatrix(iLinha, iGrid_ValorAdiantamento_Col) = Format(.dValor, "Standard")
            
            'Exibe o saldo ainda não utilizado do adiantamento no formato moeda
            GridAdiantamentos.TextMatrix(iLinha, iGrid_SaldoAdiantamento_Col) = Format(.dSaldoNaoApropriado, "Standard")
        
        End With
    
    Next
        
    'Passa para o Obj o número de Parcelas passadas pela Coleção
    objGridAdiantamentos.iLinhasExistentes = gobjBaixaAntecipDebCliente.colAntecipDebCliente.Count

    'Atualiza as checkboxes
    Call Grid_Refresh_Checkbox(objGridAdiantamentos)
    
    Exit Sub

Erro_Traz_Adiantamentos_Tela:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143248)
        
    End Select
    
    Exit Sub
    
End Sub

Sub Traz_Debitos_Tela()

Dim objAntecipDebCliente As New ClassInfoBaixaAntecipDebCli
Dim iLinha As Integer

On Error GoTo Erro_Traz_Debitos_Tela
    
    For Each objAntecipDebCliente In gobjBaixaAntecipDebCliente.colAntecipDebCliente
    
        'Incrementa a variável de controle da linha atual
        iLinha = iLinha + 1
        
       With objAntecipDebCliente
        
            'Exibe a filial do débito no grid
            GridDebitos.TextMatrix(iLinha, iGrid_FilialDebito_Col) = .iFilial
            
            If .dtDataEmissao <> DATA_NULA Then

                'Exibe a data do débito
                GridDebitos.TextMatrix(iLinha, iGrid_DataDebito_Col) = Format(.dtDataEmissao, "dd/mm/yyyy")
            
            End If
            
            'Exibe a sigla do tipo do débito
            GridDebitos.TextMatrix(iLinha, iGrid_TipoDebito_Col) = .sSiglaDocumento
            
            'Exibe o número do débito
            GridDebitos.TextMatrix(iLinha, iGrid_NumDebito_Col) = .lNumero
            
            'Exibe o valor total do débito no formato moeda
            GridDebitos.TextMatrix(iLinha, iGrid_ValorDebito_Col) = Format(.dValor, "Standard")
            
            'Exibe o saldo ainda não utilizado do débito no formato moeda
            GridDebitos.TextMatrix(iLinha, iGrid_SaldoDebito_Col) = Format(.dSaldoNaoApropriado, "Standard")
        
        End With
    
    Next
        
    'Passa para o Obj o número de Parcelas passadas pela Coleção
    objGridDebitos.iLinhasExistentes = gobjBaixaAntecipDebCliente.colAntecipDebCliente.Count

    'Atualiza as checkboxes
    Call Grid_Refresh_Checkbox(objGridDebitos)
    
    Exit Sub

Erro_Traz_Debitos_Tela:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143249)
        
    End Select
    
    Exit Sub
    
End Sub

Private Sub Limpa_Tela_BaixaAntecipDebCliente()

    Call Limpa_Tela(Me)
    
    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade
    
    TipoBaixaAdiantamento.Value = True
    
    Filial.ListIndex = COMBO_INDICE
    
    EmissaoInic.PromptInclude = False
    EmissaoInic.Text = ""
    EmissaoInic.PromptInclude = True

    EmissaoFim.PromptInclude = False
    EmissaoFim.Text = ""
    EmissaoFim.PromptInclude = True
    
    CtaCorrenteTodas.Value = True
    TipoDebitoTodos.Value = True
    
    ContaCorrenteSeleciona.ListIndex = COMBO_INDICE
    TipoDebitoSeleciona.ListIndex = COMBO_INDICE
    
    Call Limpa_Tab_AdiantamentosDebitos

    iClienteAlterado = 0
    iFramePrincipalAlterado = 0
    
    Call Inicializa_gobjBaixaAntecipDebCliente(TELA_TODA)
    
    iAlterado = 0
    
End Sub

Private Sub Limpa_Tab_AdiantamentosDebitos()

On Error GoTo Erro_Limpa_Tab_AdiantamentosDebitos

    Call Grid_Limpa(objGridAdiantamentos)
    Call Grid_Limpa(objGridDebitos)
    
    LabelValorBaixa2.Caption = Format(0, "Fixed")
    
    ContaCorrente.ListIndex = COMBO_INDICE
    
    Historico.Text = ""
    
    DataBaixa.PromptInclude = False
    DataBaixa.Text = Format(gdtDataHoje, "dd/mm/yy")
    DataBaixa.PromptInclude = True
    
    Call Inicializa_gobjBaixaAntecipDebCliente
    
    Exit Sub
    
Erro_Limpa_Tab_AdiantamentosDebitos:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143250)
    
    End Select
    
    Exit Sub
    
End Sub

'inicio contabilidade

Private Sub CTBBotaoModeloPadrao_Click()

    Call objContabil.Contabil_BotaoModeloPadrao_Click

End Sub

Private Sub CTBModelo_Click()

    Call objContabil.Contabil_Modelo_Click

End Sub

Private Sub CTBGridContabil_Click()

    Call objContabil.Contabil_GridContabil_Click

End Sub

Private Sub CTBGridContabil_EnterCell()

    Call objContabil.Contabil_GridContabil_EnterCell

End Sub

Private Sub CTBGridContabil_GotFocus()

    Call objContabil.Contabil_GridContabil_GotFocus

End Sub

Private Sub CTBGridContabil_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_GridContabil_KeyPress(KeyAscii)

End Sub

Private Sub CTBGridContabil_KeyDown(KeyCode As Integer, Shift As Integer)

    Call objContabil.Contabil_GridContabil_KeyDown(KeyCode)
    
End Sub


Private Sub CTBGridContabil_LeaveCell()

        Call objContabil.Contabil_GridContabil_LeaveCell

End Sub

Private Sub CTBGridContabil_Validate(Cancel As Boolean)

    Call objContabil.Contabil_GridContabil_Validate(Cancel)

End Sub

Private Sub CTBGridContabil_RowColChange()

    Call objContabil.Contabil_GridContabil_RowColChange

End Sub

Private Sub CTBGridContabil_Scroll()

    Call objContabil.Contabil_GridContabil_Scroll

End Sub

Private Sub CTBConta_Change()

    Call objContabil.Contabil_Conta_Change

End Sub

Private Sub CTBConta_GotFocus()

    Call objContabil.Contabil_Conta_GotFocus

End Sub

Private Sub CTBConta_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Conta_KeyPress(KeyAscii)

End Sub

Private Sub CTBConta_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Conta_Validate(Cancel)

End Sub

Private Sub CTBCcl_Change()

    Call objContabil.Contabil_Ccl_Change

End Sub

Private Sub CTBCcl_GotFocus()

    Call objContabil.Contabil_Ccl_GotFocus

End Sub

Private Sub CTBCcl_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Ccl_KeyPress(KeyAscii)

End Sub

Private Sub CTBCcl_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Ccl_Validate(Cancel)

End Sub

Private Sub CTBCredito_Change()

    Call objContabil.Contabil_Credito_Change

End Sub

Private Sub CTBCredito_GotFocus()

    Call objContabil.Contabil_Credito_GotFocus

End Sub

Private Sub CTBCredito_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Credito_KeyPress(KeyAscii)

End Sub

Private Sub CTBCredito_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Credito_Validate(Cancel)

End Sub

Private Sub CTBDebito_Change()

    Call objContabil.Contabil_Debito_Change

End Sub

Private Sub CTBDebito_GotFocus()

    Call objContabil.Contabil_Debito_GotFocus

End Sub

Private Sub CTBDebito_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Debito_KeyPress(KeyAscii)

End Sub

Private Sub CTBDebito_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Debito_Validate(Cancel)

End Sub

Private Sub CTBSeqContraPartida_Change()

    Call objContabil.Contabil_SeqContraPartida_Change

End Sub

Private Sub CTBSeqContraPartida_GotFocus()

    Call objContabil.Contabil_SeqContraPartida_GotFocus

End Sub

Private Sub CTBSeqContraPartida_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_SeqContraPartida_KeyPress(KeyAscii)

End Sub

Private Sub CTBSeqContraPartida_Validate(Cancel As Boolean)

    Call objContabil.Contabil_SeqContraPartida_Validate(Cancel)

End Sub

Private Sub CTBHistorico_Change()

    Call objContabil.Contabil_Historico_Change

End Sub

Private Sub CTBHistorico_GotFocus()

    Call objContabil.Contabil_Historico_GotFocus

End Sub

Private Sub CTBHistorico_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Historico_KeyPress(KeyAscii)

End Sub

Private Sub CTBHistorico_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Historico_Validate(Cancel)

End Sub

Private Sub CTBLancAutomatico_Click()

    Call objContabil.Contabil_LancAutomatico_Click

End Sub

Private Sub CTBAglutina_Click()
    
    Call objContabil.Contabil_Aglutina_Click

End Sub

Private Sub CTBAglutina_GotFocus()

    Call objContabil.Contabil_Aglutina_GotFocus

End Sub

Private Sub CTBAglutina_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Aglutina_KeyPress(KeyAscii)

End Sub

Private Sub CTBAglutina_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Aglutina_Validate(Cancel)

End Sub

Private Sub CTBTvwContas_NodeClick(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwContas_NodeClick(Node)

End Sub

Private Sub CTBTvwContas_Expand(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwContas_Expand(Node, CTBTvwContas.Nodes)

End Sub

Private Sub CTBTvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwCcls_NodeClick(Node)

End Sub

Private Sub CTBListHistoricos_DblClick()

    Call objContabil.Contabil_ListHistoricos_DblClick

End Sub

Private Sub CTBBotaoLimparGrid_Click()

    Call objContabil.Contabil_Limpa_GridContabil

End Sub

Private Sub CTBLote_Change()

    Call objContabil.Contabil_Lote_Change

End Sub

Private Sub CTBLote_GotFocus()

    Call objContabil.Contabil_Lote_GotFocus

End Sub

Private Sub CTBLote_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Lote_Validate(Cancel, Parent)

End Sub

Private Sub CTBDataContabil_Change()

    Call objContabil.Contabil_DataContabil_Change

End Sub

Private Sub CTBDataContabil_GotFocus()

    Call objContabil.Contabil_DataContabil_GotFocus

End Sub

Private Sub CTBDataContabil_Validate(Cancel As Boolean)

    Call objContabil.Contabil_DataContabil_Validate(Cancel, Parent)

End Sub

Private Sub objEventoLote_evSelecao(obj1 As Object)
'traz o lote selecionado para a tela

    Call objContabil.Contabil_objEventoLote_evSelecao(obj1)

End Sub

Private Sub objEventoDoc_evSelecao(obj1 As Object)

    Call objContabil.Contabil_objEventoDoc_evSelecao(obj1)

End Sub

Private Sub CTBDocumento_Change()

    Call objContabil.Contabil_Documento_Change

End Sub

Private Sub CTBDocumento_GotFocus()

    Call objContabil.Contabil_Documento_GotFocus

End Sub

Private Sub CTBBotaoImprimir_Click()
    
    Call objContabil.Contabil_BotaoImprimir_Click

End Sub

Private Sub CTBUpDown_DownClick()

    Call objContabil.Contabil_UpDown_DownClick
    
End Sub

Private Sub CTBUpDown_UpClick()

    Call objContabil.Contabil_UpDown_UpClick

End Sub

Private Sub CTBLabelDoc_Click()

    Call objContabil.Contabil_LabelDoc_Click
    
End Sub

Private Sub CTBLabelLote_Click()

    Call objContabil.Contabil_LabelLote_Click
    
End Sub

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objFilial As New ClassFilialCliente, objConta As New ClassContasCorrentesInternas
Dim sContaTela As String

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico

        Case CLIENTE_COD
            
            'Preenche NomeReduzido com o Cliente da tela
            If Len(Trim(Cliente.Text)) > 0 Then
                
                objCliente.sNomeReduzido = Cliente.Text
                lErro = CF("Cliente_Le_NomeReduzido", objCliente)
                If lErro <> SUCESSO Then gError 92806
                
                objMnemonicoValor.colValor.Add objCliente.lCodigo
                
            Else
                
                objMnemonicoValor.colValor.Add 0
                
            End If
            
        Case CLIENTE_NOME
        
            'Preenche NomeReduzido com o Cliente da tela
            If Len(Trim(Cliente.Text)) > 0 Then
                
                objCliente.sNomeReduzido = Cliente.Text
                lErro = CF("Cliente_Le_NomeReduzido", objCliente)
                If lErro <> SUCESSO Then gError 92807
            
                objMnemonicoValor.colValor.Add objCliente.sRazaoSocial
        
            Else
            
                objMnemonicoValor.colValor.Add ""
                
            End If
        
        Case FILIAL_COD
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                objMnemonicoValor.colValor.Add objFilial.iCodFilial
            
            Else
                
                objMnemonicoValor.colValor.Add 0
            
            End If
            
        Case FILIAL_NOME_RED
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Cliente.Text, objFilial)
                If lErro <> SUCESSO Then gError 92808
                
                objMnemonicoValor.colValor.Add objFilial.sNome
            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case FILIAL_CONTA
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Cliente.Text, objFilial)
                If lErro <> SUCESSO Then gError 92809
                
                If objFilial.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError 92810
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case FILIAL_CGC_CPF
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Cliente.Text, objFilial)
                If lErro <> SUCESSO Then gError 92811
                
                objMnemonicoValor.colValor.Add objFilial.sCgc
            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
        
        Case DATA1
            If Len(DataBaixa.ClipText) > 0 Then
                objMnemonicoValor.colValor.Add CDate(DataBaixa.FormattedText)
            Else
                objMnemonicoValor.colValor.Add DATA_NULA
            End If
        
        Case CONTA_COD
            If ContaCorrente.ListIndex <> -1 Then
                objMnemonicoValor.colValor.Add ContaCorrente.ItemData(ContaCorrente.ListIndex)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
        
        Case CONTA_CONTABIL_CONTA
        
            If ContaCorrente.ListIndex <> -1 Then
            
                objConta.iCodigo = ContaCorrente.ItemData(ContaCorrente.ListIndex)
                lErro = CF("ContaCorrenteInt_Le", objConta.iCodigo, objConta)
                If lErro <> SUCESSO Then gError 92812
                
                If objConta.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objConta.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError 92813
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                
            Else
                objMnemonicoValor.colValor.Add ""
            End If
                                
        Case VALOR1
            If Len(LabelValorBaixa2.Caption) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(LabelValorBaixa2.Caption)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
        
        Case HISTORICO1
            If Len(Historico.Text) > 0 Then
                objMnemonicoValor.colValor.Add Historico.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case Else
            gError 92814

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 92806 To 92813
        
        Case 92814
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143251)

    End Select

    Exit Function

End Function

Private Sub Cliente_Preenche()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objCliente As Object
    
On Error GoTo Erro_Cliente_Preenche
    
    Set objCliente = Cliente
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objCliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134028

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 134028

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143252)

    End Select
    
    Exit Sub

End Sub


Private Sub CTBGerencial_Click()
    
    Call objContabil.Contabil_Gerencial_Click

End Sub

Private Sub CTBGerencial_GotFocus()

    Call objContabil.Contabil_Gerencial_GotFocus

End Sub

Private Sub CTBGerencial_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Gerencial_KeyPress(KeyAscii)

End Sub

Private Sub CTBGerencial_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Gerencial_Validate(Cancel)

End Sub

Sub Refaz_Grid(ByVal objGridInt As AdmGrid, ByVal iNumLinhas As Integer)
    objGridInt.objGrid.Rows = iNumLinhas + 1

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
End Sub

Public Sub mnuGridMarcarTodos_Click()
Dim iIndice As Integer
    Select Case iGridAtual
         Case GRID_DEVOLUCOES
            Call Grid_Marca_Desmarca(objGridDebitos, iGrid_DebitoSeleciona_Col, MARCADO)
            For iIndice = 1 To gobjBaixaAntecipDebCliente.colAntecipDebCliente.Count
                gobjBaixaAntecipDebCliente.colAntecipDebCliente(iIndice).iSelecionado = MARCADO
            Next
            Call Calcula_Total_Baixar(TIPOBAIXA_DEBITOS)

         Case GRID_ADIANTAMENTOS
            Call Grid_Marca_Desmarca(objGridAdiantamentos, iGrid_AdiantamentoSeleciona_Col, MARCADO)
            For iIndice = 1 To gobjBaixaAntecipDebCliente.colAntecipDebCliente.Count
                gobjBaixaAntecipDebCliente.colAntecipDebCliente(iIndice).iSelecionado = MARCADO
            Next
            Call Calcula_Total_Baixar(TIPOBAIXA_ADIANTAMENTO)
    End Select
End Sub

Public Sub mnuGridDesmarcarTodos_Click()
Dim iIndice As Integer
    Select Case iGridAtual
         Case GRID_DEVOLUCOES
            Call Grid_Marca_Desmarca(objGridDebitos, iGrid_DebitoSeleciona_Col, DESMARCADO)
            For iIndice = 1 To gobjBaixaAntecipDebCliente.colAntecipDebCliente.Count
                gobjBaixaAntecipDebCliente.colAntecipDebCliente(iIndice).iSelecionado = DESMARCADO
            Next
            Call Calcula_Total_Baixar(TIPOBAIXA_DEBITOS)

         Case GRID_ADIANTAMENTOS
            Call Grid_Marca_Desmarca(objGridAdiantamentos, iGrid_AdiantamentoSeleciona_Col, DESMARCADO)
            For iIndice = 1 To gobjBaixaAntecipDebCliente.colAntecipDebCliente.Count
                gobjBaixaAntecipDebCliente.colAntecipDebCliente(iIndice).iSelecionado = DESMARCADO
            Next
            Call Calcula_Total_Baixar(TIPOBAIXA_ADIANTAMENTO)
    End Select
End Sub

Public Sub mnuGridConsultaDocOriginal_Click()
'Chama a tela de consulta de Títulos a Receber quando essa opção for selecionada no grid

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente

On Error GoTo Erro_mnuGridConsultaDocOriginal_Click
       
    'Verifica em qual grid está o documento que será consultado
    Select Case iGridAtual
    
        'Se for o grid de créditos/ devoluções
         Case GRID_DEVOLUCOES
                      
             'Abre a tela de cadastro de créditos/devoluções
             lErro = Consulta_DebitosReceber(objFilialCliente)
             If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
         
         'Se for o grid de adiantamentos
         Case GRID_ADIANTAMENTOS
         
             'Abre a tela para cadastro de adiantamentos
             lErro = Consulta_RecebAntecipados(objFilialCliente)
             If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End Select
    
    Exit Sub
    
Erro_mnuGridConsultaDocOriginal_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143441)
            
    End Select
    
    Exit Sub
    
End Sub

Private Function Consulta_DebitosReceber(objFilialCliente As ClassFilialCliente) As Long
'Exibe o débito com cliente em sua tela de cadastro

Dim lErro As Long
Dim objDebitoReceber As New ClassDebitoRecCli
Dim objAntecipDebCliente As New ClassInfoBaixaAntecipDebCli

On Error GoTo Erro_Consulta_DebitosReceber

    'Se não existe linha selecionada no grid de devoluções => erro
    If GridDebitos.Row <= 0 Then gError 79921
    
    Set objAntecipDebCliente = gobjBaixaAntecipDebCliente.colAntecipDebCliente.Item(GridDebitos.Row)
    
    'Seleciona o objDebitoReceber referente à linha cujo documento será consultado
    objDebitoReceber.lNumIntDoc = objAntecipDebCliente.lNumIntDocumento

    'Chama a tela de cadastro de Débitos a Receber e exibe o débito selecionado
    Call Chama_Tela("DebitosReceb", objDebitoReceber)

    Consulta_DebitosReceber = SUCESSO

    Exit Function

Erro_Consulta_DebitosReceber:

    Consulta_DebitosReceber = gErr

    Select Case gErr

        Case 79921
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143443)

    End Select

    Exit Function

End Function

Private Function Consulta_RecebAntecipados(objFilialCliente As ClassFilialCliente) As Long
'Exibe o Adiantamento à Cliente em sua tela de cadastro

Dim lErro As Long
Dim objAntecipReceb As New ClassAntecipReceb
Dim objMovContaCorrente As New ClassMovContaCorrente
Dim objAntecipDebCliente As New ClassInfoBaixaAntecipDebCli

On Error GoTo Erro_Consulta_RecebAntecipados

    'Se não existe linha selecionada no grid de adiantamentos=> erro
    If GridAdiantamentos.Row <= 0 Then gError 79929

    Set objAntecipDebCliente = gobjBaixaAntecipDebCliente.colAntecipDebCliente.Item(GridAdiantamentos.Row)

    objAntecipReceb.lNumMovto = objAntecipDebCliente.lNumMovto
    objAntecipReceb.lNumIntRec = objAntecipDebCliente.lNumIntDocumento

    'Exibe o Adiantamento em sua tela de cadastro
    Call Chama_Tela("AntecipReceb", objAntecipReceb)

    Consulta_RecebAntecipados = SUCESSO

    Exit Function

Erro_Consulta_RecebAntecipados:

    Consulta_RecebAntecipados = gErr

    Select Case gErr

        Case 79929
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143444)

    End Select

    Exit Function
    
End Function
