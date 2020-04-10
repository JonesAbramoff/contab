VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl GeracaoNFiscal 
   ClientHeight    =   5595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8880
   KeyPreview      =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   8880
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4920
      Index           =   2
      Left            =   150
      TabIndex        =   10
      Top             =   585
      Visible         =   0   'False
      Width           =   8565
      Begin VB.CommandButton BotaoImprimirPI 
         Caption         =   "Pedido Interno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4710
         Picture         =   "GeracaoNFiscalArt.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   110
         ToolTipText     =   "Imprimir Pedido Interno"
         Top             =   4260
         Width           =   1485
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   4395
         Index           =   8
         Left            =   7845
         TabIndex        =   67
         Top             =   765
         Visible         =   0   'False
         Width           =   9240
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
            Left            =   3480
            TabIndex        =   81
            Top             =   930
            Value           =   1  'Checked
            Width           =   2745
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
            Height          =   375
            Left            =   7710
            TabIndex        =   80
            Top             =   135
            Width           =   1245
         End
         Begin VB.Frame CTBFrame7 
            Caption         =   "Descrição do Elemento Selecionado"
            Height          =   1050
            Left            =   195
            TabIndex        =   75
            Top             =   3330
            Width           =   5895
            Begin VB.Label CTBCclDescricao 
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1845
               TabIndex        =   79
               Top             =   645
               Visible         =   0   'False
               Width           =   3720
            End
            Begin VB.Label CTBContaDescricao 
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1845
               TabIndex        =   78
               Top             =   285
               Width           =   3720
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
               Height          =   195
               Left            =   1125
               TabIndex        =   77
               Top             =   315
               Width           =   570
            End
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
               Height          =   195
               Left            =   240
               TabIndex        =   76
               Top             =   660
               Visible         =   0   'False
               Width           =   1440
            End
         End
         Begin VB.ComboBox CTBModelo 
            Height          =   315
            Left            =   7740
            Style           =   2  'Dropdown List
            TabIndex        =   74
            Top             =   810
            Width           =   1260
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
            Height          =   375
            Left            =   6330
            TabIndex        =   73
            Top             =   120
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
            Height          =   510
            Left            =   6330
            TabIndex        =   72
            Top             =   630
            Width           =   1245
         End
         Begin VB.ListBox CTBListHistoricos 
            Height          =   2790
            Left            =   6330
            TabIndex        =   71
            Top             =   1560
            Visible         =   0   'False
            Width           =   2625
         End
         Begin VB.TextBox CTBHistorico 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   4245
            MaxLength       =   150
            TabIndex        =   70
            Top             =   2175
            Width           =   1770
         End
         Begin VB.CheckBox CTBAglutina 
            Height          =   210
            Left            =   4470
            TabIndex        =   69
            Top             =   2565
            Width           =   870
         End
         Begin MSMask.MaskEdBox CTBSeqContraPartida 
            Height          =   225
            Left            =   4800
            TabIndex        =   68
            Top             =   1560
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
            TabIndex        =   82
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
            TabIndex        =   83
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
            TabIndex        =   84
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
            TabIndex        =   85
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
         Begin MSComCtl2.UpDown CTBUpDown3 
            Height          =   300
            Left            =   1635
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   540
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox CTBDataContabil3 
            Height          =   300
            Left            =   570
            TabIndex        =   87
            Top             =   525
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CTBLote3 
            Height          =   300
            Left            =   5580
            TabIndex        =   88
            Top             =   135
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
            Left            =   1845
            TabIndex        =   89
            Top             =   3030
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   5
            Mask            =   "#####"
            PromptChar      =   " "
         End
         Begin MSComctlLib.TreeView CTBTvwCcls 
            Height          =   2790
            Left            =   6330
            TabIndex        =   90
            Top             =   1560
            Visible         =   0   'False
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   4921
            _Version        =   393217
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            BorderStyle     =   1
            Appearance      =   1
         End
         Begin MSComctlLib.TreeView CTBTvwContas 
            Height          =   2790
            Left            =   6330
            TabIndex        =   91
            Top             =   1560
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   4921
            _Version        =   393217
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            BorderStyle     =   1
            Appearance      =   1
         End
         Begin MSFlexGridLib.MSFlexGrid CTBGridContabil 
            Height          =   1860
            Left            =   0
            TabIndex        =   92
            Top             =   1170
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   3281
            _Version        =   393216
            Rows            =   7
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin VB.Label CTBLabelLote3 
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
            Height          =   195
            Left            =   5100
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   109
            Top             =   165
            Width           =   450
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
            Height          =   195
            Left            =   750
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   108
            Top             =   3075
            Width           =   1035
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
            TabIndex        =   107
            Top             =   555
            Width           =   480
         End
         Begin VB.Label CTBTotalCredito 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2460
            TabIndex        =   106
            Top             =   3030
            Width           =   1155
         End
         Begin VB.Label CTBTotalDebito 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3705
            TabIndex        =   105
            Top             =   3030
            Width           =   1155
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
            Height          =   225
            Left            =   1800
            TabIndex        =   104
            Top             =   3045
            Width           =   615
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
            Height          =   195
            Left            =   7755
            TabIndex        =   103
            Top             =   585
            Width           =   690
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
            Left            =   6360
            TabIndex        =   102
            Top             =   1290
            Visible         =   0   'False
            Width           =   2490
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
            TabIndex        =   101
            Top             =   1305
            Width           =   2340
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
            TabIndex        =   100
            Top             =   1275
            Visible         =   0   'False
            Width           =   1005
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
            Top             =   945
            Width           =   1140
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
            TabIndex        =   98
            Top             =   585
            Width           =   870
         End
         Begin VB.Label CTBExercicio 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2910
            TabIndex        =   97
            Top             =   555
            Width           =   1185
         End
         Begin VB.Label CTBPeriodo 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5010
            TabIndex        =   96
            Top             =   570
            Width           =   1185
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
            TabIndex        =   95
            Top             =   600
            Width           =   735
         End
         Begin VB.Label CTBOrigem 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4305
            TabIndex        =   94
            Top             =   3075
            Width           =   1530
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
            Height          =   255
            Left            =   3600
            TabIndex        =   93
            Top             =   3120
            Width           =   720
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GridItens 
         Height          =   1575
         Left            =   7935
         TabIndex        =   66
         Top             =   615
         Visible         =   0   'False
         Width           =   8880
         _ExtentX        =   15663
         _ExtentY        =   2778
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin MSComCtl2.UpDown UpDownSaida 
         Height          =   300
         Left            =   3660
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   105
         Width           =   225
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Motivo 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   2520
         TabIndex        =   27
         Text            =   "Motivo"
         Top             =   3795
         Width           =   3690
      End
      Begin VB.TextBox Transportadora 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   2775
         TabIndex        =   26
         Text            =   "Transportadora"
         Top             =   3345
         Width           =   2055
      End
      Begin VB.ComboBox Serie 
         Height          =   315
         Left            =   5325
         TabIndex        =   12
         Top             =   105
         Width           =   990
      End
      Begin VB.CommandButton BotaoNFiscalFaturaImprime 
         Caption         =   "Gerar Nota Fiscal Fatura e Imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4035
         TabIndex        =   16
         Top             =   1035
         Width           =   3225
      End
      Begin VB.CommandButton BotaoNFiscalFatura 
         Caption         =   "Gerar Nota Fiscal Fatura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1110
         TabIndex        =   15
         Top             =   1035
         Width           =   2370
      End
      Begin VB.TextBox Bairro 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   6480
         TabIndex        =   32
         Text            =   "Bairro"
         Top             =   3495
         Width           =   1665
      End
      Begin VB.CommandButton BotaoPedido 
         Caption         =   "Editar Pedido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6345
         Picture         =   "GeracaoNFiscalArt.ctx":0102
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   4260
         Width           =   1485
      End
      Begin VB.ComboBox Ordenados 
         Height          =   315
         ItemData        =   "GeracaoNFiscalArt.ctx":0D80
         Left            =   2835
         List            =   "GeracaoNFiscalArt.ctx":0D82
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1950
         Width           =   3480
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2040
         Picture         =   "GeracaoNFiscalArt.ctx":0D84
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   4245
         Width           =   1800
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   105
         Picture         =   "GeracaoNFiscalArt.ctx":1F66
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   4245
         Width           =   1800
      End
      Begin VB.TextBox DataEmissao 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   4980
         TabIndex        =   25
         Text            =   "Emissão"
         Top             =   2550
         Width           =   1050
      End
      Begin VB.TextBox Cliente 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   1485
         TabIndex        =   22
         Text            =   "Cliente"
         Top             =   2535
         Width           =   1080
      End
      Begin VB.TextBox Pedido 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   615
         TabIndex        =   21
         Text            =   "Pedido"
         Top             =   2535
         Width           =   852
      End
      Begin VB.CheckBox GeraNFiscal 
         DragMode        =   1  'Automatic
         Height          =   210
         Left            =   195
         TabIndex        =   20
         Top             =   2535
         Width           =   816
      End
      Begin VB.CommandButton BotaoNFiscal 
         Caption         =   "Gerar Nota Fiscal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1095
         TabIndex        =   17
         Top             =   1470
         Width           =   2370
      End
      Begin VB.CommandButton BotaoNFiscalImprime 
         Caption         =   "Gerar Nota Fiscal e Imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4035
         TabIndex        =   18
         Top             =   1470
         Width           =   3240
      End
      Begin VB.TextBox DataEntrega 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   6270
         TabIndex        =   28
         Text            =   "Entrega"
         Top             =   2505
         Width           =   1095
      End
      Begin VB.TextBox NomeReduzido 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   2520
         TabIndex        =   23
         Text            =   "Nome Reduzido"
         Top             =   2520
         Width           =   2040
      End
      Begin VB.TextBox Filial 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   4485
         TabIndex        =   24
         Text            =   "Filial"
         Top             =   2535
         Width           =   540
      End
      Begin VB.TextBox Estado 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   6390
         TabIndex        =   30
         Text            =   "Estado"
         Top             =   2880
         Width           =   1665
      End
      Begin VB.TextBox Cidade 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   6495
         TabIndex        =   31
         Text            =   "Cidade"
         Top             =   3165
         Width           =   1680
      End
      Begin VB.TextBox FilialEmpresa 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   6585
         TabIndex        =   33
         Text            =   "Filial"
         Top             =   3810
         Width           =   1560
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   240
         Left            =   7350
         TabIndex        =   29
         Top             =   2580
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSFlexGridLib.MSFlexGrid GridPedido 
         Height          =   1845
         Left            =   90
         TabIndex        =   34
         Top             =   2355
         Width           =   8400
         _ExtentX        =   14817
         _ExtentY        =   3254
         _Version        =   393216
         Rows            =   10
         Cols            =   7
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin MSMask.MaskEdBox DataSaida 
         Height          =   300
         Left            =   2565
         TabIndex        =   11
         Top             =   105
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown CTBUpDown 
         Height          =   300
         Left            =   3645
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   570
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox CTBDataContabil 
         Height          =   300
         Left            =   2565
         TabIndex        =   13
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
         Left            =   5325
         TabIndex        =   14
         Top             =   570
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridAlocacao 
         Height          =   1860
         Left            =   7590
         TabIndex        =   65
         Top             =   435
         Visible         =   0   'False
         Width           =   7290
         _ExtentX        =   12859
         _ExtentY        =   3281
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Pedidos"
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
         Left            =   150
         TabIndex        =   64
         Top             =   2160
         Width           =   690
      End
      Begin VB.Label CTBLabelLote 
         AutoSize        =   -1  'True
         Caption         =   "Lote Contábil:"
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
         Left            =   4050
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   63
         Top             =   615
         Width           =   1200
      End
      Begin VB.Label LabelDataContabil 
         AutoSize        =   -1  'True
         Caption         =   "Data de Contabilização:"
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
         Left            =   480
         TabIndex        =   62
         Top             =   615
         Width           =   2055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ordenados por:"
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
         Left            =   1440
         TabIndex        =   60
         Top             =   1995
         Width           =   1320
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Data de Saída :"
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
         Left            =   1140
         TabIndex        =   57
         Top             =   135
         Width           =   1380
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Série:"
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
         Left            =   4725
         TabIndex        =   59
         Top             =   165
         Width           =   510
      End
   End
   Begin VB.CommandButton BotaoFechar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7500
      Picture         =   "GeracaoNFiscalArt.ctx":2F80
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Fechar"
      Top             =   120
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4860
      Index           =   1
      Left            =   180
      TabIndex        =   0
      Top             =   615
      Width           =   8565
      Begin VB.Frame Frame2 
         Caption         =   "Exibe Pedidos"
         Height          =   4305
         Left            =   870
         TabIndex        =   40
         Top             =   180
         Width           =   6270
         Begin VB.Frame Frame3 
            Caption         =   "Data Emissão"
            Height          =   810
            Left            =   450
            TabIndex        =   47
            Top             =   2310
            Width           =   5505
            Begin MSMask.MaskEdBox DataEmissaoDe 
               Height          =   300
               Left            =   795
               TabIndex        =   6
               Top             =   315
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownEmissaoDe 
               Height          =   300
               Left            =   1950
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   315
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataEmissaoAte 
               Height          =   300
               Left            =   3435
               TabIndex        =   7
               Top             =   315
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownEmissaoAte 
               Height          =   300
               Left            =   4605
               TabIndex        =   51
               TabStop         =   0   'False
               Top             =   315
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin VB.Label Label1 
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
               Height          =   195
               Index           =   1
               Left            =   2985
               TabIndex        =   50
               Top             =   368
               Width           =   360
            End
            Begin VB.Label Label1 
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
               Height          =   195
               Index           =   0
               Left            =   345
               TabIndex        =   48
               Top             =   368
               Width           =   315
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Pedidos"
            Height          =   810
            Left            =   420
            TabIndex        =   41
            Top             =   495
            Width           =   5520
            Begin MSMask.MaskEdBox PedidoInicial 
               Height          =   300
               Left            =   810
               TabIndex        =   2
               Top             =   315
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox PedidoFinal 
               Height          =   300
               Left            =   3450
               TabIndex        =   3
               Top             =   315
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin VB.Label LabelPedidoDe 
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
               Left            =   315
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   42
               Top             =   368
               Width           =   315
            End
            Begin VB.Label LabelPedidoAte 
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
               Height          =   195
               Left            =   2985
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   43
               Top             =   368
               Width           =   360
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Clientes"
            Height          =   810
            Left            =   435
            TabIndex        =   44
            Top             =   1395
            Width           =   5520
            Begin MSMask.MaskEdBox ClienteDe 
               Height          =   300
               Left            =   810
               TabIndex        =   4
               Top             =   330
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ClienteAte 
               Height          =   300
               Left            =   3450
               TabIndex        =   5
               Top             =   330
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin VB.Label LabelClienteAte 
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
               Height          =   195
               Left            =   2970
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   46
               Top             =   390
               Width           =   360
            End
            Begin VB.Label LabelClienteDe 
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
               Left            =   315
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   45
               Top             =   383
               Width           =   315
            End
         End
         Begin VB.CheckBox ExibeTodos 
            Caption         =   "Exibe Todos os Pedidos"
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
            Left            =   1650
            TabIndex        =   1
            Top             =   270
            Width           =   2430
         End
         Begin VB.Frame Frame5 
            Caption         =   "Data Entrega"
            Height          =   810
            Left            =   450
            TabIndex        =   52
            Top             =   3330
            Width           =   5505
            Begin MSMask.MaskEdBox DataEntregaDe 
               Height          =   300
               Left            =   795
               TabIndex        =   8
               Top             =   315
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownEntregaDe 
               Height          =   300
               Left            =   1935
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   315
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataEntregaAte 
               Height          =   300
               Left            =   3420
               TabIndex        =   9
               Top             =   315
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownEntregaAte 
               Height          =   300
               Left            =   4590
               TabIndex        =   56
               TabStop         =   0   'False
               Top             =   315
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin VB.Label Label1 
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
               Height          =   195
               Index           =   2
               Left            =   330
               TabIndex        =   53
               Top             =   368
               Width           =   315
            End
            Begin VB.Label Label1 
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
               Height          =   195
               Index           =   3
               Left            =   2985
               TabIndex        =   55
               Top             =   368
               Width           =   360
            End
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5280
      Left            =   105
      TabIndex        =   39
      Top             =   270
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   9313
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pedidos"
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
Attribute VB_Name = "GeracaoNFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'???? Alguns campos da nota fiscal não estão sendo preenchidos: DataEntrada,FilialPedido,NumPedidoVenda e ValorProdutos. Rotina NFiscal_ExtrairPV no ClassFATGrava.
'???? Alterar a rotina de leitura dos pedidos só incluindo as condições no SELECT se a opção de exibir todos os pedidos não estiver "setada".
'???? Chamei a tela de GeracaoNFiscal e depois apertei Gerar NFiscal. Voou tudo e travou o VB. Fecha e abrir os comando caso haja erro na geração de alguma NFiscal.
'???? Depois de gerar a NF não está atualizando o grid de pedidos. E quando eu coloco para atualizar aparece a msg de que não existem mais pedidos… Não devo exibir msg.
'???? O erro 44295 não está sendo tartado em Processa_Gravacao_NFiscal
'???? O grid de pedido não cabe na tela.
'???? Sem comentário:BotaoNFiscal , BotaoNFiscalFatura, NFiscalImprime, BotaoNFiscalImprime, ClienteDe_Validate

'???? Como a data de entrega é escolhida pela menor data de entrega dos itens, se algum item não possuir data de entrega a data da tela fica em branco, pois a data nula é menor que qq outra data digitda pelo usuário
'???? Para abrir a tela de pedido não deveria apenas ser necessário que um pedido estivesse com o foco?

Option Explicit

Event Unload()

Private WithEvents objCT As CTGeracaoNFiscal
Attribute objCT.VB_VarHelpID = -1

Private Sub UserControl_Initialize()
    Set objCT = New CTGeracaoNFiscal
    Set objCT.objUserControl = Me
End Sub

Private Sub BotaoDesmarcarTodos_Click()
     Call objCT.BotaoDesmarcarTodos_Click
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub BotaoMarcarTodos_Click()
     Call objCT.BotaoMarcarTodos_Click
End Sub

Private Sub BotaoNFiscal_Click()
     Call objCT.BotaoNFiscal_Click
End Sub

Private Sub BotaoNFiscalFatura_Click()
     Call objCT.BotaoNFiscalFatura_Click
End Sub

Private Sub BotaoNFiscalFaturaImprime_Click()
     Call objCT.BotaoNFiscalFaturaImprime_Click
End Sub

Private Sub BotaoNFiscalImprime_Click()
     Call objCT.BotaoNFiscalImprime_Click
End Sub

Private Sub BotaoPedido_Click()
     Call objCT.BotaoPedido_Click
End Sub

Private Sub ClienteAte_Change()
     Call objCT.ClienteAte_Change
End Sub

Private Sub ClienteAte_GotFocus()
     Call objCT.ClienteAte_GotFocus
End Sub

Private Sub ClienteAte_Validate(Cancel As Boolean)
     Call objCT.ClienteAte_Validate(Cancel)
End Sub

Private Sub ClienteDe_Change()
     Call objCT.ClienteDe_Change
End Sub

Private Sub ClienteDe_GotFocus()
     Call objCT.ClienteDe_GotFocus
End Sub

Private Sub ClienteDe_Validate(Cancel As Boolean)
     Call objCT.ClienteDe_Validate(Cancel)
End Sub

Private Sub CTBUpDown_DownClick()
     Call objCT.CTBUpDown_DownClick
End Sub

Private Sub CTBUpDown_UpClick()
     Call objCT.CTBUpDown_UpClick
End Sub

Private Sub DataEmissaoAte_Change()
     Call objCT.DataEmissaoAte_Change
End Sub

Private Sub DataEmissaoAte_GotFocus()
     Call objCT.DataEmissaoAte_GotFocus
End Sub

Private Sub DataEmissaoAte_Validate(Cancel As Boolean)
     Call objCT.DataEmissaoAte_Validate(Cancel)
End Sub

Private Sub DataEmissaoDe_Change()
     Call objCT.DataEmissaoDe_Change
End Sub

Private Sub DataEmissaoDe_GotFocus()
     Call objCT.DataEmissaoDe_GotFocus
End Sub

Private Sub DataEmissaoDe_Validate(Cancel As Boolean)
     Call objCT.DataEmissaoDe_Validate(Cancel)
End Sub

Private Sub DataEntregaAte_Change()
     Call objCT.DataEntregaAte_Change
End Sub

Private Sub DataEntregaAte_GotFocus()
     Call objCT.DataEntregaAte_GotFocus
End Sub

Private Sub DataEntregaAte_Validate(Cancel As Boolean)
     Call objCT.DataEntregaAte_Validate(Cancel)
End Sub

Private Sub DataEntregaDe_Change()
     Call objCT.DataEntregaDe_Change
End Sub

Private Sub DataEntregaDe_GotFocus()
     Call objCT.DataEntregaDe_GotFocus
End Sub

Private Sub DataEntregaDe_Validate(Cancel As Boolean)
     Call objCT.DataEntregaDe_Validate(Cancel)
End Sub

Private Sub DataSaida_Change()
     Call objCT.DataSaida_Change
End Sub

Private Sub DataSaida_GotFocus()
     Call objCT.DataSaida_GotFocus
End Sub

Private Sub ExibeTodos_Click()
     Call objCT.ExibeTodos_Click
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Private Sub GeraNFiscal_Click()
     Call objCT.GeraNFiscal_Click
End Sub

Private Sub GeraNFiscal_GotFocus()
     Call objCT.GeraNFiscal_GotFocus
End Sub

Private Sub GeraNFiscal_KeyPress(KeyAscii As Integer)
     Call objCT.GeraNFiscal_KeyPress(KeyAscii)
End Sub

Private Sub GeraNFiscal_Validate(Cancel As Boolean)
     Call objCT.GeraNFiscal_Validate(Cancel)
End Sub

Private Sub GridPedido_Click()
     Call objCT.GridPedido_Click
End Sub

Private Sub GridPedido_EnterCell()
     Call objCT.GridPedido_EnterCell
End Sub

Private Sub GridPedido_GotFocus()
     Call objCT.GridPedido_GotFocus
End Sub

Private Sub GridPedido_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridPedido_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridPedido_KeyPress(KeyAscii As Integer)
     Call objCT.GridPedido_KeyPress(KeyAscii)
End Sub

Private Sub GridPedido_LeaveCell()
     Call objCT.GridPedido_LeaveCell
End Sub

Private Sub GridPedido_Validate(Cancel As Boolean)
     Call objCT.GridPedido_Validate(Cancel)
End Sub

Private Sub GridPedido_RowColChange()
     Call objCT.GridPedido_RowColChange
End Sub

Private Sub GridPedido_Scroll()
     Call objCT.GridPedido_Scroll
End Sub

Private Sub LabelClienteAte_Click()
     Call objCT.LabelClienteAte_Click
End Sub

Private Sub LabelClienteDe_Click()
     Call objCT.LabelClienteDe_Click
End Sub

Private Sub LabelPedidoAte_Click()
     Call objCT.LabelPedidoAte_Click
End Sub

Private Sub LabelPedidoDe_Click()
     Call objCT.LabelPedidoDe_Click
End Sub

Private Sub Ordenados_Change()
     Call objCT.Ordenados_Change
End Sub

Private Sub Ordenados_Click()
     Call objCT.Ordenados_Click
End Sub

Private Sub PedidoFinal_Change()
     Call objCT.PedidoFinal_Change
End Sub

Private Sub PedidoFinal_GotFocus()
     Call objCT.PedidoFinal_GotFocus
End Sub

Private Sub PedidoFinal_Validate(Cancel As Boolean)
     Call objCT.PedidoFinal_Validate(Cancel)
End Sub

Private Sub PedidoInicial_Change()
     Call objCT.PedidoInicial_Change
End Sub

Private Sub PedidoInicial_GotFocus()
     Call objCT.PedidoInicial_GotFocus
End Sub

Private Sub PedidoInicial_Validate(Cancel As Boolean)
     Call objCT.PedidoInicial_Validate(Cancel)
End Sub

Private Sub Serie_Validate(Cancel As Boolean)
     Call objCT.Serie_Validate(Cancel)
End Sub

Private Sub TabStrip1_Click()
     Call objCT.TabStrip1_Click
End Sub

Private Sub UpDownEmissaoAte_DownClick()
     Call objCT.UpDownEmissaoAte_DownClick
End Sub

Private Sub UpDownEmissaoAte_UpClick()
     Call objCT.UpDownEmissaoAte_UpClick
End Sub

Private Sub UpDownEmissaoDe_DownClick()
     Call objCT.UpDownEmissaoDe_DownClick
End Sub

Private Sub UpDownEmissaoDe_UpClick()
     Call objCT.UpDownEmissaoDe_UpClick
End Sub

Private Sub UpDownEntregaAte_DownClick()
     Call objCT.UpDownEntregaAte_DownClick
End Sub

Private Sub UpDownEntregaAte_UpClick()
     Call objCT.UpDownEntregaAte_UpClick
End Sub

Private Sub UpDownEntregaDe_DownClick()
     Call objCT.UpDownEntregaDe_DownClick
End Sub

Private Sub UpDownEntregaDe_UpClick()
     Call objCT.UpDownEntregaDe_UpClick
End Sub

Private Sub DataSaida_Validate(Cancel As Boolean)
     Call objCT.DataSaida_Validate(Cancel)
End Sub

Private Sub UpDownSaida_DownClick()
     Call objCT.UpDownSaida_DownClick
End Sub

Private Sub UpDownSaida_UpClick()
     Call objCT.UpDownSaida_UpClick
End Sub

Function Trata_Parametros() As Long
     Trata_Parametros = objCT.Trata_Parametros()
End Function

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub
Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub
Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub
Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub
Private Sub label_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(label, Source, X, Y)
End Sub
Private Sub label_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(label, Button, Shift, X, Y)
End Sub
Private Sub Label19_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label19, Source, X, Y)
End Sub
Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label19, Button, Shift, X, Y)
End Sub
Private Sub LabelPedidoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelPedidoDe, Source, X, Y)
End Sub
Private Sub LabelPedidoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelPedidoDe, Button, Shift, X, Y)
End Sub
Private Sub LabelPedidoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelPedidoAte, Source, X, Y)
End Sub
Private Sub LabelPedidoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelPedidoAte, Button, Shift, X, Y)
End Sub
Private Sub LabelClienteAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteAte, Source, X, Y)
End Sub
Private Sub LabelClienteAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteAte, Button, Shift, X, Y)
End Sub
Private Sub LabelClienteDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteDe, Source, X, Y)
End Sub
Private Sub LabelClienteDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteDe, Button, Shift, X, Y)
End Sub
Private Sub CTBLancAutomatico_Click()
     Call objCT.CTBLancAutomatico_Click
End Sub

Private Sub CTBAglutina_Click()
     Call objCT.CTBAglutina_Click
End Sub

Private Sub LabelDataContabil_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataContabil, Button, Shift, X, Y)
End Sub
Private Sub LabelDataContabil_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataContabil, Source, X, Y)
End Sub
Private Sub CTBLabelLote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelLote, Button, Shift, X, Y)
End Sub
Private Sub CTBLabelLote_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelLote, Source, X, Y)
End Sub
Private Sub TabStrip1_BeforeClick(Cancel As Integer)
     Call objCT.TabStrip1_BeforeClick(Cancel)
End Sub

Private Sub CTBCclDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclDescricao, Source, X, Y)
End Sub
Private Sub CTBCclDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclDescricao, Button, Shift, X, Y)
End Sub
Private Sub CTBContaDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBContaDescricao, Source, X, Y)
End Sub
Private Sub CTBContaDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBContaDescricao, Button, Shift, X, Y)
End Sub
Private Sub CTBLabel7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel7, Source, X, Y)
End Sub
Private Sub CTBLabel7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel7, Button, Shift, X, Y)
End Sub
Private Sub CTBCclLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclLabel, Source, X, Y)
End Sub
Private Sub CTBCclLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclLabel, Button, Shift, X, Y)
End Sub
Private Sub CTBLabelLote3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelLote3, Source, X, Y)
End Sub
Private Sub CTBLabelLote3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelLote3, Button, Shift, X, Y)
End Sub
Private Sub CTBLabelDoc_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelDoc, Source, X, Y)
End Sub
Private Sub CTBLabelDoc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelDoc, Button, Shift, X, Y)
End Sub
Private Sub CTBLabel8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel8, Source, X, Y)
End Sub
Private Sub CTBLabel8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel8, Button, Shift, X, Y)
End Sub
Private Sub CTBTotalCredito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBTotalCredito, Source, X, Y)
End Sub
Private Sub CTBTotalCredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBTotalCredito, Button, Shift, X, Y)
End Sub
Private Sub CTBTotalDebito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBTotalDebito, Source, X, Y)
End Sub
Private Sub CTBTotalDebito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBTotalDebito, Button, Shift, X, Y)
End Sub
Private Sub CTBLabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelTotais, Source, X, Y)
End Sub
Private Sub CTBLabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelTotais, Button, Shift, X, Y)
End Sub
Private Sub CTBLabel1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel1, Source, X, Y)
End Sub
Private Sub CTBLabel1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel1, Button, Shift, X, Y)
End Sub
Private Sub CTBLabelCcl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelCcl, Source, X, Y)
End Sub
Private Sub CTBLabelCcl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelCcl, Button, Shift, X, Y)
End Sub
Private Sub CTBLabelContas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelContas, Source, X, Y)
End Sub
Private Sub CTBLabelContas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelContas, Button, Shift, X, Y)
End Sub
Private Sub CTBLabelHistoricos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelHistoricos, Source, X, Y)
End Sub
Private Sub CTBLabelHistoricos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelHistoricos, Button, Shift, X, Y)
End Sub
Private Sub CTBLabel5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel5, Source, X, Y)
End Sub
Private Sub CTBLabel5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel5, Button, Shift, X, Y)
End Sub
Private Sub CTBLabel13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel13, Source, X, Y)
End Sub
Private Sub CTBLabel13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel13, Button, Shift, X, Y)
End Sub
Private Sub CTBExercicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBExercicio, Source, X, Y)
End Sub
Private Sub CTBExercicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBExercicio, Button, Shift, X, Y)
End Sub
Private Sub CTBPeriodo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBPeriodo, Source, X, Y)
End Sub
Private Sub CTBPeriodo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBPeriodo, Button, Shift, X, Y)
End Sub
Private Sub CTBLabel14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel14, Source, X, Y)
End Sub
Private Sub CTBLabel14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel14, Button, Shift, X, Y)
End Sub
Private Sub CTBOrigem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBOrigem, Source, X, Y)
End Sub
Private Sub CTBOrigem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBOrigem, Button, Shift, X, Y)
End Sub
Private Sub CTBLabel21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel21, Source, X, Y)
End Sub
Private Sub CTBLabel21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel21, Button, Shift, X, Y)
End Sub
Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub
Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub
Public Function Form_Load_Ocx() As Object

    Call objCT.Form_Load_Ocx
    Set Form_Load_Ocx = Me

End Function

Public Sub Form_Unload(Cancel As Integer)
    If Not (objCT Is Nothing) Then
        Call objCT.Form_Unload(Cancel)
        If Cancel = False Then
             Set objCT.objUserControl = Nothing
             Set objCT = Nothing
        End If
    End If
End Sub

Private Sub objCT_Unload()
   RaiseEvent Unload
End Sub

Public Function Name() As String
    Call objCT.Name
End Function

Public Sub Show()
    Call objCT.Show
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

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Property Get Caption() As String
    Caption = objCT.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    objCT.Caption = New_Caption
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Call objCT.UserControl_KeyDown(KeyCode, Shift)
End Sub

'##########################################
'Inserido por Wagner
Private Sub BotaoImprimirPI_Click()
     Call objCT.BotaoImprimirPI_Click
End Sub
'##########################################

