VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ConcorrenciaConsOcx 
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9030
   ScaleHeight     =   5685
   ScaleWidth      =   9030
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   7290
      ScaleHeight     =   480
      ScaleWidth      =   1590
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   30
      Width           =   1650
      Begin VB.CommandButton BotaoBaixar 
         Height          =   360
         Left            =   585
         Picture         =   "ConcorrenciaConsOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Baixar"
         Top             =   72
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1110
         Picture         =   "ConcorrenciaConsOcx.ctx":01C2
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Fechar"
         Top             =   72
         Width           =   420
      End
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   75
         Picture         =   "ConcorrenciaConsOcx.ctx":0340
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Imprimir"
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4920
      Index           =   4
      Left            =   90
      TabIndex        =   4
      Top             =   675
      Visible         =   0   'False
      Width           =   8730
      Begin VB.Frame FrameCotacoes 
         Caption         =   "Cotações"
         Height          =   2610
         Index           =   2
         Left            =   168
         TabIndex        =   73
         Top             =   528
         Width           =   8505
         Begin MSMask.MaskEdBox ValorRecebido 
            Height          =   228
            Left            =   3384
            TabIndex        =   103
            Top             =   1440
            Visible         =   0   'False
            Width           =   1056
            _ExtentX        =   1852
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
            PromptChar      =   " "
         End
         Begin VB.ComboBox Moeda 
            Enabled         =   0   'False
            Height          =   288
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   101
            Top             =   1368
            Width           =   1665
         End
         Begin MSMask.MaskEdBox Taxa 
            Height          =   228
            Left            =   2124
            TabIndex        =   100
            Top             =   1404
            Visible         =   0   'False
            Width           =   1056
            _ExtentX        =   1852
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
            PromptChar      =   " "
         End
         Begin VB.ComboBox MotivoEscolhaCot 
            Enabled         =   0   'False
            Height          =   288
            Left            =   6360
            TabIndex        =   77
            Text            =   "MotivoEscolhaCot"
            Top             =   2145
            Width           =   1995
         End
         Begin VB.CheckBox EscolhidoCot 
            Enabled         =   0   'False
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
            Left            =   465
            TabIndex        =   76
            Top             =   240
            Width           =   840
         End
         Begin VB.TextBox DescProdutoCot 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   2520
            MaxLength       =   50
            TabIndex        =   75
            Top             =   270
            Width           =   1455
         End
         Begin VB.ComboBox TipoTributacaoCot 
            Enabled         =   0   'False
            Height          =   288
            Left            =   2220
            Style           =   2  'Dropdown List
            TabIndex        =   74
            Top             =   990
            Width           =   2565
         End
         Begin MSMask.MaskEdBox DataCotacao 
            Height          =   225
            Left            =   4230
            TabIndex        =   78
            Top             =   270
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            Format          =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox AliquotaICMS 
            Height          =   225
            Left            =   840
            TabIndex        =   79
            Top             =   135
            Width           =   1380
            _ExtentX        =   2434
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PrecoUnitario 
            Height          =   225
            Left            =   2925
            TabIndex        =   80
            Top             =   2310
            Width           =   1080
            _ExtentX        =   1905
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
            Format          =   "#,##0.00###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PedCotacao 
            Height          =   225
            Left            =   7260
            TabIndex        =   81
            Top             =   2175
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataValidade 
            Height          =   225
            Left            =   180
            TabIndex        =   82
            Top             =   2310
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   10
            Format          =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantidadeEntrega 
            Height          =   225
            Left            =   4935
            TabIndex        =   83
            Top             =   2235
            Width           =   1125
            _ExtentX        =   1984
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorPresente 
            Height          =   225
            Left            =   4095
            TabIndex        =   84
            Top             =   2250
            Width           =   1170
            _ExtentX        =   2064
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
         Begin MSMask.MaskEdBox DataNecessidade 
            Height          =   225
            Left            =   3450
            TabIndex        =   85
            Top             =   2115
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   10
            Format          =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataEntrega 
            Height          =   225
            Left            =   2430
            TabIndex        =   86
            Top             =   2325
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   10
            Format          =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PrazoEntrega 
            Height          =   225
            Left            =   1260
            TabIndex        =   87
            Top             =   2160
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   3
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantComprarCot 
            Height          =   225
            Left            =   6585
            TabIndex        =   88
            Top             =   2295
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CondPagto 
            Height          =   225
            Left            =   1245
            TabIndex        =   89
            Top             =   2310
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   30
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox UnidadeMedCot 
            Height          =   225
            Left            =   4005
            TabIndex        =   90
            Top             =   270
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   5
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FilialFornCot 
            Height          =   225
            Left            =   285
            TabIndex        =   91
            Top             =   2175
            Visible         =   0   'False
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FornecedorCot 
            Height          =   225
            Left            =   6195
            TabIndex        =   92
            Top             =   330
            Visible         =   0   'False
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantidadeCot 
            Height          =   225
            Left            =   5160
            TabIndex        =   93
            Top             =   270
            Width           =   1575
            _ExtentX        =   2778
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdutoCot 
            Height          =   225
            Left            =   1260
            TabIndex        =   94
            Top             =   270
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridCotacoes 
            Height          =   1845
            Left            =   150
            TabIndex        =   95
            Top             =   300
            Width           =   7980
            _ExtentX        =   14076
            _ExtentY        =   3254
            _Version        =   393216
            Rows            =   12
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox Preferencia 
            Height          =   225
            Left            =   6060
            TabIndex        =   96
            Top             =   2280
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorItem 
            Height          =   255
            Left            =   1980
            TabIndex        =   97
            Top             =   2145
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   450
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox AliquotaIPI 
            Height          =   225
            Left            =   15
            TabIndex        =   98
            Top             =   0
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
            PromptChar      =   " "
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Pedidos de Compra"
         Height          =   1584
         Left            =   168
         TabIndex        =   42
         Top             =   3204
         Width           =   8505
         Begin VB.ComboBox MoedaPC 
            Enabled         =   0   'False
            Height          =   288
            Left            =   6192
            Style           =   2  'Dropdown List
            TabIndex        =   102
            Top             =   792
            Width           =   1665
         End
         Begin MSMask.MaskEdBox ProdutoPC 
            Height          =   228
            Left            =   2700
            TabIndex        =   99
            Top             =   864
            Width           =   1212
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox UMPedido 
            Height          =   228
            Left            =   5148
            TabIndex        =   43
            Top             =   864
            Width           =   996
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantPedido 
            Height          =   228
            Left            =   4032
            TabIndex        =   44
            Top             =   864
            Width           =   996
            _ExtentX        =   1746
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodPedido 
            Height          =   228
            Left            =   1728
            TabIndex        =   45
            Top             =   864
            Width           =   876
            _ExtentX        =   1561
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
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridPedidos 
            Height          =   960
            Left            =   336
            TabIndex        =   46
            Top             =   300
            Width           =   7848
            _ExtentX        =   13864
            _ExtentY        =   1693
            _Version        =   393216
         End
      End
      Begin VB.CommandButton BotaoPedCotacao 
         Caption         =   "Pedido de Cotação"
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
         Left            =   6696
         TabIndex        =   41
         Top             =   144
         Width           =   1950
      End
      Begin VB.Label TaxaEmpresa 
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   1764
         TabIndex        =   6
         Top             =   132
         Width           =   1140
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Taxa Financeira:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   228
         TabIndex        =   5
         Top             =   168
         Width           =   1440
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4860
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   705
      Width           =   8655
      Begin VB.Frame Frame2 
         Caption         =   "Concorrência"
         Height          =   870
         Left            =   195
         TabIndex        =   68
         Top             =   75
         Width           =   8160
         Begin VB.Label Label5 
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
            Height          =   195
            Left            =   2970
            TabIndex        =   72
            Top             =   465
            Width           =   930
         End
         Begin VB.Label DescricaoLabel 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4035
            TabIndex        =   71
            Top             =   435
            Width           =   3690
         End
         Begin VB.Label LabelConcorrencia 
            AutoSize        =   -1  'True
            Caption         =   "Concorrência:"
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
            Left            =   165
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   70
            Top             =   450
            Width           =   1200
         End
         Begin VB.Label ConcorrenciaLabel 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1425
            TabIndex        =   69
            Top             =   420
            Width           =   1140
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Requisições de Compra"
         Height          =   3555
         Left            =   195
         TabIndex        =   7
         Top             =   1110
         Width           =   8160
         Begin VB.TextBox ObservacaoReq 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   1245
            MaxLength       =   255
            TabIndex        =   9
            Top             =   4170
            Width           =   2415
         End
         Begin VB.CheckBox UrgenteReq 
            Enabled         =   0   'False
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
            Left            =   5010
            TabIndex        =   8
            Top             =   315
            Width           =   870
         End
         Begin MSMask.MaskEdBox FilialEmpresaReq 
            Height          =   225
            Left            =   285
            TabIndex        =   10
            Top             =   330
            Visible         =   0   'False
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Requisitante 
            Height          =   240
            Left            =   5160
            TabIndex        =   11
            Top             =   315
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CclReq 
            Height          =   225
            Left            =   225
            TabIndex        =   12
            Top             =   4170
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataLimite 
            Height          =   225
            Left            =   2655
            TabIndex        =   13
            Top             =   315
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodigoRequisicao 
            Height          =   225
            Left            =   1845
            TabIndex        =   14
            Top             =   315
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataReq 
            Height          =   225
            Left            =   3825
            TabIndex        =   15
            Top             =   330
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridRequisicoes 
            Height          =   3000
            Left            =   180
            TabIndex        =   16
            Top             =   375
            Width           =   7785
            _ExtentX        =   13732
            _ExtentY        =   5292
            _Version        =   393216
            Rows            =   16
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4890
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   705
      Visible         =   0   'False
      Width           =   8640
      Begin VB.Frame FrameProdutos 
         BorderStyle     =   0  'None
         Height          =   3930
         Index           =   1
         Left            =   195
         TabIndex        =   58
         Top             =   405
         Width           =   8250
         Begin VB.TextBox DescProduto1 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   2520
            MaxLength       =   50
            TabIndex        =   59
            Top             =   255
            Width           =   1455
         End
         Begin MSMask.MaskEdBox UnidadeMed1 
            Height          =   225
            Left            =   4005
            TabIndex        =   60
            Top             =   315
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   5
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FilialForn1 
            Height          =   165
            Left            =   3030
            TabIndex        =   61
            Top             =   3540
            Visible         =   0   'False
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   291
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Fornecedor1 
            Height          =   165
            Left            =   1080
            TabIndex        =   62
            Top             =   3540
            Visible         =   0   'False
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   291
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantComprar1 
            Height          =   225
            Left            =   5145
            TabIndex        =   63
            Top             =   315
            Width           =   990
            _ExtentX        =   1746
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Produto1 
            Height          =   225
            Left            =   1230
            TabIndex        =   64
            Top             =   270
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantUrgente1 
            Height          =   225
            Left            =   6270
            TabIndex        =   65
            Top             =   255
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridProdutos1 
            Height          =   3255
            Left            =   240
            TabIndex        =   66
            Top             =   285
            Width           =   7395
            _ExtentX        =   13044
            _ExtentY        =   5741
            _Version        =   393216
            Rows            =   12
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.Frame FrameProdutos 
         BorderStyle     =   0  'None
         Caption         =   "Produtos"
         Height          =   3885
         Index           =   2
         Left            =   180
         TabIndex        =   47
         Top             =   405
         Visible         =   0   'False
         Width           =   8205
         Begin VB.TextBox DescProduto2 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   2430
            MaxLength       =   50
            TabIndex        =   48
            Top             =   285
            Width           =   1455
         End
         Begin MSMask.MaskEdBox FilialDestino 
            Height          =   225
            Left            =   645
            TabIndex        =   49
            Top             =   3255
            Width           =   1065
            _ExtentX        =   1879
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Destino 
            Height          =   225
            Left            =   7035
            TabIndex        =   50
            Top             =   300
            Width           =   1065
            _ExtentX        =   1879
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TipDestino 
            Height          =   225
            Left            =   6000
            TabIndex        =   51
            Top             =   315
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Fornecedor2 
            Height          =   225
            Left            =   1770
            TabIndex        =   52
            Top             =   3300
            Visible         =   0   'False
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox UnidadeMed2 
            Height          =   225
            Left            =   3945
            TabIndex        =   53
            Top             =   270
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   5
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FilialForn2 
            Height          =   225
            Left            =   3735
            TabIndex        =   54
            Top             =   3300
            Visible         =   0   'False
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantComprar2 
            Height          =   225
            Left            =   5010
            TabIndex        =   55
            Top             =   270
            Width           =   990
            _ExtentX        =   1746
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Produto2 
            Height          =   225
            Left            =   825
            TabIndex        =   56
            Top             =   285
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridProdutos2 
            Height          =   3255
            Left            =   120
            TabIndex        =   57
            Top             =   210
            Width           =   7650
            _ExtentX        =   13494
            _ExtentY        =   5741
            _Version        =   393216
            Rows            =   12
            Cols            =   9
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.CommandButton BotaoEditarProduto 
         Caption         =   "Produto ..."
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
         Left            =   7095
         TabIndex        =   17
         Top             =   4485
         Width           =   1395
      End
      Begin MSComctlLib.TabStrip TabProdutos 
         Height          =   4365
         Left            =   135
         TabIndex        =   67
         Top             =   90
         Width           =   8400
         _ExtentX        =   14817
         _ExtentY        =   7699
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Seleção"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Quantidades por Destino"
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4875
      Index           =   2
      Left            =   135
      TabIndex        =   2
      Top             =   690
      Visible         =   0   'False
      Width           =   8730
      Begin VB.Frame Frame5 
         Caption         =   "Itens de Requisições"
         Height          =   4380
         Left            =   195
         TabIndex        =   18
         Top             =   180
         Width           =   8295
         Begin VB.TextBox DescProdutoItemReq 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   4500
            MaxLength       =   50
            TabIndex        =   21
            Top             =   450
            Width           =   1455
         End
         Begin VB.CheckBox EscolhidoItemReq 
            Enabled         =   0   'False
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
            Left            =   180
            TabIndex        =   20
            Top             =   465
            Width           =   915
         End
         Begin VB.TextBox ObservacaoItemReq 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   5310
            MaxLength       =   255
            TabIndex        =   19
            Top             =   3915
            Width           =   2355
         End
         Begin MSMask.MaskEdBox FilialItemReq 
            Height          =   225
            Left            =   885
            TabIndex        =   22
            Top             =   585
            Visible         =   0   'False
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CclItemReq 
            Height          =   225
            Left            =   2520
            TabIndex        =   23
            Top             =   4005
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Item 
            Height          =   225
            Left            =   3090
            TabIndex        =   24
            Top             =   555
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   3
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodigoReq 
            Height          =   225
            Left            =   2250
            TabIndex        =   25
            Top             =   405
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ExclusivoItemReq 
            Height          =   225
            Left            =   6345
            TabIndex        =   26
            Top             =   4020
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Almoxarifado 
            Height          =   225
            Left            =   1200
            TabIndex        =   27
            Top             =   4005
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox UMItemReq 
            Height          =   225
            Left            =   6480
            TabIndex        =   28
            Top             =   345
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   5
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantComprarItemReq 
            Height          =   225
            Left            =   7035
            TabIndex        =   29
            Top             =   330
            Width           =   990
            _ExtentX        =   1746
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantRecebida 
            Height          =   225
            Left            =   165
            TabIndex        =   30
            Top             =   4005
            Width           =   990
            _ExtentX        =   1746
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantPedida 
            Height          =   225
            Left            =   1725
            TabIndex        =   31
            Top             =   4050
            Width           =   990
            _ExtentX        =   1746
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FilialFornItemReq 
            Height          =   225
            Left            =   5055
            TabIndex        =   32
            Top             =   4005
            Visible         =   0   'False
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FornecedorItemReq 
            Height          =   225
            Left            =   3315
            TabIndex        =   33
            Top             =   4005
            Visible         =   0   'False
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantidadeItemRC 
            Height          =   225
            Left            =   600
            TabIndex        =   34
            Top             =   4050
            Width           =   990
            _ExtentX        =   1746
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdutoItemReq 
            Height          =   225
            Left            =   3435
            TabIndex        =   35
            Top             =   345
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridItensReq 
            Height          =   3135
            Left            =   285
            TabIndex        =   36
            Top             =   630
            Width           =   7665
            _ExtentX        =   13520
            _ExtentY        =   5530
            _Version        =   393216
            Rows            =   15
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5295
      Left            =   75
      TabIndex        =   0
      Top             =   345
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9340
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Requisições"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Itens de Requisições"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Produtos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Cotações - Pedidos"
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
Attribute VB_Name = "ConcorrenciaConsOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Variáveis globais
Dim iFrameAtual As Integer
Dim iFrameProdutoAtual As Integer
Dim iAlterado As Integer
Dim iFrameTipoDestinoAtual As Integer
Dim gsTipoTributacao As String
Dim gcolRequisicaoCompra As Collection

'GridRequisicoes
Dim objGridRequisicoes As AdmGrid
Dim iGrid_FilialEmpresaReq_Col As Integer
Dim iGrid_CodigoReq_Col  As Integer
Dim iGrid_DataLimiteReq_Col As Integer
Dim iGrid_DataReq_Col As Integer
Dim iGrid_UrgenteReq_Col As Integer
Dim iGrid_RequisitanteReq_Col As Integer
Dim iGrid_CclReq_Col As Integer
Dim iGrid_ObservacaoReq_Col As Integer
Dim iGrid_TipoDestinoReq_Col As Integer
Dim iGrid_DestinoReq_Col As Integer
Dim iGrid_FilialDestReq_Col As Integer

'GridItensReq
Dim objGridItensReq As AdmGrid
Dim iGrid_EscolhidoItemReq_Col As Integer
Dim iGrid_FilialItemReq_Col As Integer
Dim iGrid_CodigoReqItemReq_Col As Integer
Dim iGrid_ItemItemReq_Col As Integer
Dim iGrid_ProdutoItemReq_Col As Integer
Dim iGrid_DescProdutoItemReq_Col As Integer
Dim iGrid_UMItemReq_Col As Integer
Dim iGrid_QuantComprarItemReq_Col As Integer
Dim iGrid_QuantidadeItemReq_Col As Integer
Dim iGrid_QuantPedidaItemReq_Col As Integer
Dim iGrid_QuantRecebidaItemReq_Col As Integer
Dim iGrid_AlmoxarifadoItemReq_Col As Integer
Dim iGrid_CclItemReq_Col As Integer
Dim iGrid_FornecedorItemReq_Col As Integer
Dim iGrid_FilialFornItemReq_Col As Integer
Dim iGrid_ExclusivoItemReq_Col As Integer
Dim iGrid_ObservacaoItemReq_Col As Integer

'GridProdutos1
Dim objGridProdutos1 As AdmGrid
Dim iGrid_Produto1_Col As Integer
Dim iGrid_DescProduto1_Col As Integer
Dim iGrid_UnidadeMed1_Col As Integer
Dim iGrid_QuantComprar1_Col As Integer
Dim iGrid_Urgente1_Col As Integer
Dim iGrid_Fornecedor1_Col As Integer
Dim iGrid_FilialForn1_Col As Integer

'GridProdutos2
Dim objGridProdutos2 As AdmGrid
Dim iGrid_Escolhido2_Col As Integer
Dim iGrid_Produto2_Col As Integer
Dim iGrid_DescProduto2_Col As Integer
Dim iGrid_UnidadeMed2_Col As Integer
Dim iGrid_QuantComprar2_Col As Integer
Dim iGrid_Urgente2_Col As Integer
Dim iGrid_TipoDestino_Col As Integer
Dim iGrid_Destino_Col As Integer
Dim iGrid_FilialDestino_Col As Integer
Dim iGrid_Fornecedor2_Col As Integer
Dim iGrid_FilialForn2_Col As Integer

'GridCotacoes
Dim objGridCotacoes As AdmGrid
Dim iGrid_EscolhidoCot_Col As Integer
Dim iGrid_ProdutoCot_Col As Integer
Dim iGrid_DescProdutoCot_Col As Integer
Dim iGrid_UMCot_Col As Integer
Dim iGrid_QuantidadeCot_Col As Integer
Dim iGrid_FornecedorCot_Col As Integer
Dim iGrid_FilialFornCot_Col As Integer
Dim iGrid_CondPagtoCot_Col As Integer
Dim iGrid_PrecoUnitario_Col As Integer
Dim iGrid_ValorPresenteCot_Col As Integer
Dim iGrid_TipoTributacaoCot_Col As Integer
Dim iGrid_AliquotaIPI_Col As Integer
Dim iGrid_AliquotaICMS_Col As Integer
Dim iGrid_PedidoCot_Col As Integer
Dim iGrid_DataValidadeCot_Col As Integer
Dim iGrid_PrazoEntrega_Col As Integer
Dim iGrid_DataNecessidade_Col As Integer
Dim iGrid_QuantidadeEntrega_Col As Integer
Dim iGrid_Preferencia_Col As Integer
Dim iGrid_QuantComprarCot_Col As Integer
Dim iGrid_MotivoEscolhaCot_Col As Integer
Dim iGrid_DataEntrega_Col As Integer
Dim iGrid_ValorItem_Col As Integer
Dim iGrid_DataCotacaoCot_Col As Integer
Dim iGrid_MoedaCot_Col As Integer
Dim iGrid_TaxaCot_Col As Integer
Dim iGrid_Unitario_RS_Col As Integer

'GridPedidos
Dim objGridPedidos As AdmGrid
Dim iGrid_CodPedido_Col As Integer
Dim iGrid_ProdutoPC_Col As Integer
Dim iGrid_QuantPedido_Col As Integer
Dim iGrid_UMPedido_Col As Integer
Dim iGrid_MoedaPC_Col As Integer

Dim gobjConcorrencia As ClassConcorrencia

'Eventos dos Browses
Private WithEvents objEventoConcorrencia As AdmEvento
Attribute objEventoConcorrencia.VB_VarHelpID = -1

Function Trata_Parametros(Optional objConcorrencia As ClassConcorrencia)

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not objConcorrencia Is Nothing Then

        Set gobjConcorrencia = objConcorrencia
        lErro = Traz_Concorrencia_Tela(objConcorrencia)
        If lErro <> SUCESSO Then gError 74852
    End If

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 74852

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154556)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long
Dim objComprador As New ClassComprador
Dim colCod_Descricao As AdmColCodigoNome
Dim colCod_DescReduzida As New AdmColCodigoNome
Dim lCotacao As Long
Dim objUsuario As New ClassUsuario
Dim objConfiguraCOM As New ClassConfiguraCOM
Dim iTipoTrib As Integer
Dim sDescricao As String
Dim colCodigoDescricao As New AdmColCodigoNome

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    iFrameProdutoAtual = 1
    
    '##########################
    'Inserido por Wagner
    Call Formata_Controles
    '##########################
    
    'Carrega a combo de moedas ...
    lErro = Carrega_Moeda
    If lErro <> SUCESSO Then gError 114533
    
    Set gobjConcorrencia = New ClassConcorrencia
    
    Set objGridPedidos = New AdmGrid
    Set objGridRequisicoes = New AdmGrid
    Set objGridItensReq = New AdmGrid
    Set objGridProdutos1 = New AdmGrid
    Set objGridProdutos2 = New AdmGrid
    Set objGridCotacoes = New AdmGrid
    Set objEventoConcorrencia = New AdmEvento

    'Inicializa coleção de Requisição de Compras
    Set gcolRequisicaoCompra = New Collection

    'Inicializa as máscaras dos Produtos
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoItemReq)
    If lErro <> SUCESSO Then gError 67644
    
    Produto1.Mask = ProdutoItemReq.Mask
    Produto2.Mask = ProdutoItemReq.Mask
    ProdutoCot.Mask = ProdutoItemReq.Mask

    'Inicializa mascara do Ccl
    lErro = Inicializa_MascaraCcl()
    If lErro <> SUCESSO Then gError 67648

    'Coloca as Quantidades da tela no formato de Estoque
    QuantComprarItemReq.Format = FORMATO_ESTOQUE
    QuantidadeItemRC.Format = FORMATO_ESTOQUE
    QuantPedida.Format = FORMATO_ESTOQUE
    QuantRecebida.Format = FORMATO_ESTOQUE
    QuantComprar1.Format = FORMATO_ESTOQUE
    QuantComprar2.Format = FORMATO_ESTOQUE
    QuantComprarCot.Format = FORMATO_ESTOQUE

    'Carrega Motivos de Escolha
    lErro = Carrega_MotivoEscolhaCot()
    If lErro <> SUCESSO Then gError 67649

    'Inicializa o GridRequisicoes
    lErro = Inicializa_Grid_Requisicoes(objGridRequisicoes)
    If lErro <> SUCESSO Then gError 67650

    'Inicializa o GridItensReq
    lErro = Inicializa_Grid_ItensReq(objGridItensReq)
    If lErro <> SUCESSO Then gError 67652

    'Inicializa o GridProdutos1
    lErro = Inicializa_Grid_Produtos1(objGridProdutos1)
    If lErro <> SUCESSO Then gError 67653

    'Inicializa o GridProdutos2
    lErro = Inicializa_Grid_Produtos2(objGridProdutos2)
    If lErro <> SUCESSO Then gError 67654

    'Inicializa o GridProdutos2
    lErro = Inicializa_Grid_Cotacoes(objGridCotacoes)
    If lErro <> SUCESSO Then gError 67655

    'Inicializa o GridPedidos
    lErro = Inicializa_Grid_Pedidos(objGridPedidos)
    If lErro <> SUCESSO Then gError 67651

    lErro = Carrega_TipoTributacao()
    If lErro <> SUCESSO Then gError 67657

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case gErr

        Case 67644 To 67657, 114533
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154557)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Private Function Inicializa_MascaraCcl() As Long
'Inicializa a mascara do centro de custo

Dim sMascaraCcl As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_mascaraccl

    sMascaraCcl = String(STRING_CCL, 0)

    'Le a mascara dos centros de custo/lucro
    lErro = MascaraCcl(sMascaraCcl)
    If lErro <> SUCESSO Then gError 67658

    CclReq.Mask = sMascaraCcl
    CclItemReq.Mask = sMascaraCcl

    Inicializa_MascaraCcl = SUCESSO

    Exit Function

Erro_Inicializa_mascaraccl:

    Inicializa_MascaraCcl = gErr

    Select Case gErr

        Case 67658
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154558)

    End Select

    Exit Function

End Function

Private Function Carrega_MotivoEscolhaCot() As Long
'Carrega a combobox FilialEmpresa

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_MotivoEscolhaCot

    'Lê o Código e o Nome de todo MotivoEscolhaCot do BD
    lErro = CF("Cod_Nomes_Le", "Motivo", "Codigo", "Motivo", STRING_NOME_TABELA, colCodigoNome)
    If lErro <> SUCESSO Then gError 67659

    'Carrega a combo de Motivo Escolha com código e nome
    For Each objCodigoNome In colCodigoNome

        'Verifica se o MotivoEscolhaCot é diferente de Exclusividade
        If objCodigoNome.iCodigo <> MOTIVO_EXCLUSIVO Then

            MotivoEscolhaCot.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
            MotivoEscolhaCot.ItemData(MotivoEscolhaCot.NewIndex) = objCodigoNome.iCodigo

        End If

    Next

    Carrega_MotivoEscolhaCot = SUCESSO

    Exit Function

Erro_Carrega_MotivoEscolhaCot:

    Carrega_MotivoEscolhaCot = gErr

    Select Case gErr

        Case 67659
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154559)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Requisicoes(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Requisições

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Requisicoes

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")

    'Para versao Full
    If giTipoVersao = VERSAO_FULL Then
        objGridInt.colColuna.Add ("Filial Empresa")
    End If

    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Data Limite")
    objGridInt.colColuna.Add ("Data RC")
    objGridInt.colColuna.Add ("Urgente")
    objGridInt.colColuna.Add ("Requisitante")
    objGridInt.colColuna.Add ("Ccl")
    objGridInt.colColuna.Add ("Observação")

    'campos de edição do grid

    'Para versao Full
    If giTipoVersao = VERSAO_FULL Then
        objGridInt.colCampo.Add (FilialEmpresaReq.Name)
    End If

    objGridInt.colCampo.Add (CodigoRequisicao.Name)
    objGridInt.colCampo.Add (DataLimite.Name)
    objGridInt.colCampo.Add (DataReq.Name)
    objGridInt.colCampo.Add (UrgenteReq.Name)
    objGridInt.colCampo.Add (Requisitante.Name)
    objGridInt.colCampo.Add (CclReq.Name)
    objGridInt.colCampo.Add (ObservacaoReq.Name)

    'indica onde estao situadas as colunas do grid para versao Full
    If giTipoVersao = VERSAO_FULL Then

        iGrid_FilialEmpresaReq_Col = 1
        iGrid_CodigoReq_Col = 2
        iGrid_DataLimiteReq_Col = 3
        iGrid_DataReq_Col = 4
        iGrid_UrgenteReq_Col = 5
        iGrid_RequisitanteReq_Col = 6
        iGrid_CclReq_Col = 7
        iGrid_ObservacaoReq_Col = 8

    End If

    'indica onde estao situadas as colunas do grid para Versao Light
    If giTipoVersao = VERSAO_LIGHT Then

        iGrid_CodigoReq_Col = 1
        iGrid_DataLimiteReq_Col = 2
        iGrid_DataReq_Col = 3
        iGrid_UrgenteReq_Col = 4
        iGrid_RequisitanteReq_Col = 5
        iGrid_CclReq_Col = 6
        iGrid_ObservacaoReq_Col = 7

    End If

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridRequisicoes

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_REQUISICOES + 1

    'Não permite incluir e excluir linhas do grid
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 8

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Requisicoes = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Requisicoes:

    Inicializa_Grid_Requisicoes = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 154560)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_ItensReq(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Itens de Requisições

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_ItensReq

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Escolhido")

    'Para a Versao Full
    If giTipoVersao = VERSAO_FULL Then
        objGridInt.colColuna.Add ("Filial Empresa")
    End If

    objGridInt.colColuna.Add ("Requisição")
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("A Comprar")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Em Pedido")
    objGridInt.colColuna.Add ("Recebido")
    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("Ccl")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial Fornecedor")
    objGridInt.colColuna.Add ("Exclusivo")
    objGridInt.colColuna.Add ("Observação")

    'campos de edição do grid
    objGridInt.colCampo.Add (EscolhidoItemReq.Name)
    'Para a versao full
    If giTipoVersao = VERSAO_FULL Then
        objGridInt.colCampo.Add (FilialItemReq.Name)
    End If
    objGridInt.colCampo.Add (CodigoReq.Name)
    objGridInt.colCampo.Add (Item.Name)
    objGridInt.colCampo.Add (ProdutoItemReq.Name)
    objGridInt.colCampo.Add (DescProdutoItemReq.Name)
    objGridInt.colCampo.Add (UMItemReq.Name)
    objGridInt.colCampo.Add (QuantComprarItemReq.Name)
    objGridInt.colCampo.Add (QuantidadeItemRC.Name)
    objGridInt.colCampo.Add (QuantPedida.Name)
    objGridInt.colCampo.Add (QuantRecebida.Name)
    objGridInt.colCampo.Add (Almoxarifado.Name)
    objGridInt.colCampo.Add (CclItemReq.Name)
    objGridInt.colCampo.Add (FornecedorItemReq.Name)
    objGridInt.colCampo.Add (FilialFornItemReq.Name)
    objGridInt.colCampo.Add (ExclusivoItemReq.Name)
    objGridInt.colCampo.Add (ObservacaoItemReq.Name)

    'indica onde estao situadas as colunas do grid, para a versao FULL
    If giTipoVersao = VERSAO_FULL Then

        iGrid_EscolhidoItemReq_Col = 1
        iGrid_FilialItemReq_Col = 2
        iGrid_CodigoReqItemReq_Col = 3
        iGrid_ItemItemReq_Col = 4
        iGrid_ProdutoItemReq_Col = 5
        iGrid_DescProdutoItemReq_Col = 6
        iGrid_UMItemReq_Col = 7
        iGrid_QuantComprarItemReq_Col = 8
        iGrid_QuantidadeItemReq_Col = 9
        iGrid_QuantPedidaItemReq_Col = 10
        iGrid_QuantRecebidaItemReq_Col = 11
        iGrid_AlmoxarifadoItemReq_Col = 12
        iGrid_CclItemReq_Col = 13
        iGrid_FornecedorItemReq_Col = 14
        iGrid_FilialFornItemReq_Col = 15
        iGrid_ExclusivoItemReq_Col = 16
        iGrid_ObservacaoItemReq_Col = 17

    End If

    'indica onde estao situadas as colunas do grid, para a versao LIGHT
    If giTipoVersao = VERSAO_LIGHT Then

        iGrid_EscolhidoItemReq_Col = 1
        iGrid_CodigoReqItemReq_Col = 2
        iGrid_ItemItemReq_Col = 3
        iGrid_ProdutoItemReq_Col = 4
        iGrid_DescProdutoItemReq_Col = 5
        iGrid_UMItemReq_Col = 6
        iGrid_QuantComprarItemReq_Col = 7
        iGrid_QuantidadeItemReq_Col = 8
        iGrid_QuantPedidaItemReq_Col = 9
        iGrid_QuantRecebidaItemReq_Col = 10
        iGrid_AlmoxarifadoItemReq_Col = 11
        iGrid_CclItemReq_Col = 12
        iGrid_FornecedorItemReq_Col = 13
        iGrid_FilialFornItemReq_Col = 14
        iGrid_ExclusivoItemReq_Col = 15
        iGrid_ObservacaoItemReq_Col = 16

    End If

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridItensReq

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_REQUISICAO + 1

    'Não permite incluir e excluir linhas do grid
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 10

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_ItensReq = SUCESSO

    Exit Function

Erro_Inicializa_Grid_ItensReq:

    Inicializa_Grid_ItensReq = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 154561)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Produtos1(objGridInt As AdmGrid) As Long
'Executa a Inicialização do Grid Produtos1

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Produtos1

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("A Comprar")
    objGridInt.colColuna.Add ("Urgente")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial Fornecedor")

    'campos de edição do grid
    objGridInt.colCampo.Add (Produto1.Name)
    objGridInt.colCampo.Add (DescProduto1.Name)
    objGridInt.colCampo.Add (UnidadeMed1.Name)
    objGridInt.colCampo.Add (QuantComprar1.Name)
    objGridInt.colCampo.Add (QuantUrgente1.Name)
    objGridInt.colCampo.Add (Fornecedor1.Name)
    objGridInt.colCampo.Add (FilialForn1.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_Produto1_Col = 1
    iGrid_DescProduto1_Col = 2
    iGrid_UnidadeMed1_Col = 3
    iGrid_QuantComprar1_Col = 4
    iGrid_Urgente1_Col = 5
    iGrid_Fornecedor1_Col = 6
    iGrid_FilialForn1_Col = 7

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridProdutos1

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_GERACAO + 1

    'Não permite incluir e excluir linhas do grid
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 10

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Produtos1 = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Produtos1:

    Inicializa_Grid_Produtos1 = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 154562)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Produtos2(objGridInt As AdmGrid) As Long
'Executa a Inicialização do Grid Produtos2

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Produtos2

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("A Comprar")
    objGridInt.colColuna.Add ("Tipo Destino")
    objGridInt.colColuna.Add ("Destino")
    objGridInt.colColuna.Add ("FilialDestino")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial Fornecedor")

    'campos de edição do grid
    objGridInt.colCampo.Add (Produto2.Name)
    objGridInt.colCampo.Add (DescProduto2.Name)
    objGridInt.colCampo.Add (UnidadeMed2.Name)
    objGridInt.colCampo.Add (QuantComprar2.Name)
    objGridInt.colCampo.Add (TipDestino.Name)
    objGridInt.colCampo.Add (Destino.Name)
    objGridInt.colCampo.Add (FilialDestino.Name)
    objGridInt.colCampo.Add (Fornecedor2.Name)
    objGridInt.colCampo.Add (FilialForn2.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_Produto2_Col = 1
    iGrid_DescProduto2_Col = 2
    iGrid_UnidadeMed2_Col = 3
    iGrid_QuantComprar2_Col = 4
    iGrid_TipoDestino_Col = 5
    iGrid_Destino_Col = 6
    iGrid_FilialDestino_Col = 7
    iGrid_Fornecedor2_Col = 8
    iGrid_FilialForn2_Col = 9

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridProdutos2

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_GERACAO + 1

    'Não permite incluir e excluir linhas do grid
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 10

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Produtos2 = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Produtos2:

    Inicializa_Grid_Produtos2 = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 154563)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Cotacoes(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Cotacoes

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Cotacoes

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Escolhido")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Preferência")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial Fornecedor")
    objGridInt.colColuna.Add ("Cond. Pagto")
    objGridInt.colColuna.Add ("Quant. Cotada")
    objGridInt.colColuna.Add ("A Comprar")
    objGridInt.colColuna.Add ("UM")
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("Moeda")
    objGridInt.colColuna.Add ("Taxa")
    objGridInt.colColuna.Add ("Unitário R$")
    objGridInt.colColuna.Add ("Valor Presente")
    objGridInt.colColuna.Add ("Valor Item")
    objGridInt.colColuna.Add ("Tipo Tributacao")
    objGridInt.colColuna.Add ("Alíquota IPI")
    objGridInt.colColuna.Add ("Alíquota ICMS")
    objGridInt.colColuna.Add ("Ped. Cotação")
    objGridInt.colColuna.Add ("Data Cotação")
    objGridInt.colColuna.Add ("Data Validade")
    objGridInt.colColuna.Add ("Prazo Entrega")
    objGridInt.colColuna.Add ("Data Entrega")
    objGridInt.colColuna.Add ("Data Necessidade")
    objGridInt.colColuna.Add ("Para Entrega")
    objGridInt.colColuna.Add ("Motivo da Escolha")

    'campos de edição do grid
    objGridInt.colCampo.Add (EscolhidoCot.Name)
    objGridInt.colCampo.Add (ProdutoCot.Name)
    objGridInt.colCampo.Add (DescProdutoCot.Name)
    objGridInt.colCampo.Add (Preferencia.Name)
    objGridInt.colCampo.Add (FornecedorCot.Name)
    objGridInt.colCampo.Add (FilialFornCot.Name)
    objGridInt.colCampo.Add (CondPagto.Name)
    objGridInt.colCampo.Add (QuantidadeCot.Name)
    objGridInt.colCampo.Add (QuantComprarCot.Name)
    objGridInt.colCampo.Add (UnidadeMedCot.Name)
    objGridInt.colCampo.Add (PrecoUnitario.Name)
    objGridInt.colCampo.Add (Moeda.Name)
    objGridInt.colCampo.Add (Taxa.Name)
    objGridInt.colCampo.Add (ValorRecebido.Name)
    objGridInt.colCampo.Add (ValorPresente.Name)
    objGridInt.colCampo.Add (ValorItem.Name)
    objGridInt.colCampo.Add (TipoTributacaoCot.Name)
    objGridInt.colCampo.Add (AliquotaIPI.Name)
    objGridInt.colCampo.Add (AliquotaICMS.Name)
    objGridInt.colCampo.Add (PedCotacao.Name)
    objGridInt.colCampo.Add (DataCotacao.Name)
    objGridInt.colCampo.Add (DataValidade.Name)
    objGridInt.colCampo.Add (PrazoEntrega.Name)
    objGridInt.colCampo.Add (DataEntrega.Name)
    objGridInt.colCampo.Add (DataNecessidade.Name)
    objGridInt.colCampo.Add (QuantidadeEntrega.Name)
    objGridInt.colCampo.Add (MotivoEscolhaCot.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_EscolhidoCot_Col = 1
    iGrid_ProdutoCot_Col = 2
    iGrid_DescProdutoCot_Col = 3
    iGrid_Preferencia_Col = 4
    iGrid_FornecedorCot_Col = 5
    iGrid_FilialFornCot_Col = 6
    iGrid_CondPagtoCot_Col = 7
    iGrid_QuantidadeCot_Col = 8
    iGrid_QuantComprarCot_Col = 9
    iGrid_UMCot_Col = 10
    iGrid_PrecoUnitario_Col = 11
    iGrid_MoedaCot_Col = 12
    iGrid_TaxaCot_Col = 13
    iGrid_Unitario_RS_Col = 14
    iGrid_ValorPresenteCot_Col = 15
    iGrid_ValorItem_Col = 16
    iGrid_TipoTributacaoCot_Col = 17
    iGrid_AliquotaIPI_Col = 18
    iGrid_AliquotaICMS_Col = 19
    iGrid_PedidoCot_Col = 20
    iGrid_DataCotacaoCot_Col = 21
    iGrid_DataValidadeCot_Col = 22
    iGrid_PrazoEntrega_Col = 23
    iGrid_DataEntrega_Col = 24
    iGrid_DataNecessidade_Col = 25
    iGrid_QuantidadeEntrega_Col = 26
    iGrid_MotivoEscolhaCot_Col = 27

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridCotacoes

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_COTACOES + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 4

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    GridCotacoes.Width = 8295
    GridCotacoes.ColWidth(0) = 350

    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Cotacoes = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Cotacoes:

    Inicializa_Grid_Cotacoes = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 154564)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Pedidos(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Concorrências

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Pedidos

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Código")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Moeda")

    'campos de edição do grid
    objGridInt.colCampo.Add (CodPedido.Name)
    objGridInt.colCampo.Add (ProdutoPC.Name)
    objGridInt.colCampo.Add (QuantPedido.Name)
    objGridInt.colCampo.Add (UMPedido.Name)
    objGridInt.colCampo.Add (MoedaPC.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_CodPedido_Col = 1
    iGrid_ProdutoPC_Col = 2
    iGrid_QuantPedido_Col = 3
    iGrid_UMPedido_Col = 4
    iGrid_MoedaPC_Col = 5

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridPedidos

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_PEDIDOS + 1

    'Não permite incluir e excluir linhas do grid
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 2

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Pedidos = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Pedidos:

    Inicializa_Grid_Pedidos = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 154565)

    End Select

    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objConcorrencia As New ClassConcorrencia

On Error GoTo Erro_Tela_Extrai

    sTabela = "ConcorrenciaTodas"

    'Move os dados da tela para a memoria
    lErro = Move_Tela_Memoria(objConcorrencia)
    If lErro <> SUCESSO Then gError 67660

    'Preenche a coleção colCampoValor
    colCampoValor.Add "Codigo", objConcorrencia.lCodigo, 0, "Codigo"
    colCampoValor.Add "TaxaFinanceira", objConcorrencia.dTaxaFinanceira, 0, "TaxaFinanceira"
    colCampoValor.Add "FilialEmpresa", objConcorrencia.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "NumIntDoc", objConcorrencia.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "Descricao", objConcorrencia.sDescricao, STRING_DESCRICAO_CAMPO, "Descricao"

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 67660

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154566)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objConcorrencia As New ClassConcorrencia

On Error GoTo Erro_Tela_Preenche

    'Carrega objPedidoCompra com os dados passados em colCampoValor
    objConcorrencia.lCodigo = colCampoValor.Item("Codigo").vValor
    objConcorrencia.dTaxaFinanceira = colCampoValor.Item("TaxaFinanceira").vValor
    objConcorrencia.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
    objConcorrencia.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor
    objConcorrencia.sDescricao = colCampoValor.Item("Descricao").vValor

    'Traz os dados da Concorrência para a tela
    lErro = Traz_Concorrencia_Tela(objConcorrencia)
    If lErro <> SUCESSO Then gError 67661

    iAlterado = 0

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 67661

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154567)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    'Libera objetos globais
    Set objGridPedidos = Nothing
    Set objGridRequisicoes = Nothing
    Set objGridItensReq = Nothing
    Set objGridProdutos1 = Nothing
    Set objGridProdutos2 = Nothing
    Set objGridCotacoes = Nothing

    Set objEventoConcorrencia = Nothing

    Set gcolRequisicaoCompra = Nothing
    Set gobjConcorrencia = Nothing
    Call ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub BotaoImprimir_Click()

Dim lErro As Long
Dim objConcorrencia As New ClassConcorrencia
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_BotaoImprimir_Click

    'Verifica se o código da concorrencia esta preenchido
    If Len(Trim(ConcorrenciaLabel.Caption)) = 0 Then gError 76084

    objConcorrencia.lCodigo = StrParaLong(ConcorrenciaLabel.Caption)
    objConcorrencia.iFilialEmpresa = giFilialEmpresa

    'Lê a Concorrencia
    lErro = CF("ConcorrenciaN_Le", objConcorrencia)
    If lErro <> SUCESSO And lErro <> 89865 Then gError 76079

    'Se não encontrou a concorrencia ==> erro
    If lErro = 89865 Then gError 76080

    'Executa o relatório
    lErro = objRelatorio.ExecutarDireto("Geracao Pedido Compra Avulsa", "CONCORTO.NumIntDoc = @NCONCORR", 1, "CONCORR", "NCONCORR", objConcorrencia.lNumIntDoc)
    If lErro <> SUCESSO Then gError 76081

    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case 76079, 76081

        Case 76080
            Call Rotina_Erro(vbOKOnly, "ERRO_CONCORRENCIA_NAO_CADASTRADA", gErr, objConcorrencia.lCodigo)

        Case 76084
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CONCORRENCIA_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154568)

    End Select

    Exit Sub

End Sub

Private Sub BotaoPedCotacao_Click()

Dim objPedidoCotacao As New ClassPedidoCotacao

On Error GoTo Erro_BotaoPedCotacao_Click

    'Se nenhuma linha foi selecionada no Grid, sai da rotina
    If GridCotacoes.Row = 0 Then gError 89438

    objPedidoCotacao.lCodigo = StrParaLong(GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_PedidoCot_Col))
    objPedidoCotacao.iFilialEmpresa = giFilialEmpresa

    Call Chama_Tela("PedidoCotacaoCons", objPedidoCotacao)

    Exit Sub

Erro_BotaoPedCotacao_Click:

    Select Case gErr
    
        Case 89438
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154569)
            
    End Select
    
    Exit Sub

End Sub

Private Sub LabelConcorrencia_Click()

Dim objConcorrencia As New ClassConcorrencia
Dim colSelecao As New Collection

    'Coloca no objPedidoCotacao o código do pedido da tela
    objConcorrencia.lCodigo = StrParaLong(ConcorrenciaLabel.Caption)

    'Chama a tela de PedidoCotacaoTodosLista
    Call Chama_Tela("ConcorrenciaTodasLista", colSelecao, objConcorrencia, objEventoConcorrencia)

    Exit Sub

End Sub

Private Sub objEventoConcorrencia_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objConcorrencia As New ClassConcorrencia

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objConcorrencia = obj1

    'Chama Traz_PedidoCotacao_Tela
    lErro = Traz_Concorrencia_Tela(objConcorrencia)
    If lErro <> SUCESSO Then gError 89258

    'Fecha o sistema de setas
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 89258 'Erro tratado na rotina chamada.

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154570)

    End Select

    Exit Sub

End Sub

Public Sub BotaoEditarProduto_Click()

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer

On Error GoTo Erro_BotaoEditarProduto_Click

    'Se está editando um produto do GridProdutos1
    If FrameProdutos(1).Visible = True Then

        'Verifica se tem alguma linha selecionada no GridProdutos1
        If GridProdutos1.Row = 0 Then gError 67664

        'Verifica se o Produto está preenchido
        If Len(Trim(GridProdutos1.TextMatrix(GridProdutos1.Row, iGrid_Produto1_Col))) > 0 Then
            lErro = CF("Produto_Formata", GridProdutos1.TextMatrix(GridProdutos1.Row, iGrid_Produto1_Col), sProduto, iPreenchido)
            If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""
        End If

    'Se está editando um produto do GridProdutos2
    Else

        'Verifica se tem alguma linha selecionada no GridProdutos1
        If GridProdutos2.Row = 0 Then gError 67663

        'Verifica se o Produto está preenchido
        If Len(Trim(GridProdutos2.TextMatrix(GridProdutos2.Row, iGrid_Produto2_Col))) > 0 Then
            lErro = CF("Produto_Formata", GridProdutos2.TextMatrix(GridProdutos2.Row, iGrid_Produto2_Col), sProduto, iPreenchido)
            If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""
        End If

    End If

    objProduto.sCodigo = sProduto

    'Chama a Tela ProdutoCompraLista
    Call Chama_Tela("Produto", objProduto)

    Exit Sub

Erro_BotaoEditarProduto_Click:

    Select Case gErr

        Case 67663, 67664
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154571)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip1_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index

    End If

End Sub

Private Sub TabProdutos_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabProdutos.SelectedItem.Index <> iFrameProdutoAtual Then

        If TabStrip_PodeTrocarTab(iFrameProdutoAtual, TabProdutos, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        FrameProdutos(TabProdutos.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        FrameProdutos(iFrameProdutoAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameProdutoAtual = TabProdutos.SelectedItem.Index

    End If

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub


Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'se for o GridCotacoes
            Case GridCotacoes.Name

                lErro = Saida_Celula_GridCotacoes(objGridInt)
                If lErro <> SUCESSO Then gError 67719


        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 67720

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 67719, 67720
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154572)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridCotacoes(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridCotacoes

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Motivo Escolha
        Case iGrid_MotivoEscolhaCot_Col
            lErro = Saida_Celula_MotivoEscolhaCot(objGridInt)
            If lErro <> SUCESSO Then gError 67721

    End Select

    Saida_Celula_GridCotacoes = SUCESSO

    Exit Function

Erro_Saida_Celula_GridCotacoes:

    Saida_Celula_GridCotacoes = gErr

    Select Case gErr

        Case 67721

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154573)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MotivoEscolhaCot(objGridInt As AdmGrid) As Long
'Faz a saida de celula de MotivoEscolhaCot

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Saida_Celula_MotivoEscolhaCot

    Set objGridInt.objControle = MotivoEscolhaCot

    'Verifica se o MotivoEscolhaCot está preenchido
    If Len(Trim(MotivoEscolhaCot.Text)) > 0 Then

        lErro = Combo_Seleciona(MotivoEscolhaCot, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 67722

        'Se não encontrou o Motivo de Escolha ==> erro
        If lErro = 6730 Or lErro = 6731 Then gError 67724

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 67723

    Saida_Celula_MotivoEscolhaCot = SUCESSO

    Exit Function

Erro_Saida_Celula_MotivoEscolhaCot:

    Saida_Celula_MotivoEscolhaCot = gErr

    Select Case gErr

        Case 67722, 67723
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 67724
            Call Rotina_Erro(vbOKOnly, "ERRO_MOTIVO_NAO_ENCONTRADO", gErr, MotivoEscolhaCot.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154574)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Concorrência Consulta"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ConcorrenciaCons"

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

Private Sub Unload(objme As Object)

   RaiseEvent Unload

End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

'**** fim do trecho a ser copiado *****

'Tratamento dos Grids

'GridRequisicoes
Private Sub GridRequisicoes_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridRequisicoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRequisicoes, iAlterado)
    End If

End Sub

Private Sub GridRequisicoes_EnterCell()

    Call Grid_Entrada_Celula(objGridRequisicoes, iAlterado)

End Sub

Private Sub GridRequisicoes_GotFocus()

    Call Grid_Recebe_Foco(objGridRequisicoes)

End Sub

Private Sub GridRequisicoes_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridRequisicoes, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRequisicoes, iAlterado)
    End If

End Sub

Private Sub GridRequisicoes_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long

On Error GoTo Erro_GridRequisicoes_KeyDown

    Call Grid_Trata_Tecla1(KeyCode, objGridRequisicoes)

    Exit Sub

Erro_GridRequisicoes_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154575)

    End Select

    Exit Sub

End Sub

Private Sub GridRequisicoes_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridRequisicoes)

End Sub

Private Sub GridRequisicoes_RowColChange()

    Call Grid_RowColChange(objGridRequisicoes)

End Sub

Private Sub GridRequisicoes_Scroll()

    Call Grid_Scroll(objGridRequisicoes)

End Sub

'GridItensReq
Private Sub GridItensReq_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItensReq, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItensReq, iAlterado)
    End If

End Sub

Private Sub GridItensReq_EnterCell()

    Call Grid_Entrada_Celula(objGridItensReq, iAlterado)

End Sub

Private Sub GridItensReq_GotFocus()

    Call Grid_Recebe_Foco(objGridItensReq)

End Sub

Private Sub GridItensReq_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItensReq, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItensReq, iAlterado)
    End If

End Sub

Private Sub GridItensReq_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long

On Error GoTo Erro_GridItensReq_KeyDown

    Call Grid_Trata_Tecla1(KeyCode, objGridItensReq)

    Exit Sub

Erro_GridItensReq_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154576)

    End Select

    Exit Sub

End Sub

Private Sub GridItensReq_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridItensReq)

End Sub

Private Sub GridItensReq_RowColChange()

    Call Grid_RowColChange(objGridItensReq)

End Sub

Private Sub GridItensReq_Scroll()

    Call Grid_Scroll(objGridItensReq)

End Sub

'GridProdutos1
Private Sub GridProdutos1_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridProdutos1, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridProdutos1, iAlterado)
    End If

End Sub

Private Sub GridProdutos1_EnterCell()

    Call Grid_Entrada_Celula(objGridProdutos1, iAlterado)

End Sub

Private Sub GridProdutos1_GotFocus()

    Call Grid_Recebe_Foco(objGridProdutos1)

End Sub

Private Sub GridProdutos1_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridProdutos1, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridProdutos1, iAlterado)
    End If

End Sub

Private Sub GridProdutos1_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long

On Error GoTo Erro_GridProdutos1_KeyDown

    Call Grid_Trata_Tecla1(KeyCode, objGridProdutos1)

    Exit Sub

Erro_GridProdutos1_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154577)

    End Select

    Exit Sub

End Sub

Private Sub GridProdutos1_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridProdutos1)

End Sub

Private Sub GridProdutos1_RowColChange()

    Call Grid_RowColChange(objGridProdutos1)

End Sub

Private Sub GridProdutos1_Scroll()

    Call Grid_Scroll(objGridProdutos1)

End Sub

'GridProdutos2
Private Sub GridProdutos2_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridProdutos2, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridProdutos2, iAlterado)
    End If

End Sub

Private Sub GridProdutos2_EnterCell()

    Call Grid_Entrada_Celula(objGridProdutos2, iAlterado)

End Sub

Private Sub GridProdutos2_GotFocus()

    Call Grid_Recebe_Foco(objGridProdutos2)

End Sub

Private Sub GridProdutos2_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridProdutos2, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridProdutos2, iAlterado)
    End If

End Sub

Private Sub GridProdutos2_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long

On Error GoTo Erro_GridProdutos2_KeyDown

    Call Grid_Trata_Tecla1(KeyCode, objGridProdutos2)

    Exit Sub

Erro_GridProdutos2_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154578)

    End Select

    Exit Sub

End Sub

Private Sub GridProdutos2_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridProdutos2)

End Sub

Private Sub GridProdutos2_RowColChange()

    Call Grid_RowColChange(objGridProdutos2)

End Sub

Private Sub GridProdutos2_Scroll()

    Call Grid_Scroll(objGridProdutos2)

End Sub

Private Sub GridCotacoes_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridCotacoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCotacoes, iAlterado)
    End If

End Sub

Private Sub GridCotacoes_EnterCell()

    Call Grid_Entrada_Celula(objGridCotacoes, iAlterado)

End Sub

Private Sub GridCotacoes_GotFocus()

    Call Grid_Recebe_Foco(objGridCotacoes)

End Sub

Private Sub GridCotacoes_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCotacoes, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCotacoes, iAlterado)
    End If

End Sub

Private Sub GridCotacoes_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long

On Error GoTo Erro_GridCotacoes_KeyDown

    Call Grid_Trata_Tecla1(KeyCode, objGridCotacoes)

    Exit Sub

Erro_GridCotacoes_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154579)

    End Select

    Exit Sub

End Sub

Private Sub GridCotacoes_LeaveCell()

    Call Saida_Celula(objGridCotacoes)

End Sub

Private Sub GridCotacoes_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridCotacoes)

End Sub

Private Sub GridCotacoes_RowColChange()

    Call Grid_RowColChange(objGridCotacoes)

End Sub

Private Sub GridCotacoes_Scroll()

    Call Grid_Scroll(objGridCotacoes)

End Sub

Private Sub MotivoEscolhaCot_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MotivoEscolhaCot_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCotacoes)

End Sub

Private Sub MotivoEscolhaCot_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCotacoes)

End Sub

Private Sub MotivoEscolhaCot_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCotacoes.objControle = MotivoEscolhaCot
    lErro = Grid_Campo_Libera_Foco(objGridCotacoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

'GridPedidos
Private Sub GridPedidos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridPedidos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPedidos, iAlterado)
    End If

End Sub

Private Sub GridPedidos_EnterCell()

    Call Grid_Entrada_Celula(objGridPedidos, iAlterado)

End Sub

Private Sub GridPedidos_GotFocus()

    Call Grid_Recebe_Foco(objGridPedidos)

End Sub

Private Sub GridPedidos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridPedidos, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPedidos, iAlterado)
    End If

End Sub

Private Sub GridPedidos_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long

On Error GoTo Erro_GridPedidos_KeyDown

    Call Grid_Trata_Tecla1(KeyCode, objGridPedidos)

    Exit Sub

Erro_GridPedidos_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154580)

    End Select

    Exit Sub

End Sub

Private Sub GridPedidos_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridPedidos)

End Sub

Private Sub GridPedidos_RowColChange()

    Call Grid_RowColChange(objGridPedidos)

End Sub

Private Sub GridPedidos_Scroll()

    Call Grid_Scroll(objGridPedidos)

End Sub


Private Sub TaxaEmpresa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TaxaEmpresa, Source, X, Y)
End Sub

Private Sub TaxaEmpresa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TaxaEmpresa, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub ConcorrenciaLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ConcorrenciaLabel, Source, X, Y)
End Sub

Private Sub ConcorrenciaLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ConcorrenciaLabel, Button, Shift, X, Y)
End Sub


Function Carrega_TipoTributacao() As Long
'Carrega Tipos de Tributação

Dim lErro As Long
Dim colTributacao As New AdmColCodigoNome
Dim iIndice As Integer

On Error GoTo Erro_Carrega_TipoTributacao

    'Lê os Tipos de Tributação associadas a Compras
    lErro = CF("TiposTributacaoCompras_Le", colTributacao)
    If lErro <> SUCESSO Then gError 66123

    'Carrega Tipos de Tributação
    For iIndice = 1 To colTributacao.Count
        TipoTributacaoCot.AddItem colTributacao(iIndice).iCodigo & SEPARADOR & colTributacao(iIndice).sNome
        TipoTributacaoCot.ItemData(TipoTributacaoCot.NewIndex) = colTributacao(iIndice).iCodigo
    Next

    Carrega_TipoTributacao = SUCESSO

    Exit Function

Erro_Carrega_TipoTributacao:

    Carrega_TipoTributacao = gErr

    Select Case gErr

        Case 66123

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154581)

    End Select

    Exit Function

End Function

Function Traz_Concorrencia_Tela(objConcorrencia As ClassConcorrencia) As Long
'Traz os dados da Concorrência para a tela

Dim lErro As Long
Dim objItemConcorrencia As ClassItemConcorrencia
Dim colPedidos As New Collection
Dim colItensCotacao As New Collection

On Error GoTo Erro_Traz_Concorrencia_Tela

    'Dados da Concorrência
    ConcorrenciaLabel.Caption = objConcorrencia.lCodigo
    DescricaoLabel.Caption = objConcorrencia.sDescricao
    TaxaEmpresa.Caption = Format(objConcorrencia.dTaxaFinanceira, "Percent")

    'Lê os Itens da Concorrência em questão
    lErro = CF("ItensConcorrenciaTodos_Le", objConcorrencia)
    If lErro <> SUCESSO Then gError 67665

    Set gcolRequisicaoCompra = New Collection

    'Lê as Requisições vinculadas a concorrência
    lErro = CF("RequisicoesTodas_Le_Concorrencia", objConcorrencia, gcolRequisicaoCompra)
    If lErro <> SUCESSO Then gError 67666

    'Lê Pedidos de Compras vinculados a Concorrência
    lErro = CF("PedidoCompraTodos_Le_Concorrencia", objConcorrencia, colPedidos)
    If lErro <> SUCESSO Then gError 67672

    'Preenche o Grid de Requisições
    lErro = GridRequisicoes_Preenche()
    If lErro <> SUCESSO Then gError 67667

    Call Seleciona_ItensRelacionados(objConcorrencia)

    'Traz os itens de Requisição para a tela
    lErro = GridItensReq_Preenche()
    If lErro <> SUCESSO Then gError 67671
    
    If gcolRequisicaoCompra.Count = 0 Then
        Call Inicializa_QuantSupl(objConcorrencia)
    End If

    'Atualiza Grid de Cotações
    lErro = Grids_Produto_Preenche(objConcorrencia)
    If lErro <> SUCESSO Then gError 67670

    'Preenche Grid de Pedidos
    lErro = Preenche_GridPedidos(colPedidos)
    If lErro <> SUCESSO Then gError 67673

    Traz_Concorrencia_Tela = SUCESSO

    Exit Function

Erro_Traz_Concorrencia_Tela:

    Traz_Concorrencia_Tela = gErr

    Select Case gErr

        Case 67665, 67666, 67667, 67668, 67669, 67670, 67671, 67672, 67673

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154582)

    End Select

    Exit Function

End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function GridRequisicoes_Preenche() As Long

Dim lErro As Long
Dim objRequisicao As New ClassRequisicaoCompras
Dim iLinha As Long
Dim objFilialEmpresa As New AdmFiliais
Dim objRequisitante As New ClassRequisitante
Dim sCclMascarado As String
Dim colRequisitantes As New AdmCollCodigoNome
Dim colFiliais As New AdmCollCodigoNome
Dim objlCodigoNome As AdmlCodigoNome
Dim iPosicao As Integer

On Error GoTo Erro_GridRequisicoes_Preenche

    'Limpa o Grid de Requisições
    Call Grid_Limpa(objGridRequisicoes)

    If gcolRequisicaoCompra.Count > 0 Then

        'Preenche o GridRequisicoes
        For Each objRequisicao In gcolRequisicaoCompra
            objRequisicao.iSelecionado = MARCADO
            iLinha = objGridRequisicoes.iLinhasExistentes + 1
    
            Call Busca_Na_Colecao(colFiliais, objRequisicao.iFilialEmpresa, iPosicao)
    
            If iPosicao = 0 Then
    
                objFilialEmpresa.iCodFilial = objRequisicao.iFilialEmpresa
    
                'Lê a FilialEmpresa
                lErro = CF("FilialEmpresa_Le", objFilialEmpresa, True)
                If lErro <> SUCESSO And lErro <> 27378 Then gError 63976
    
                'Se não encontrou ==>Erro
                If lErro = 27378 Then gError 63977
    
                Set objlCodigoNome = New AdmlCodigoNome
    
                objlCodigoNome.lCodigo = objFilialEmpresa.iCodFilial
                objlCodigoNome.sNome = objFilialEmpresa.sNome
    
                colFiliais.Add objlCodigoNome.lCodigo, objlCodigoNome.sNome
    
            Else
    
                Set objlCodigoNome = colFiliais(iPosicao)
    
            End If
    
            'Preenche a Filial de Requisicao com código e nome reduzido
            GridRequisicoes.TextMatrix(iLinha, iGrid_FilialEmpresaReq_Col) = objlCodigoNome.lCodigo & SEPARADOR & objlCodigoNome.sNome
            GridRequisicoes.TextMatrix(iLinha, iGrid_CodigoReq_Col) = objRequisicao.lCodigo
    
            'Verifica se DataLimite é diferente de Data Nula
            If objRequisicao.dtDataLimite <> DATA_NULA Then GridRequisicoes.TextMatrix(iLinha, iGrid_DataLimiteReq_Col) = Format(objRequisicao.dtDataLimite, "dd/mm/yyyy")
    
            'Verifica se Data é diferente de Data Nula
            If objRequisicao.dtData <> DATA_NULA Then GridRequisicoes.TextMatrix(iLinha, iGrid_DataReq_Col) = Format(objRequisicao.dtData, "dd/mm/yyyy")
    
            GridRequisicoes.TextMatrix(iLinha, iGrid_UrgenteReq_Col) = objRequisicao.lUrgente
    
            Call Busca_Na_Colecao(colRequisitantes, objRequisicao.lRequisitante, iPosicao)
            
            If iPosicao = 0 Then
                objRequisitante.lCodigo = objRequisicao.lRequisitante
        
                'Lê o requisitante
                lErro = CF("Requisitante_Le", objRequisitante)
                If lErro <> SUCESSO And lErro <> 49084 Then gError 63978
        
                'Se não encontrou o Requisitante ==> Erro
                If lErro = 49084 Then gError 63979
                
                Set objlCodigoNome = New AdmlCodigoNome
                
                objlCodigoNome.lCodigo = objRequisitante.lCodigo
                objlCodigoNome.sNome = objRequisitante.sNomeReduzido
                
                colRequisitantes.Add objlCodigoNome.lCodigo, objlCodigoNome.sNome
                
            Else
                Set objlCodigoNome = colRequisitantes(iPosicao)
            End If
            
            'Preenche o Requisitante com o código e o nome reduzido
            GridRequisicoes.TextMatrix(iLinha, iGrid_RequisitanteReq_Col) = objlCodigoNome.lCodigo & SEPARADOR & objlCodigoNome.sNome
    
            'Se o Ccl está preenchida
            If Len(Trim(objRequisicao.sCcl)) > 0 Then
    
                'Mascara o Produto
                lErro = Mascara_MascararCcl(objRequisicao.sCcl, sCclMascarado)
                If lErro <> SUCESSO Then gError 63980
    
                'Preenche o Ccl
                GridRequisicoes.TextMatrix(iLinha, iGrid_CclReq_Col) = sCclMascarado
    
            End If
    
            'Preenche a Observacao
            GridRequisicoes.TextMatrix(iLinha, iGrid_ObservacaoReq_Col) = objRequisicao.sObservacao
                           
            objGridRequisicoes.iLinhasExistentes = iLinha
        
        Next
    
        Call Grid_Refresh_Checkbox(objGridRequisicoes)

    End If
    
    GridRequisicoes_Preenche = SUCESSO
    
    Exit Function
    
Erro_GridRequisicoes_Preenche:

    GridRequisicoes_Preenche = gErr
    
    Select Case gErr
    
        Case 63976, 63978, 63980
            'Erros tratados nas rotinas chamadas

        Case 63977
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case 63979
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_NAO_CADASTRADO", gErr, objRequisitante.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154583)

    End Select
        
    Exit Function
        
End Function

Function GridItensReq_Preenche() As Long
'Preenche o GridItensReq com os Itens da Requisicao passada como parametro

Dim lErro As Long
Dim objRequisicao As ClassRequisicaoCompras
Dim objItemReqCompras As New ClassItemReqCompras
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objFornecedor As New ClassFornecedor
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim iLinha As Integer
Dim sCclMascarado As String
Dim objFilialEmpresa As New AdmFiliais
Dim sProdutoMascarado As String
Dim objlCodigoNome As AdmlCodigoNome
Dim colFiliais As New AdmCollCodigoNome
Dim colAlmoxarifados As New AdmCollCodigoNome
Dim colFornecedor As New AdmCollCodigoNome
Dim colFilialForn As New Collection
Dim iPosicao As Integer

On Error GoTo Erro_GridItensReq_Preenche

    'Limpa o grid de itens
    Call Grid_Limpa(objGridItensReq)
    
    'Para cada requisicao
    For Each objRequisicao In gcolRequisicaoCompra
        'Se a req está selecionada
        If objRequisicao.iSelecionado = MARCADO Then
            'Para cada item
            For Each objItemReqCompras In objRequisicao.colItens
        
                iLinha = iLinha + 1
                'BUsca a filial da req na colfiliais
                Call Busca_Na_Colecao(colFiliais, objRequisicao.iFilialEmpresa, iPosicao)
            
                If iPosicao = 0 Then
               
                    objFilialEmpresa.iCodFilial = objRequisicao.iFilialEmpresa
                    'Lê a FilialEmpresa
                    lErro = CF("FilialEmpresa_Le", objFilialEmpresa, True)
                    If lErro <> SUCESSO And lErro <> 27378 Then gError 68059
        
                    'Se não encontrou a filial ==>erro
                    If lErro = 27378 Then gError 68060
        
                    Set objlCodigoNome = New AdmlCodigoNome
                    
                    objlCodigoNome.lCodigo = objFilialEmpresa.iCodFilial
                    objlCodigoNome.sNome = objFilialEmpresa.sNome
                    
                    colFiliais.Add objlCodigoNome.lCodigo, objlCodigoNome.sNome
        
                Else
                    Set objlCodigoNome = colFiliais(iPosicao)
                End If
        
                GridItensReq.TextMatrix(iLinha, iGrid_EscolhidoItemReq_Col) = objItemReqCompras.iSelecionado
                GridItensReq.TextMatrix(iLinha, iGrid_FilialItemReq_Col) = objlCodigoNome.lCodigo & SEPARADOR & objlCodigoNome.sNome
                GridItensReq.TextMatrix(iLinha, iGrid_CodigoReqItemReq_Col) = objRequisicao.lCodigo
        
                GridItensReq.TextMatrix(iLinha, iGrid_ItemItemReq_Col) = objItemReqCompras.iItem
        
                'Mascara o Produto
                lErro = Mascara_RetornaProdutoEnxuto(objItemReqCompras.sProduto, sProdutoMascarado)
                If lErro <> SUCESSO Then gError 68064
        
                ProdutoItemReq.PromptInclude = False
                ProdutoItemReq.Text = sProdutoMascarado
                ProdutoItemReq.PromptInclude = True
                GridItensReq.TextMatrix(iLinha, iGrid_ProdutoItemReq_Col) = ProdutoItemReq.Text
                GridItensReq.TextMatrix(iLinha, iGrid_DescProdutoItemReq_Col) = objItemReqCompras.sDescProduto
                
                GridItensReq.TextMatrix(iLinha, iGrid_UMItemReq_Col) = objItemReqCompras.sUM
                GridItensReq.TextMatrix(iLinha, iGrid_QuantidadeItemReq_Col) = Formata_Estoque(objItemReqCompras.dQuantidade)
                GridItensReq.TextMatrix(iLinha, iGrid_QuantPedidaItemReq_Col) = Formata_Estoque(objItemReqCompras.dQuantPedida)
                GridItensReq.TextMatrix(iLinha, iGrid_QuantRecebidaItemReq_Col) = Formata_Estoque(objItemReqCompras.dQuantRecebida)
        
                GridItensReq.TextMatrix(iLinha, iGrid_QuantComprarItemReq_Col) = Formata_Estoque(objItemReqCompras.dQuantNoPedido)
        
                If objItemReqCompras.iAlmoxarifado <> 0 Then
                    
                    Call Busca_Na_Colecao(colAlmoxarifados, objItemReqCompras.iAlmoxarifado, iPosicao)
                
                    If iPosicao = 0 Then
                
                        objAlmoxarifado.iCodigo = objItemReqCompras.iAlmoxarifado
            
                        'Lê o almoxarifado
                        lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                        If lErro <> SUCESSO And lErro <> 25056 Then gError 63984
            
                        'Se não encontrou ==> Erro
                        If lErro = 25056 Then gError 63985
        
                        Set objlCodigoNome = New AdmlCodigoNome
                        
                        objlCodigoNome.lCodigo = objAlmoxarifado.iCodigo
                        objlCodigoNome.sNome = objAlmoxarifado.sNomeReduzido
                        
                        colAlmoxarifados.Add objlCodigoNome.lCodigo, objlCodigoNome.sNome
        
                    Else
                        Set objlCodigoNome = colAlmoxarifados(iPosicao)
                    End If
                
                    GridItensReq.TextMatrix(iLinha, iGrid_AlmoxarifadoItemReq_Col) = objlCodigoNome.sNome
                
                End If
        
        
                If Len(Trim(objItemReqCompras.sCcl)) > 0 Then
        
                    'Mascara o Ccl
                    lErro = Mascara_MascararCcl(objItemReqCompras.sCcl, sCclMascarado)
                    If lErro <> SUCESSO Then gError 63986
                    
                    GridItensReq.TextMatrix(iLinha, iGrid_CclItemReq_Col) = sCclMascarado
                End If
        
        
                If objItemReqCompras.lFornecedor <> 0 And objItemReqCompras.iFilial <> 0 Then
                    
                    Call Busca_Na_Colecao(colFornecedor, objItemReqCompras.lFornecedor, iPosicao)
        
                    If iPosicao = 0 Then
        
                        objFornecedor.lCodigo = objItemReqCompras.lFornecedor
            
                        'Lê o Fornecedor
                        lErro = CF("Fornecedor_Le", objFornecedor)
                        If lErro <> SUCESSO And lErro <> 12729 Then gError 63987
            
                        'Se não encontrou o Fornecedor==> Erro
                        If lErro = 12729 Then gError 63988
                        
                        Set objlCodigoNome = New AdmlCodigoNome
                    
                        objlCodigoNome.lCodigo = objFornecedor.lCodigo
                        objlCodigoNome.sNome = objFornecedor.sNomeReduzido
                        
                        colFornecedor.Add objlCodigoNome.lCodigo, objlCodigoNome.sNome
                    
                    Else
                        Set objlCodigoNome = colFornecedor(iPosicao)
                    End If
        
                    GridItensReq.TextMatrix(iLinha, iGrid_FornecedorItemReq_Col) = objlCodigoNome.sNome
        
                    Call Busca_FilialForn(colFilialForn, objItemReqCompras.lFornecedor, objItemReqCompras.iFilial, iPosicao)
                    
                    If iPosicao = 0 Then
                        Set objFilialFornecedor = New ClassFilialFornecedor
                        objFilialFornecedor.iCodFilial = objItemReqCompras.iFilial
                        objFilialFornecedor.lCodFornecedor = objItemReqCompras.lFornecedor
                        
                        'Lê a FilialFornecedor
                        lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
                        If lErro <> SUCESSO And lErro <> 12929 Then gError 63989
            
                        'Se não encontrou==>Erro
                        If lErro = 12929 Then gError 63990
                    Else
                        Set objFilialFornecedor = colFilialForn(iPosicao)
                    End If
        
                    GridItensReq.TextMatrix(iLinha, iGrid_FilialFornItemReq_Col) = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome
                
                    If objItemReqCompras.iExclusivo = MARCADO Then
                        GridItensReq.TextMatrix(iLinha, iGrid_ExclusivoItemReq_Col) = "Exclusivo"
                    Else
                        GridItensReq.TextMatrix(iLinha, iGrid_ExclusivoItemReq_Col) = "Preferencial"
                    End If
                    
                End If
        
                GridItensReq.TextMatrix(iLinha, iGrid_ObservacaoItemReq_Col) = objItemReqCompras.sObservacao
        
            Next
        End If
    Next
    
    'Atualiza o número de linhas existentes do GridItensReq
    objGridItensReq.iLinhasExistentes = iLinha
    
    Call Grid_Refresh_Checkbox(objGridItensReq)
    
    GridItensReq_Preenche = SUCESSO

    Exit Function

Erro_GridItensReq_Preenche:

    GridItensReq_Preenche = gErr

    Select Case gErr

        Case 63982, 63984, 63986, 63987, 63989, 68059, 68064
            'Erros tratados nas rotinas chamadas

        Case 63983
            Call Rotina_Erro(vbOKOnly, "ERRO_ITENSREQCOMPRA_NAO_CADASTRADO", gErr, objRequisicao.lNumIntDoc, objItemReqCompras.lReqCompra)

        Case 63985
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)

        Case 63988
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)

        Case 63990
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objFilialFornecedor.iCodFilial, objFornecedor.lCodigo)

        Case 68060
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154584)

    End Select

    Exit Function
    
End Function

Function Grids_Produto_Preenche(objConcorrencia As ClassConcorrencia) As Long

Dim iLinha1 As Integer, iLinha2 As Integer
Dim objItemConc As New ClassItemConcorrencia
Dim sProdutoEnxuto As String
Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objQuantSupl As ClassQuantSuplementar
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim colItensSaida As New Collection
Dim colCampos As New Collection
Dim colFilForn As New Collection
Dim colFornec As New AdmCollCodigoNome
Dim objCodNome As New AdmlCodigoNome
Dim iPosicao As Integer

On Error GoTo Erro_Grids_Produto_Preenche
    
    'Limpa o grid de produtos1
    Call Grid_Limpa(objGridProdutos1)
    
        
    colCampos.Add "sProduto"
    colCampos.Add "lFornecedor"
    colCampos.Add "iFilial"
    
    'Ordena os itens de concorrência por produto
    lErro = Ordena_Colecao(objConcorrencia.colItens, colItensSaida, colCampos)
    If lErro <> SUCESSO Then gError 63808

    Set objConcorrencia.colItens = colItensSaida
    
    iLinha1 = 0
    iLinha2 = 0
    
    'Para cada item de concorrência
    For Each objItemConc In objConcorrencia.colItens
        
        iLinha1 = iLinha1 + 1
        
        lErro = Mascara_RetornaProdutoEnxuto(objItemConc.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 62778
        
        Produto1.PromptInclude = False
        Produto1.Text = sProdutoEnxuto
        Produto1.PromptInclude = True
        
        'Preenche o produto
        GridProdutos1.TextMatrix(iLinha1, iGrid_Produto1_Col) = Produto1.Text
        GridProdutos1.TextMatrix(iLinha1, iGrid_DescProduto1_Col) = objItemConc.sDescricao
        GridProdutos1.TextMatrix(iLinha1, iGrid_UnidadeMed1_Col) = objItemConc.sUM
        GridProdutos1.TextMatrix(iLinha1, iGrid_QuantComprar1_Col) = Formata_Estoque(objItemConc.dQuantidade)
        GridProdutos1.TextMatrix(iLinha1, iGrid_Urgente1_Col) = Formata_Estoque(objItemConc.dQuantUrgente)
        
        'Se o Fornecedor está preenchido
        If objItemConc.lFornecedor > 0 And objItemConc.iFilial > 0 Then
            
            'verifica se esse forn já foi lido
            Call Busca_Na_Colecao(colFornec, objItemConc.lFornecedor, iPosicao)
        
            If iPosicao = 0 Then
                objFornecedor.lCodigo = objItemConc.lFornecedor
                
                lErro = CF("Fornecedor_Le", objFornecedor)
                If lErro <> SUCESSO And lErro <> 12729 Then gError 62779
                If lErro <> SUCESSO Then gError 62780
                            
                Set objCodNome = New AdmlCodigoNome
                
                objCodNome.lCodigo = objFornecedor.lCodigo
                objCodNome.sNome = objFornecedor.sNomeReduzido
                
                colFornec.Add objCodNome.lCodigo, objCodNome.sNome
            Else
                Set objCodNome = colFornec(iPosicao)
            End If
            
            'Preenche o fornecedor
            GridProdutos1.TextMatrix(iLinha1, iGrid_Fornecedor1_Col) = objCodNome.sNome
            
            'Verifica se essa filial já foi lida
            Call Busca_FilialForn(colFilForn, objItemConc.lFornecedor, objItemConc.iFilial, iPosicao)
            
            If iPosicao = 0 Then
                Set objFilialFornecedor = New ClassFilialFornecedor
                objFilialFornecedor.lCodFornecedor = objItemConc.lFornecedor
                objFilialFornecedor.iCodFilial = objItemConc.iFilial
                
                lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
                If lErro <> SUCESSO And lErro <> 12929 Then gError 63989
                
                'Se não encontrou==>Erro
                If lErro = 12929 Then gError 63990
                
                colFilForn.Add objFilialFornecedor
            Else
                Set objFilialFornecedor = colFilForn(iPosicao)
            End If
            'Preenche a filial
            GridProdutos1.TextMatrix(iLinha1, iGrid_FilialForn1_Col) = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome
        
        End If
    Next
    
    objGridProdutos1.iLinhasExistentes = iLinha1
    
    Call Grid_Refresh_Checkbox(objGridProdutos1)
    
    'Preenche o grid de produtos 2
    lErro = GridProdutos2_Preenche(objConcorrencia)
    If lErro <> SUCESSO Then gError 62781
    
    Call GridCotacoes_Preenche(objConcorrencia)
    
    Grids_Produto_Preenche = SUCESSO
    
    Exit Function
    
Erro_Grids_Produto_Preenche:

    Grids_Produto_Preenche = gErr
    
    Select Case gErr
        
        Case 63808, 62778, 62779, 63989, 62781
        
        Case 62780
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)
            
        Case 63990
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objFilialFornecedor.iCodFilial, objFilialFornecedor.lCodFornecedor)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154585)
            
    End Select
    
    Exit Function

End Function
Private Function GridProdutos2_Preenche(objConcorrencia As ClassConcorrencia) As Long
'Preenche o grid de produtos 2

Dim objItemConc As ClassItemConcorrencia
Dim objQuantSupl As ClassQuantSuplementar
Dim iLinha2 As Integer, lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim iLinha1 As Integer
Dim colFilEmp As New AdmCollCodigoNome
Dim colFilForn As New Collection
Dim colForn As New AdmCollCodigoNome
Dim objCodNome As AdmlCodigoNome
Dim objFilEmp As New AdmFiliais
Dim iPosicao As Integer

On Error GoTo Erro_GridProdutos2_Preenche
    
    'Limpa o grid de produtos2
    Call Grid_Limpa(objGridProdutos2)
       
    iLinha1 = 0
    iLinha2 = 0
    
    'Para cada item de conc
    For Each objItemConc In objConcorrencia.colItens
        iLinha1 = iLinha1 + 1
           
        'Para cada quant supl
        For Each objQuantSupl In objItemConc.colQuantSuplementar
        
            iLinha2 = iLinha2 + 1
            'Preenche com os dados do item de conorrência
            GridProdutos2.TextMatrix(iLinha2, iGrid_Produto2_Col) = GridProdutos1.TextMatrix(iLinha1, iGrid_Produto1_Col)
            GridProdutos2.TextMatrix(iLinha2, iGrid_DescProduto2_Col) = objItemConc.sDescricao
            GridProdutos2.TextMatrix(iLinha2, iGrid_UnidadeMed2_Col) = objItemConc.sUM
            GridProdutos2.TextMatrix(iLinha2, iGrid_QuantComprar2_Col) = Formata_Estoque(objQuantSupl.dQuantidade)
              
            If objQuantSupl.iTipoDestino = TIPO_DESTINO_EMPRESA Then
                
                Call Busca_Na_Colecao(colFilEmp, objQuantSupl.iFilialDestino, iPosicao)
                
                If iPosicao = 0 Then
                
                    objFilEmp.lCodEmpresa = glEmpresa
                    objFilEmp.iCodFilial = objQuantSupl.iFilialDestino
                                                            
                    lErro = CF("FilialEmpresa_Le", objFilEmp, True)
                    If lErro <> SUCESSO And lErro <> 27378 Then gError 62788
                    If lErro <> SUCESSO Then gError 62789
                    
                    Set objCodNome = New AdmlCodigoNome
                    
                    objCodNome.sNome = objFilEmp.sNome
                    objCodNome.lCodigo = objFilEmp.iCodFilial
                    
                    colFilEmp.Add objCodNome.lCodigo, objCodNome.sNome
                
                Else
                    Set objCodNome = colFilEmp(iPosicao)
                End If
                'Preenche os dados do destino
                GridProdutos2.TextMatrix(iLinha2, iGrid_TipoDestino_Col) = "Empresa"
                GridProdutos2.TextMatrix(iLinha2, iGrid_Destino_Col) = ""
              
                GridProdutos2.TextMatrix(iLinha2, iGrid_FilialDestino_Col) = objCodNome.lCodigo & SEPARADOR & objCodNome.sNome
              
            ElseIf objQuantSupl.iTipoDestino = TIPO_DESTINO_FORNECEDOR Then
                
                GridProdutos2.TextMatrix(iLinha2, iGrid_TipoDestino_Col) = "Fornecedor"
                                      
                Call Busca_Na_Colecao(colForn, objQuantSupl.lFornCliDestino, iPosicao)
                                    
                If iPosicao = 0 Then
                    objFornecedor.lCodigo = objQuantSupl.lFornCliDestino
                    
                    'Lê o fornecedor
                    lErro = CF("Fornecedor_Le", objFornecedor)
                    If lErro <> SUCESSO And lErro <> 12729 Then gError 62790
                    If lErro <> SUCESSO Then gError 62791
                                        
                    Set objCodNome = New AdmlCodigoNome
                    
                    objCodNome.lCodigo = objFornecedor.lCodigo
                    objCodNome.sNome = objFornecedor.sNomeReduzido
                
                    colForn.Add objCodNome.lCodigo, objCodNome.sNome
                Else
                    Set objCodNome = colForn(iPosicao)
                End If
                
                GridProdutos2.TextMatrix(iLinha2, iGrid_Destino_Col) = objCodNome.sNome
                  
                Call Busca_FilialForn(colFilForn, objQuantSupl.lFornCliDestino, objQuantSupl.iFilialDestino, iPosicao)
                
                If iPosicao = 0 Then
                    Set objFilialFornecedor = New ClassFilialFornecedor
                    
                    objFilialFornecedor.lCodFornecedor = objQuantSupl.lFornCliDestino
                    objFilialFornecedor.iCodFilial = objQuantSupl.iFilialDestino
                
                    lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
                    If lErro <> SUCESSO And lErro <> 12929 Then gError 63989
                
                    'Se não encontrou==>Erro
                    If lErro = 12929 Then gError 63990
                                     
                    colFilForn.Add objFilialFornecedor
                Else
                    Set objFilialFornecedor = colFilForn(iPosicao)
                End If
                'Preenche os dados do destino
                GridProdutos2.TextMatrix(iLinha2, iGrid_FilialDestino_Col) = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome
              
            End If
            
            GridProdutos2.TextMatrix(iLinha2, iGrid_Fornecedor2_Col) = GridProdutos1.TextMatrix(iLinha1, iGrid_Fornecedor1_Col)
            GridProdutos2.TextMatrix(iLinha2, iGrid_FilialForn2_Col) = GridProdutos1.TextMatrix(iLinha1, iGrid_FilialForn1_Col)
                
        Next
    Next
    
    objGridProdutos2.iLinhasExistentes = iLinha2

    Call Grid_Refresh_Checkbox(objGridProdutos2)
    
    GridProdutos2_Preenche = SUCESSO
    
    Exit Function
    
Erro_GridProdutos2_Preenche:

    GridProdutos2_Preenche = gErr
    
    Select Case gErr
    
        Case 62788, 62790, 63989

        Case 62789
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilEmp.iCodFilial)
        
        Case 62791
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)

        Case 63990
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objFilialFornecedor.iCodFilial, objFornecedor.lCodigo)

    End Select
    
    Exit Function
    
End Function
Function Move_Tela_Memoria(objConcorrencia As ClassConcorrencia) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    'Guarda os dados da Concorrência
    objConcorrencia.lCodigo = StrParaLong(ConcorrenciaLabel.Caption)
    objConcorrencia.dTaxaFinanceira = PercentParaDbl(TaxaEmpresa.Caption)
    objConcorrencia.iFilialEmpresa = giFilialEmpresa

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154586)

    End Select

    Exit Function

End Function

Sub Calcula_Preferencia(objCotItemConc As ClassCotacaoItemConc, sProduto As String, dQuantComprar As Double)
'Calcula a Preferência

Dim iIndice As Integer
Dim dQuantPreferencial As Double
Dim dQuantComprarItem As Double
    
    dQuantPreferencial = 0
    
    If dQuantComprar = 0 Then Exit Sub
    
    For iIndice = 1 To objGridItensReq.iLinhasExistentes
    
        If StrParaInt(GridItensReq.TextMatrix(iIndice, iGrid_EscolhidoItemReq_Col)) = MARCADO Then
        
            If GridItensReq.TextMatrix(iIndice, iGrid_ProdutoItemReq_Col) = sProduto And _
              GridItensReq.TextMatrix(iIndice, iGrid_FilialFornItemReq_Col) = objCotItemConc.sFilial And _
              GridItensReq.TextMatrix(iIndice, iGrid_FornecedorItemReq_Col) = objCotItemConc.sFornecedor And _
              GridItensReq.TextMatrix(iIndice, iGrid_ExclusivoItemReq_Col) = "Preferencial" Then
                
                Call Busca_QuantComprar_ItemReq(StrParaLong(GridItensReq.TextMatrix(iIndice, iGrid_CodigoReqItemReq_Col)), Codigo_Extrai(GridItensReq.TextMatrix(iIndice, iGrid_FilialItemReq_Col)), StrParaInt(GridItensReq.TextMatrix(iIndice, iGrid_ItemItemReq_Col)), dQuantComprarItem)
              
                dQuantPreferencial = dQuantPreferencial + dQuantComprarItem
            End If
        End If
    Next
            
    objCotItemConc.dPreferencia = dQuantPreferencial / dQuantComprar
    Exit Sub

End Sub

Function Preenche_GridPedidos(colPedidos As Collection) As Long

Dim lErro As Long
Dim objItemPCInfo As ClassItemPedCompraInfo
Dim iLinha As Integer
Dim sProdutoMascarado As String
Dim iIndice As Integer

On Error GoTo Erro_Preenche_GridPedidos

    Call Grid_Limpa(objGridPedidos)

    'Para cada ItemPC da coleção
    For Each objItemPCInfo In colPedidos

        iLinha = iLinha + 1
        
        lErro = Mascara_MascararProduto(objItemPCInfo.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 86158

        'Preenche uma linha do GridPedidos
        GridPedidos.TextMatrix(iLinha, iGrid_CodPedido_Col) = objItemPCInfo.lPedCompra
        GridPedidos.TextMatrix(iLinha, iGrid_ProdutoPC_Col) = sProdutoMascarado
        GridPedidos.TextMatrix(iLinha, iGrid_QuantPedido_Col) = Formata_Estoque(objItemPCInfo.dQuantReceber)
        GridPedidos.TextMatrix(iLinha, iGrid_UMPedido_Col) = objItemPCInfo.sUM
        
        For iIndice = 0 To MoedaPC.ListCount - 1
            If objItemPCInfo.iMoeda = MoedaPC.ItemData(iIndice) Then
                GridPedidos.TextMatrix(iLinha, iGrid_MoedaPC_Col) = MoedaPC.List(iIndice)
                Exit For
            End If
        Next

    Next

    objGridPedidos.iLinhasExistentes = iLinha

    Preenche_GridPedidos = SUCESSO

    Exit Function

Erro_Preenche_GridPedidos:

    Preenche_GridPedidos = gErr

    Select Case gErr
    
        Case 86158

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154587)

    End Select

    Exit Function

End Function

Private Function Seleciona_ItensRelacionados(objConcorrencia As ClassConcorrencia)

Dim objReqCompra As ClassRequisicaoCompras
Dim objItemRC As ClassItemReqCompras
Dim objItemRCItemConc As ClassItemRCItemConcorrencia
Dim objItemConc As ClassItemConcorrencia
Dim objQtSupl As ClassQuantSuplementar
Dim bAchou As Boolean, bAChouQtSup As Boolean
Dim iTipoTribItem As Integer, dQuantMaior As Double
Dim objCotItemConc As ClassCotacaoItemConc
Dim objProduto As New ClassProduto
Dim lErro As Long
Dim dFator As Double

On Error GoTo Erro_Seleciona_ItensRelacionados

    'Para cada Item da concorrência
    For Each objItemConc In objConcorrencia.colItens
    
        objProduto.sCodigo = objItemConc.sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 23080 Then gError 86151
        If lErro <> SUCESSO Then gError 86152
        
        iTipoTribItem = 0
        dQuantMaior = 0
        'Para cada Item de Requisisão ligado ao Item de conc
        For Each objItemRCItemConc In objItemConc.colItemRCItemConcorrencia
            bAchou = False
            'Busca o Item de RC nas Requisições
            For Each objReqCompra In gcolRequisicaoCompra
                For Each objItemRC In objReqCompra.colItens
                    'Quando encontrar o item
                    If objItemRC.lNumIntDoc = objItemRCItemConc.lItemReqCompra Then
                        
                        lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemConc.sUM, objItemRC.sUM, dFator)
                        If lErro <> SUCESSO Then gError 86153
                        
                        'Seleciona o item de Requisição de compras
                        objItemRC.iSelecionado = MARCADO
                        objItemRC.dQuantNoPedido = objItemRC.dQuantNoPedido + objItemRCItemConc.dQuantidade * dFator
                        
                        If objItemRCItemConc.dQuantidade > dQuantMaior Then
                            dQuantMaior = objItemRCItemConc.dQuantidade
                            iTipoTribItem = objItemRC.iTipoTributacao
                        End If

                        bAChouQtSup = False

                        For Each objQtSupl In objItemConc.colQuantSuplementar
                            If objQtSupl.iFilialDestino = objReqCompra.iFilialDestino And _
                               objQtSupl.iTipoDestino = objReqCompra.iTipoDestino And _
                               objQtSupl.lFornCliDestino = objReqCompra.lFornCliDestino Then
                               
                               objQtSupl.dQuantRequisitada = objQtSupl.dQuantRequisitada + objItemRCItemConc.dQuantidade
                               objQtSupl.dQuantidade = objQtSupl.dQuantidade + objItemRCItemConc.dQuantidade
                                bAChouQtSup = True
                                Exit For
                            End If
                        Next
                        
                        If Not bAChouQtSup Then
                        
                            Set objQtSupl = New ClassQuantSuplementar
                            
                            objQtSupl.dQuantidade = objItemRCItemConc.dQuantidade
                            objQtSupl.dQuantRequisitada = objItemRCItemConc.dQuantidade
                            objQtSupl.iFilialDestino = objReqCompra.iFilialDestino
                            objQtSupl.iTipoDestino = objReqCompra.iTipoDestino
                            objQtSupl.lFornCliDestino = objReqCompra.lFornCliDestino
                            
                            objItemConc.colQuantSuplementar.Add objQtSupl
                        
                        End If
                        
                        'Marca que a achou o item que estava sendo procurado
                        bAchou = True
                        Exit For
                    End If
                Next
                'Se o Item foi encontrado termina a busca
                If bAchou Then Exit For
            Next
        Next
        For Each objCotItemConc In objItemConc.colCotacaoItemConc
            objCotItemConc.iTipoTributacao = iTipoTribItem
        Next
    Next
    
    Seleciona_ItensRelacionados = SUCESSO
    
    Exit Function
    
Erro_Seleciona_ItensRelacionados:

    Seleciona_ItensRelacionados = Err
    
    Select Case gErr
    
        Case 86151, 86153
        
        Case 86152
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err, objProduto.sCodigo)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154588)
    
    End Select
    
    Exit Function

End Function
Private Sub Inicializa_QuantSupl(objConcorrencia As ClassConcorrencia)

Dim objItemConc As ClassItemConcorrencia
Dim objQuantSupl As ClassQuantSuplementar

    If objConcorrencia.iTipoDestino <> TIPO_DESTINO_AUSENTE Then

        For Each objItemConc In objConcorrencia.colItens

            Set objQuantSupl = New ClassQuantSuplementar

            objQuantSupl.dQuantidade = objItemConc.dQuantidade
            objQuantSupl.iTipoDestino = objConcorrencia.iTipoDestino
            objQuantSupl.iFilialDestino = objConcorrencia.iFilialDestino
            objQuantSupl.lFornCliDestino = objConcorrencia.lFornCliDestino

            objItemConc.colQuantSuplementar.Add objQuantSupl

        Next

    End If
    
    Exit Sub

End Sub

Private Sub Busca_Na_Colecao(collCodigoNome As AdmCollCodigoNome, lCodigo As Long, iPosicao As Integer)
'Busca a chave lCodigo na coleção

Dim objlCodigoNome As AdmlCodigoNome
Dim iIndice As Integer

    iPosicao = 0
    iIndice = 0
    
    'Para cada item da coleção
    For Each objlCodigoNome In collCodigoNome
        
        iIndice = iIndice + 1
        
        'Busca o item com a chave passada
        If objlCodigoNome.lCodigo = lCodigo Then
            
            iPosicao = iIndice
            Exit For
        
        End If
    
    Next
    
    Exit Sub

End Sub


Private Sub Busca_FilialForn(colFilialForn As Collection, lFornecedor As Long, iFilial As Integer, iPosicao As Integer)

Dim objFilialFornecedor As ClassFilialFornecedor
Dim iIndice As Integer

    iPosicao = 0
    
    For iIndice = 1 To colFilialForn.Count
        
        Set objFilialFornecedor = colFilialForn(iIndice)
        If objFilialFornecedor.lCodFornecedor = lFornecedor And objFilialFornecedor.iCodFilial = iFilial Then
            iPosicao = iIndice
            Exit Sub
        End If
    Next
        
    Exit Sub
    
End Sub
Private Function GridCotacoes_Preenche(objConcorrencia As ClassConcorrencia) As Long
'Preenche Grid de Cotações

Dim lErro As Long
Dim iIndice As Integer, iIndice2 As Integer
Dim iIndice3 As Integer
Dim colCampos As New Collection
Dim iCondPagto As Integer
Dim colGeracao As New Collection
Dim dValorPresente As Double
Dim colCotacaoSaida As New Collection
Dim sProdutoMascarado As String
Dim objCotItemConcAux As ClassCotacaoItemConcAux
Dim objItemCotItemConc As ClassCotacaoItemConc
Dim objItemConcorrencia As New ClassItemConcorrencia

On Error GoTo Erro_GridCotacoes_Preenche
    
    Call Grid_Limpa(objGridCotacoes)
           
    For Each objItemConcorrencia In objConcorrencia.colItens
        'Coloca na coleção as cotações que aparecem na tela
         For Each objItemCotItemConc In objItemConcorrencia.colCotacaoItemConc
                
            Set objCotItemConcAux = New ClassCotacaoItemConcAux
            
            Set objCotItemConcAux.objCotacaoItemConc = objItemCotItemConc
            objCotItemConcAux.sCondPagto = objItemCotItemConc.sCondPagto
            objCotItemConcAux.sDescricao = objItemConcorrencia.sDescricao
            objCotItemConcAux.sFilial = objItemCotItemConc.sFilial
            objCotItemConcAux.sFornecedor = objItemCotItemConc.sFornecedor
            objCotItemConcAux.sProduto = objItemConcorrencia.sProduto
            objCotItemConcAux.dtDataNecessidade = objItemConcorrencia.dtDataNecessidade
            objItemCotItemConc.sUMCompra = objItemConcorrencia.sUM
            
            colGeracao.Add objCotItemConcAux
        Next
    Next
    
    'Carrega os campos base para a ordenação utilizados na rotina de ordenação
    colCampos.Add "sProduto"
    colCampos.Add "sCondPagto"
    colCampos.Add "sFornecedor"
    colCampos.Add "sFilial"

    If colGeracao.Count > 0 Then
        lErro = Ordena_Colecao(colGeracao, colCotacaoSaida, colCampos)
        If lErro <> SUCESSO Then gError 63808
    End If
    
    Set colGeracao = colCotacaoSaida
    
    iIndice = 0
    
    For Each objCotItemConcAux In colGeracao

        iIndice = iIndice + 1
        GridCotacoes.TextMatrix(iIndice, iGrid_EscolhidoCot_Col) = objCotItemConcAux.objCotacaoItemConc.iEscolhido
        
        For iIndice3 = 0 To MoedaPC.ListCount - 1
            If objCotItemConcAux.objCotacaoItemConc.iMoeda = Moeda.ItemData(iIndice3) Then
                GridCotacoes.TextMatrix(iIndice, iGrid_MoedaCot_Col) = Moeda.List(iIndice3)
                Exit For
            End If
        Next
        
        If objCotItemConcAux.objCotacaoItemConc.dTaxa > 0 Then
            GridCotacoes.TextMatrix(iIndice, iGrid_TaxaCot_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dTaxa, "STANDARD")
            GridCotacoes.TextMatrix(iIndice, iGrid_Unitario_RS_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dTaxa * objCotItemConcAux.objCotacaoItemConc.dPrecoUnitario, gobjCOM.sFormatoPrecoUnitario) ' "STANDARD") 'Alterado por Wagner
        End If
        
        'Mascara o Produto
        lErro = Mascara_RetornaProdutoEnxuto(objCotItemConcAux.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 68358

        'Preenche o Produto com o ProdutoEnxuto
        Produto1.PromptInclude = False
        Produto1.Text = sProdutoMascarado
        Produto1.PromptInclude = True
        
        GridCotacoes.TextMatrix(iIndice, iGrid_ProdutoCot_Col) = Produto1.Text
        GridCotacoes.TextMatrix(iIndice, iGrid_DescProdutoCot_Col) = objCotItemConcAux.sDescricao
        GridCotacoes.TextMatrix(iIndice, iGrid_CondPagtoCot_Col) = objCotItemConcAux.objCotacaoItemConc.sCondPagto
        
        GridCotacoes.TextMatrix(iIndice, iGrid_QuantComprarCot_Col) = Formata_Estoque(objCotItemConcAux.objCotacaoItemConc.dQuantidadeComprar)

        GridCotacoes.TextMatrix(iIndice, iGrid_UMCot_Col) = objCotItemConcAux.objCotacaoItemConc.sUMCompra
        GridCotacoes.TextMatrix(iIndice, iGrid_PrecoUnitario_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dPrecoAjustado, gobjCOM.sFormatoPrecoUnitario) ' "STANDARD") 'Alterado por Wagner
        
        If objCotItemConcAux.objCotacaoItemConc.sMotivoEscolha <> MOTIVO_EXCLUSIVO_DESCRICAO Then
            Call Calcula_Preferencia(objCotItemConcAux.objCotacaoItemConc, Produto1.Text, objCotItemConcAux.objCotacaoItemConc.dQuantidadeComprar)
            GridCotacoes.TextMatrix(iIndice, iGrid_Preferencia_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dPreferencia, "Percent")
        Else
            GridCotacoes.TextMatrix(iIndice, iGrid_Preferencia_Col) = "Exclusivo"
        End If
                                                    
        GridCotacoes.TextMatrix(iIndice, iGrid_ValorPresenteCot_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dValorPresente, ValorPresente.Format) 'Alterado por Wagner
        GridCotacoes.TextMatrix(iIndice, iGrid_FornecedorCot_Col) = objCotItemConcAux.objCotacaoItemConc.sFornecedor
        GridCotacoes.TextMatrix(iIndice, iGrid_FilialFornCot_Col) = objCotItemConcAux.objCotacaoItemConc.sFilial
        GridCotacoes.TextMatrix(iIndice, iGrid_PedidoCot_Col) = objCotItemConcAux.objCotacaoItemConc.lPedCotacao
        If objCotItemConcAux.objCotacaoItemConc.dQuantEntrega > 0 Then GridCotacoes.TextMatrix(iIndice, iGrid_QuantidadeEntrega_Col) = Formata_Estoque(objCotItemConcAux.objCotacaoItemConc.dQuantEntrega)
        GridCotacoes.TextMatrix(iIndice, iGrid_ValorItem_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dPrecoAjustado * objCotItemConcAux.objCotacaoItemConc.dQuantidadeComprar, "STANDARD") 'Alterado por Wagner
        
        For iIndice2 = 0 To TipoTributacaoCot.ListCount - 1
            If objCotItemConcAux.objCotacaoItemConc.iTipoTributacao = TipoTributacaoCot.ItemData(iIndice2) Then
                GridCotacoes.TextMatrix(iIndice, iGrid_TipoTributacaoCot_Col) = TipoTributacaoCot.List(iIndice2)
                Exit For
            End If
        Next
        
        GridCotacoes.TextMatrix(iIndice, iGrid_AliquotaIPI_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dAliquotaIPI, "Percent")
        GridCotacoes.TextMatrix(iIndice, iGrid_AliquotaICMS_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dAliquotaICMS, "Percent")
        
        'Data da Cotacao
        If objCotItemConcAux.objCotacaoItemConc.dtDataPedidoCotacao <> DATA_NULA Then
            GridCotacoes.TextMatrix(iIndice, iGrid_DataCotacaoCot_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dtDataPedidoCotacao, "dd/mm/yyyy")
        End If
        
        'Data de Validade
        If objCotItemConcAux.objCotacaoItemConc.dtDataValidade <> DATA_NULA Then
            GridCotacoes.TextMatrix(iIndice, iGrid_DataValidadeCot_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dtDataValidade, "dd/mm/yyyy")
        End If

        'Prazo de Entrega
        If objCotItemConcAux.objCotacaoItemConc.iPrazoEntrega <> 0 Then
            GridCotacoes.TextMatrix(iIndice, iGrid_PrazoEntrega_Col) = objCotItemConcAux.objCotacaoItemConc.iPrazoEntrega
            GridCotacoes.TextMatrix(iIndice, iGrid_DataEntrega_Col) = Format(DateAdd("d", objCotItemConcAux.objCotacaoItemConc.iPrazoEntrega, Date), "dd/mm/yyyy")
        End If

        'Data de Entrega
        If objCotItemConcAux.objCotacaoItemConc.dtDataEntrega <> DATA_NULA Then
        End If
                
        'Quantidade a comprar Máxima
        GridCotacoes.TextMatrix(iIndice, iGrid_QuantidadeCot_Col) = Formata_Estoque(objCotItemConcAux.objCotacaoItemConc.dQuantCotada)

        'Motivo escolha
        GridCotacoes.TextMatrix(iIndice, iGrid_MotivoEscolhaCot_Col) = objCotItemConcAux.objCotacaoItemConc.sMotivoEscolha
        
        If objCotItemConcAux.dtDataNecessidade <> DATA_NULA Then
            GridCotacoes.TextMatrix(iIndice, iGrid_DataNecessidade_Col) = Format(objCotItemConcAux.dtDataNecessidade, "dd/mm/yyyy")
        End If
        
        objGridCotacoes.iLinhasExistentes = objGridCotacoes.iLinhasExistentes + 1
    Next

    Call Grid_Refresh_Checkbox(objGridCotacoes)
    
    Exit Function

Erro_GridCotacoes_Preenche:

    Select Case gErr

        Case 62733, 63808, 68358
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154589)

    End Select

    Exit Function

End Function

Private Function Busca_QuantComprar_ItemReq(lReqCompra As Long, iFilialReq As Integer, iItem As Integer, dQuantComprar As Double)

Dim objReqCompra As ClassRequisicaoCompras
Dim objItemRC As ClassItemReqCompras
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim lErro As Long

On Error GoTo Erro_Busca_QuantComprar_ItemReq

    dQuantComprar = 0

    'Para cada Requisição da tela
    For Each objReqCompra In gcolRequisicaoCompra
        'se for a req passada
        If objReqCompra.lCodigo = lReqCompra And objReqCompra.iFilialEmpresa = iFilialReq Then
            'Localiza o item procurado
            For Each objItemRC In objReqCompra.colItens
                If objItemRC.iItem = iItem Then
                    
                    objProduto.sCodigo = objItemRC.sProduto
                    'Lê o produto
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 23080 Then gError 62796
                    If lErro <> SUCESSO Then gError 62797
                    
                    lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemRC.sUM, objProduto.sSiglaUMCompra, dFator)
                    If lErro <> SUCESSO Then gError 62798
                    
                    'COnverte para a UM compra
                    dQuantComprar = objItemRC.dQuantComprar * dFator
                    Exit For
                End If
            Next
        End If
        
    Next
    
    Busca_QuantComprar_ItemReq = SUCESSO

    Exit Function

Erro_Busca_QuantComprar_ItemReq:

    Busca_QuantComprar_ItemReq = gErr
    
    Select Case gErr
    
        Case 62796, 62798
        
        Case 62797
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154590)
            
    End Select

    Exit Function

End Function

Function Carrega_Moeda() As Long

Dim lErro As Long
Dim objMoeda As ClassMoedas
Dim colMoedas As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Carrega_Moeda
    
    lErro = CF("Moedas_Le_Todas", colMoedas)
    If lErro <> SUCESSO Then gError 103371
    
    'se não existem moedas cadastradas
    If colMoedas.Count = 0 Then gError 103372
    
    For Each objMoeda In colMoedas
    
        Moeda.AddItem objMoeda.sNome
        Moeda.ItemData(iIndice) = objMoeda.iCodigo
        
        MoedaPC.AddItem objMoeda.sNome
        MoedaPC.ItemData(iIndice) = objMoeda.iCodigo
        
        iIndice = iIndice + 1
        
    Next

    Carrega_Moeda = SUCESSO
    
    Exit Function
    
Erro_Carrega_Moeda:

    Carrega_Moeda = gErr
    
    Select Case gErr
    
        Case 103371
        
        Case 103372
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDAS_NAO_CADASTRADAS", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154591)
    
    End Select

End Function

'##############################################
'Inserido por Wagner
Private Sub Formata_Controles()

    PrecoUnitario.Format = gobjCOM.sFormatoPrecoUnitario
    ValorRecebido.Format = gobjCOM.sFormatoPrecoUnitario

End Sub
'##############################################


Private Sub BotaoBaixar_Click()

Dim lErro As Long
Dim objConcorrencia As New ClassConcorrencia
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoBaixar_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se existe alguma Requisição de Compras
    If Len(Trim(ConcorrenciaLabel.Caption)) = 0 Then gError 189325

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_BAIXA_CONCORRENCIA", ConcorrenciaLabel.Caption)
    If vbMsgRes = vbNo Then gError 189326

    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objConcorrencia)
    If lErro <> SUCESSO Then gError 189327
    
    'Lê a Concorrencia
    lErro = CF("ConcorrenciaN_Le", objConcorrencia)
    If lErro <> SUCESSO And lErro <> 89865 Then gError 189328

    'Se não encontrou a concorrencia ==> erro
    If lErro = 89865 Then gError 189329
    
    'Baixa a Requisição de Compras
    lErro = CF("Concorrencia_Grava_Baixa", objConcorrencia)
    If lErro <> SUCESSO Then gError 189330

    'Limpa a tela
    Call Rotina_Aviso(vbOKOnly, "AVISO_BAIXA_CONCORRENCIA_SUCESSO")

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoBaixar_Click:

    Select Case gErr

        Case 189325
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CONCORRENCIA_NAO_PREENCHIDO", gErr)

        Case 189326, 189327, 189328, 189330
        
        Case 189329
            Call Rotina_Erro(vbOKOnly, "ERRO_CONCORRENCIA_NAO_CADASTRADA", gErr, objConcorrencia.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189365)

    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

End Sub

