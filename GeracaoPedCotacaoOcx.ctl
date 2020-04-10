VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl GeracaoPedCotacaoOcx 
   ClientHeight    =   9195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   ScaleHeight     =   9195
   ScaleWidth      =   16995
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8205
      Index           =   5
      Left            =   195
      TabIndex        =   93
      Top             =   795
      Visible         =   0   'False
      Width           =   16605
      Begin VB.CommandButton BotaoFornecedor 
         Caption         =   "Filial Fornecedor..."
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
         Left            =   6750
         Picture         =   "GeracaoPedCotacaoOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   116
         Top             =   7560
         Width           =   2085
      End
      Begin VB.ComboBox OrdemFornecedor 
         Height          =   315
         ItemData        =   "GeracaoPedCotacaoOcx.ctx":0DA2
         Left            =   2820
         List            =   "GeracaoPedCotacaoOcx.ctx":0DA4
         Style           =   2  'Dropdown List
         TabIndex        =   95
         Top             =   120
         Width           =   2325
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         Height          =   600
         Index           =   5
         Left            =   210
         Picture         =   "GeracaoPedCotacaoOcx.ctx":0DA6
         Style           =   1  'Graphical
         TabIndex        =   114
         Top             =   7560
         Width           =   2085
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         Height          =   600
         Index           =   5
         Left            =   2490
         Picture         =   "GeracaoPedCotacaoOcx.ctx":1DC0
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   7560
         Width           =   2085
      End
      Begin VB.Frame Frame15 
         Caption         =   "Fornecedores"
         Height          =   6990
         Left            =   120
         TabIndex        =   96
         Top             =   450
         Width           =   16320
         Begin MSMask.MaskEdBox DataUltimaCotacao 
            Height          =   225
            Left            =   6615
            TabIndex        =   104
            Top             =   300
            Visible         =   0   'False
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdutoForn 
            Height          =   225
            Left            =   450
            TabIndex        =   99
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.CheckBox EscolhidoForn 
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
            Left            =   60
            TabIndex        =   98
            Top             =   360
            Width           =   750
         End
         Begin VB.TextBox ObservacaoForn 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   90
            MaxLength       =   100
            TabIndex        =   107
            Top             =   3315
            Width           =   1635
         End
         Begin VB.TextBox DescProdutoForn 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   1770
            MaxLength       =   50
            TabIndex        =   100
            Top             =   345
            Width           =   4000
         End
         Begin VB.CheckBox Exclusivo 
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
            Left            =   6360
            TabIndex        =   103
            Top             =   360
            Width           =   780
         End
         Begin MSMask.MaskEdBox QuantRecebidaForn 
            Height          =   225
            Left            =   4470
            TabIndex        =   111
            Top             =   3225
            Visible         =   0   'False
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantPedidaForn 
            Height          =   225
            Left            =   3240
            TabIndex        =   110
            Top             =   3240
            Visible         =   0   'False
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PrazoEntrega 
            Height          =   225
            Left            =   1920
            TabIndex        =   109
            Top             =   3240
            Visible         =   0   'False
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   3
            Format          =   "###"
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox UltimaCotacao 
            Height          =   225
            Left            =   7890
            TabIndex        =   105
            Top             =   315
            Visible         =   0   'False
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox SaldoTitulos 
            Height          =   225
            Left            =   7035
            TabIndex        =   113
            Top             =   3225
            Visible         =   0   'False
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CondicaoPagto 
            Height          =   225
            Left            =   5730
            TabIndex        =   112
            Top             =   3240
            Visible         =   0   'False
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   30
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataUltimaCompra 
            Height          =   225
            Left            =   750
            TabIndex        =   108
            Top             =   3240
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TipoFrete 
            Height          =   225
            Left            =   240
            TabIndex        =   106
            Top             =   3255
            Visible         =   0   'False
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   3
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FilialFornGrid 
            Height          =   225
            Left            =   5040
            TabIndex        =   102
            Top             =   330
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FornecedorGrid 
            Height          =   225
            Left            =   3240
            TabIndex        =   101
            Top             =   360
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
         Begin MSFlexGridLib.MSFlexGrid GridFornecedores 
            Height          =   6540
            Left            =   120
            TabIndex        =   97
            Top             =   300
            Width           =   16080
            _ExtentX        =   28363
            _ExtentY        =   11536
            _Version        =   393216
            Rows            =   12
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.Label Label45 
         Caption         =   "Ordena por:"
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
         Left            =   1710
         TabIndex        =   94
         Top             =   150
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8220
      Index           =   4
      Left            =   165
      TabIndex        =   80
      Top             =   765
      Visible         =   0   'False
      Width           =   16680
      Begin VB.Frame Frame4 
         Caption         =   "Produtos"
         Height          =   7440
         Left            =   240
         TabIndex        =   81
         Top             =   150
         Width           =   16305
         Begin VB.CheckBox EscolhidoProd 
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
            Left            =   75
            TabIndex        =   83
            Top             =   225
            Width           =   930
         End
         Begin VB.TextBox DescProduto 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   85
            Top             =   480
            Width           =   4000
         End
         Begin MSMask.MaskEdBox UnidadeMedProd 
            Height          =   240
            Left            =   2940
            TabIndex        =   86
            Top             =   465
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FilialFornProd 
            Height          =   225
            Left            =   7200
            TabIndex        =   89
            Top             =   480
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
         Begin MSMask.MaskEdBox FornecedorProd 
            Height          =   225
            Left            =   5280
            TabIndex        =   88
            Top             =   480
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
         Begin MSMask.MaskEdBox QuantidadeProd 
            Height          =   225
            Left            =   4500
            TabIndex        =   87
            Top             =   480
            Width           =   990
            _ExtentX        =   1746
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
         Begin MSMask.MaskEdBox Produto 
            Height          =   225
            Left            =   120
            TabIndex        =   84
            Top             =   480
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridProdutos 
            Height          =   6945
            Left            =   90
            TabIndex        =   82
            Top             =   315
            Width           =   16110
            _ExtentX        =   28416
            _ExtentY        =   12250
            _Version        =   393216
            Rows            =   12
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         Height          =   570
         Index           =   4
         Left            =   270
         Picture         =   "GeracaoPedCotacaoOcx.ctx":2FA2
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   7605
         Width           =   1425
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         Height          =   570
         Index           =   4
         Left            =   1890
         Picture         =   "GeracaoPedCotacaoOcx.ctx":3FBC
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   7605
         Width           =   1425
      End
      Begin VB.CommandButton BotaoProduto 
         Caption         =   "Produto..."
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
         Left            =   7125
         TabIndex        =   92
         Top             =   7710
         Width           =   1680
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8175
      Index           =   2
      Left            =   150
      TabIndex        =   42
      Top             =   765
      Visible         =   0   'False
      Width           =   16590
      Begin VB.CommandButton BotaoRequisicao 
         Caption         =   "Requisição..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6900
         TabIndex        =   131
         Top             =   7710
         Width           =   1875
      End
      Begin VB.ComboBox OrdemRequisicao 
         Height          =   315
         ItemData        =   "GeracaoPedCotacaoOcx.ctx":519E
         Left            =   3390
         List            =   "GeracaoPedCotacaoOcx.ctx":51A0
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   60
         Width           =   2325
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         Height          =   570
         Index           =   2
         Left            =   180
         Picture         =   "GeracaoPedCotacaoOcx.ctx":51A2
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   7530
         Width           =   1425
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         Height          =   570
         Index           =   2
         Left            =   1800
         Picture         =   "GeracaoPedCotacaoOcx.ctx":61BC
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   7530
         Width           =   1425
      End
      Begin VB.Frame Frame7 
         Caption         =   "Requisições de Compra"
         Height          =   7065
         Left            =   240
         TabIndex        =   45
         Top             =   390
         Width           =   16215
         Begin MSMask.MaskEdBox CodigoPV 
            Height          =   240
            Left            =   2340
            TabIndex        =   136
            Top             =   930
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.CheckBox EscolhidoReq 
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
            TabIndex        =   47
            Top             =   270
            Width           =   615
         End
         Begin MSMask.MaskEdBox FilialReq 
            Height          =   225
            Left            =   195
            TabIndex        =   132
            Top             =   525
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
         Begin VB.TextBox ObservacaoReq 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   6015
            MaxLength       =   255
            TabIndex        =   54
            Top             =   255
            Width           =   4000
         End
         Begin VB.CheckBox Urgente 
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
            Left            =   3315
            TabIndex        =   51
            Top             =   525
            Width           =   735
         End
         Begin MSMask.MaskEdBox Requisitante 
            Height          =   240
            Left            =   3870
            TabIndex        =   52
            Top             =   270
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
            Left            =   5715
            TabIndex        =   53
            Top             =   495
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
            Left            =   1710
            TabIndex        =   49
            Top             =   525
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
         Begin MSMask.MaskEdBox Requisicao 
            Height          =   225
            Left            =   1050
            TabIndex        =   48
            Top             =   210
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
            Left            =   2580
            TabIndex        =   50
            Top             =   255
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
            Height          =   6570
            Left            =   165
            TabIndex        =   46
            Top             =   315
            Width           =   15900
            _ExtentX        =   28046
            _ExtentY        =   11589
            _Version        =   393216
            Rows            =   16
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.Label Label57 
         Caption         =   "Ordena por:"
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
         Left            =   2250
         TabIndex        =   43
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8280
      Index           =   6
      Left            =   150
      TabIndex        =   117
      Top             =   705
      Visible         =   0   'False
      Width           =   16575
      Begin VB.CommandButton BotaoEmail 
         Caption         =   "Gera e Envia por Email Pedidos de Cotação "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   870
         TabIndex        =   135
         Top             =   5370
         Width           =   4410
      End
      Begin VB.ListBox CondPagtos 
         Height          =   5520
         Left            =   7530
         TabIndex        =   127
         Top             =   540
         Width           =   3135
      End
      Begin VB.CommandButton BotaoImprimePedidos 
         Caption         =   "Gera e Imprime Pedidos de Cotação "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   870
         TabIndex        =   125
         Top             =   4470
         Width           =   4410
      End
      Begin VB.Frame Frame11 
         Caption         =   "Geração"
         Height          =   2820
         Left            =   345
         TabIndex        =   118
         Top             =   465
         Width           =   6435
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   2790
            Picture         =   "GeracaoPedCotacaoOcx.ctx":739E
            Style           =   1  'Graphical
            TabIndex        =   133
            ToolTipText     =   "Numeração Automática"
            Top             =   345
            Width           =   300
         End
         Begin MSMask.MaskEdBox Descricao 
            Height          =   1425
            Left            =   1980
            TabIndex        =   123
            Top             =   1200
            Width           =   4260
            _ExtentX        =   7514
            _ExtentY        =   2514
            _Version        =   393216
            MaxLength       =   50
            PromptChar      =   "_"
         End
         Begin VB.Label Cotacao 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1995
            TabIndex        =   134
            Top             =   330
            Width           =   795
         End
         Begin VB.Label CondPagto 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1980
            TabIndex        =   121
            Top             =   780
            Width           =   3255
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            Caption         =   "Cond Pagto A Prazo:"
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
            TabIndex        =   120
            Top             =   810
            Width           =   1785
         End
         Begin VB.Label Label54 
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
            Left            =   1020
            TabIndex        =   122
            Top             =   1230
            Width           =   930
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Nº Geração:"
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
            Left            =   870
            TabIndex        =   119
            Top             =   390
            Width           =   1065
         End
      End
      Begin VB.CommandButton BotaoGeraPedidos 
         Caption         =   "Gera Pedidos de Cotação"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   870
         TabIndex        =   124
         Top             =   3585
         Width           =   4410
      End
      Begin VB.Label Label40 
         Caption         =   "Condições de Pagamento"
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
         Left            =   7515
         TabIndex        =   126
         Top             =   285
         Width           =   2265
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8190
      Index           =   3
      Left            =   150
      TabIndex        =   57
      Top             =   765
      Visible         =   0   'False
      Width           =   16650
      Begin VB.Frame Frame5 
         Caption         =   "Itens de Requisições"
         Height          =   7395
         Left            =   240
         TabIndex        =   59
         Top             =   90
         Width           =   16260
         Begin MSMask.MaskEdBox FilialEmpresaItemReq 
            Height          =   225
            Left            =   735
            TabIndex        =   61
            Top             =   210
            Visible         =   0   'False
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Almoxarifado 
            Height          =   225
            Left            =   1560
            TabIndex        =   72
            Top             =   3630
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
         Begin MSMask.MaskEdBox QuantCanceladaItem 
            Height          =   225
            Left            =   435
            TabIndex        =   71
            Top             =   3585
            Width           =   1395
            _ExtentX        =   2461
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
         Begin VB.TextBox DescProdutoItem 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   3315
            MaxLength       =   50
            TabIndex        =   65
            Top             =   210
            Width           =   4000
         End
         Begin VB.CheckBox EscolhidoItem 
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
            Left            =   495
            TabIndex        =   60
            Top             =   345
            Width           =   855
         End
         Begin VB.TextBox ObservacaoItem 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   7185
            MaxLength       =   255
            TabIndex        =   77
            Top             =   3600
            Width           =   4000
         End
         Begin MSMask.MaskEdBox CclItem 
            Height          =   225
            Left            =   2700
            TabIndex        =   73
            Top             =   3615
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FornecedorItem 
            Height          =   225
            Left            =   3375
            TabIndex        =   74
            Top             =   3630
            Visible         =   0   'False
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Item 
            Height          =   225
            Left            =   1635
            TabIndex        =   63
            Top             =   330
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
            Left            =   915
            TabIndex        =   62
            Top             =   675
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ExclusivoItem 
            Height          =   225
            Left            =   5940
            TabIndex        =   76
            Top             =   3615
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
         Begin MSMask.MaskEdBox UnidadeMedItem 
            Height          =   225
            Left            =   4800
            TabIndex        =   66
            Top             =   225
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   5
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantCotarItem 
            Height          =   225
            Left            =   5940
            TabIndex        =   67
            Top             =   360
            Width           =   750
            _ExtentX        =   1323
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
         Begin MSMask.MaskEdBox QuantRecebidaItem 
            Height          =   225
            Left            =   105
            TabIndex        =   70
            Top             =   3495
            Width           =   1365
            _ExtentX        =   2408
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
         Begin MSMask.MaskEdBox QuantPedidaItem 
            Height          =   225
            Left            =   7035
            TabIndex        =   69
            Top             =   255
            Width           =   1155
            _ExtentX        =   2037
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
         Begin MSMask.MaskEdBox FilialFornItem 
            Height          =   225
            Left            =   4710
            TabIndex        =   75
            Top             =   3630
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
         Begin MSMask.MaskEdBox QuantidadeItem 
            Height          =   225
            Left            =   6975
            TabIndex        =   68
            Top             =   360
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
         Begin MSMask.MaskEdBox ProdutoItem 
            Height          =   225
            Left            =   2070
            TabIndex        =   64
            Top             =   195
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridItensRequisicoes 
            Height          =   7005
            Left            =   135
            TabIndex        =   58
            Top             =   270
            Width           =   16050
            _ExtentX        =   28310
            _ExtentY        =   12356
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
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         Height          =   570
         Index           =   3
         Left            =   240
         Picture         =   "GeracaoPedCotacaoOcx.ctx":7488
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   7605
         Width           =   1425
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         Height          =   570
         Index           =   3
         Left            =   2010
         Picture         =   "GeracaoPedCotacaoOcx.ctx":84A2
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   7590
         Width           =   1425
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8205
      Index           =   1
      Left            =   150
      TabIndex        =   1
      Top             =   780
      Width           =   16500
      Begin VB.Frame Frame6 
         Caption         =   "Exibe Requisições"
         Height          =   7500
         Left            =   255
         TabIndex        =   2
         Top             =   165
         Width           =   10365
         Begin VB.Frame Frame8 
            Caption         =   "Data Registro"
            Height          =   1380
            Left            =   3315
            TabIndex        =   17
            Top             =   3060
            Width           =   2385
            Begin MSComCtl2.UpDown UpDownDataDe 
               Height          =   300
               Left            =   1845
               TabIndex        =   20
               Top             =   360
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown UpDownDataAte 
               Height          =   300
               Left            =   1860
               TabIndex        =   23
               Top             =   870
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataDe 
               Height          =   300
               Left            =   720
               TabIndex        =   19
               Top             =   360
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox DataAte 
               Height          =   300
               Left            =   720
               TabIndex        =   22
               Top             =   870
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin VB.Label Label11 
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
               Left            =   255
               TabIndex        =   18
               Top             =   420
               Width           =   315
            End
            Begin VB.Label Label2 
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
               Left            =   285
               TabIndex        =   21
               Top             =   960
               Width           =   360
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Número"
            Height          =   1365
            Left            =   6240
            TabIndex        =   24
            Top             =   3060
            Width           =   2985
            Begin MSMask.MaskEdBox CodigoDe 
               Height          =   315
               Left            =   780
               TabIndex        =   26
               Top             =   345
               Width           =   810
               _ExtentX        =   1429
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CodigoAte 
               Height          =   315
               Left            =   780
               TabIndex        =   28
               Top             =   915
               Width           =   810
               _ExtentX        =   1429
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin VB.Label Label12 
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
               Left            =   375
               TabIndex        =   27
               Top             =   975
               Width           =   360
            End
            Begin VB.Label Label14 
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
               Left            =   375
               TabIndex        =   25
               Top             =   405
               Width           =   315
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Data Limite"
            Height          =   1380
            Left            =   375
            TabIndex        =   10
            Top             =   3060
            Width           =   2385
            Begin MSComCtl2.UpDown UpDownDataLimDe 
               Height          =   300
               Left            =   1845
               TabIndex        =   13
               Top             =   360
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown UpDownDataLimAte 
               Height          =   300
               Left            =   1860
               TabIndex        =   16
               Top             =   900
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataLimiteDe 
               Height          =   300
               Left            =   720
               TabIndex        =   12
               Top             =   360
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox DataLimiteAte 
               Height          =   300
               Left            =   720
               TabIndex        =   15
               Top             =   885
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin VB.Label Label13 
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
               Left            =   285
               TabIndex        =   14
               Top             =   960
               Width           =   360
            End
            Begin VB.Label Label17 
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
               Left            =   255
               TabIndex        =   11
               Top             =   420
               Width           =   315
            End
         End
         Begin VB.CheckBox ExibeCotadas 
            Caption         =   "Exibe Itens Requisições já cotados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   375
            TabIndex        =   5
            Top             =   960
            Width           =   3360
         End
         Begin VB.Frame Frame3 
            Caption         =   "Local de Entrega"
            Height          =   1575
            Left            =   360
            TabIndex        =   29
            Top             =   4545
            Width           =   8880
            Begin VB.Frame FrameDestino 
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   720
               Index           =   0
               Left            =   3555
               TabIndex        =   34
               Top             =   765
               Width           =   3645
               Begin VB.ComboBox FilialEmpresa 
                  Height          =   315
                  Left            =   1245
                  TabIndex        =   36
                  Top             =   240
                  Width           =   2160
               End
               Begin VB.Label FilialEmpresaLabel 
                  AutoSize        =   -1  'True
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
                  Height          =   195
                  Left            =   720
                  TabIndex        =   35
                  Top             =   270
                  Width           =   465
               End
            End
            Begin VB.Frame Frame2 
               Caption         =   "Tipo"
               Height          =   585
               Left            =   3390
               TabIndex        =   31
               Top             =   135
               Width           =   4065
               Begin VB.OptionButton TipoDestino 
                  Caption         =   "Filial Empresa"
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
                  Index           =   0
                  Left            =   375
                  TabIndex        =   32
                  Top             =   240
                  Width           =   1515
               End
               Begin VB.OptionButton TipoDestino 
                  Caption         =   "Fornecedor"
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
                  Index           =   1
                  Left            =   2310
                  TabIndex        =   33
                  Top             =   240
                  Width           =   1335
               End
            End
            Begin VB.CheckBox SelecionaDestino 
               Caption         =   "Seleciona Local Entrega"
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
               Left            =   300
               TabIndex        =   30
               Top             =   330
               Width           =   2445
            End
            Begin VB.Frame FrameDestino 
               BorderStyle     =   0  'None
               Height          =   705
               Index           =   1
               Left            =   3570
               TabIndex        =   37
               Top             =   780
               Visible         =   0   'False
               Width           =   3645
               Begin VB.ComboBox FilialFornec 
                  Height          =   315
                  Left            =   1230
                  TabIndex        =   41
                  Top             =   375
                  Width           =   2160
               End
               Begin MSMask.MaskEdBox Fornecedor 
                  Height          =   300
                  Left            =   1230
                  TabIndex        =   39
                  Top             =   0
                  Width           =   2145
                  _ExtentX        =   3784
                  _ExtentY        =   529
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   20
                  PromptChar      =   " "
               End
               Begin VB.Label FilialFornLabel 
                  AutoSize        =   -1  'True
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
                  Height          =   195
                  Left            =   690
                  TabIndex        =   40
                  Top             =   405
                  Width           =   465
               End
               Begin VB.Label FornecedorLabel 
                  AutoSize        =   -1  'True
                  Caption         =   "Fornecedor:"
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
                  Left            =   135
                  MousePointer    =   14  'Arrow and Question
                  TabIndex        =   38
                  Top             =   60
                  Width           =   1035
               End
            End
         End
         Begin VB.CommandButton BotaoDesmarcarTodos 
            Caption         =   "Desmarcar Todos"
            Height          =   570
            Index           =   1
            Left            =   7770
            Picture         =   "GeracaoPedCotacaoOcx.ctx":9684
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   960
            Width           =   1425
         End
         Begin VB.CommandButton BotaoMarcarTodos 
            Caption         =   "Marcar Todos"
            Height          =   570
            Index           =   1
            Left            =   7770
            Picture         =   "GeracaoPedCotacaoOcx.ctx":A866
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   345
            Width           =   1425
         End
         Begin VB.ListBox TipoProduto 
            Height          =   2310
            Left            =   3780
            Style           =   1  'Checkbox
            TabIndex        =   7
            Top             =   360
            Width           =   3885
         End
         Begin VB.Label Comprador 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1350
            TabIndex        =   4
            Top             =   435
            Width           =   2145
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Comprador:"
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
            Left            =   285
            TabIndex        =   3
            Top             =   465
            Width           =   975
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Tipos de Produto"
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
            Left            =   3780
            TabIndex        =   6
            Top             =   135
            Width           =   1470
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   15720
      ScaleHeight     =   495
      ScaleWidth      =   1110
      TabIndex        =   128
      TabStop         =   0   'False
      Top             =   75
      Width           =   1170
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   615
         Picture         =   "GeracaoPedCotacaoOcx.ctx":B880
         Style           =   1  'Graphical
         TabIndex        =   130
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   90
         Picture         =   "GeracaoPedCotacaoOcx.ctx":B9FE
         Style           =   1  'Graphical
         TabIndex        =   129
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   8625
      Left            =   135
      TabIndex        =   0
      Top             =   435
      Width           =   16755
      _ExtentX        =   29554
      _ExtentY        =   15214
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Requisições"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Itens de Requisições"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Produtos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Fornecedores"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Geração"
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
Attribute VB_Name = "GeracaoPedCotacaoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'CONSTANTES GLOBAIS
Const TAB_Selecao = 1
Const TAB_REQUISICAO = 2
Const TAB_ITENSREQ = 3
Const TAB_Produtos = 4
Const TAB_FORNECEDOR = 5
Const TAB_GERACAO = 6
Const PRODUTO_ENCONTRADO = 1
Const PRODUTO_NAO_ENCONTRADO = 0
Const NUM_MAX_REQUISICOES_GRID = 100
'Const NUM_MAX_ITENS_REQUISICOES = 150

Const ORDEM_DATALIMITE = 0
Const ORDEM_URGENTE = 1
Const ORDEM_DATA = 2
Const ORDEM_CODIGO_REQUISICAO = 3
Const ORDEM_CCL = 4

'=========================
'Condicao de pagamentos:
'Só pode ter duas
'Uma é a vista
'A outra o usuário escolhe
'========================
'IMPORTANTE:
'Para cada Produto, Quantidade A Cotar >= Soma dos ítens de Requisições.
'======================
'O mesmo Produto aparece no GridProdutos n vezes
'Uma vez para "sem Filial Fornecedor definido"
'n-1 vezes para cada Filial Fornecedor definido
'========================
'A combo Ordenacao no Tab Requisições deve ter as
'seguintes ordens: DataLimite (Urgencia ordem 2), Urgente (DataLimite ordem 2), Data (DataLimite ordem 2), Codigo, Requisitante (DataLimite ordem 2), Ccl (DataLimite ordem 2)

'Variáveis Globais
Public giAlterado As Integer
Dim giFrameAtual As Integer
Dim giFrameDestinoAtual As Integer
Dim gsOrdemRequisicao As String
Dim asOrdemRequisicao(4) As String
Dim asOrdemRequisicaoString(4) As String
Dim gsOrdemFornecedor As String
Dim asOrdemFornecedor(3) As String
Dim asOrdemFornecedorString(3) As String
Dim giFornecedorAlterado As Integer
Dim giTabSelecao_Alterado As Integer
Dim giTabRequisicao_Alterado As Integer
Dim giTabItens_Alterado As Integer
Dim giTabFornecedor_Alterado As Integer
Dim giTabProdutos_Alterado As Integer
Dim giPodeAumentarQuant As Integer

'Eventos da Tela
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1

'Grids da Tela
'GridRequisicoes
Dim objGridRequisicoes As AdmGrid
Dim iGrid_FilialReq_Col As Integer
Dim iGrid_EscolhidoReq_Col As Integer
Dim iGrid_Numero_Col As Integer
Dim iGrid_DataLimite_Col As Integer
Dim iGrid_DataRC_Col As Integer
Dim iGrid_Urgente_Col As Integer
Dim iGrid_Requisitante_Col As Integer
Dim iGrid_CclReq_Col As Integer
Dim iGrid_Observacao_Col As Integer
Dim iGrid_CodigoPV_Col As Integer

'GridProdutos
Dim objGridProdutos As AdmGrid
Dim iGrid_EscolhidoProd_Col As Integer
Dim iGrid_Produto_Col As Integer
Dim iGrid_Descricao_Col As Integer
Dim iGrid_UnidadeMed_Col As Integer
Dim iGrid_QuantidadeProd_Col As Integer
Dim iGrid_FornecedorProd_Col As Integer
Dim iGrid_FilialFornProd_Col As Integer

'GridItensRequisicoes
Dim objGridItensRequisicoes As AdmGrid
Dim iGrid_EscolhidoItem_Col As Integer
Dim iGrid_Requisicao_Col As Integer
Dim iGrid_FilialEmpresaItemReq_Col As Integer
Dim iGrid_Item_Col As Integer
Dim iGrid_ProdutoItem_Col As Integer
Dim iGrid_DescProdutoItem_Col As Integer
Dim iGrid_UnidadeMedItem_Col As Integer
Dim iGrid_QuantCotarItem_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_QuantPedida_Col As Integer
Dim iGrid_QuantRecebida_Col As Integer
Dim iGrid_QuantCancelada_Col As Integer
Dim iGrid_Almoxarifado_Col As Integer
Dim iGrid_CclItem_Col As Integer
Dim iGrid_FornecedorItem_Col As Integer
Dim iGrid_FilialFornItem_Col As Integer
Dim iGrid_ExclusoItem_Col As Integer
Dim iGrid_ObservacaoItem_Col As Integer

'GridFornecedores
Dim objGridFornecedores As AdmGrid
Dim iGrid_EscolhidoForn_Col As Integer
Dim iGrid_ProdutoForn_Col As Integer
Dim iGrid_DescProdutoForn_Col As Integer
Dim iGrid_FornecedorGrid_Col As Integer
Dim iGrid_FilialFornGrid_Col As Integer
Dim iGrid_Exclusivo_Col As Integer
Dim iGrid_UltimaCotacao_Col As Integer
Dim iGrid_ValorCotacao_Col As Integer
Dim iGrid_Frete_Col As Integer
Dim iGrid_UltimaCompra_Col As Integer
Dim iGrid_PrazoEntrega_Col As Integer
Dim iGrid_QuantPedidaForn_Col As Integer
Dim iGrid_QuantRecebidaForn_Col As Integer
Dim iGrid_CondicaoPagto_Col As Integer
Dim iGrid_SaldoTitulos_Col As Integer
Dim iGrid_ObservacaoForn_Col As Integer

'OBJETOS GLOBAIS DA TELA
'Coleção de objCotacaoProduto do tipo ClassCotacaoProduto
Dim gobjGeracaoCotacao As ClassGeracaoCotacao
Dim gcolCotacaoProduto As New Collection
Dim gcolFornecedorProdutoFF As New Collection
Dim gobjCotacao As ClassCotacao
Dim gcolItemGridFornecedores As New Collection
Dim iFrameAtual As Integer
Dim gcolPedidoCotacao As Collection

Public Sub Form_Load()

Dim lErro As Long
Dim objComprador As New ClassComprador
Dim objUsuario As New ClassUsuario
Dim iFilialEmpresa As Integer

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
   
    'Inicializa as variáveis globais
    Set gobjGeracaoCotacao = New ClassGeracaoCotacao
    
    Set objEventoFornecedor = New AdmEvento
    Set objGridProdutos = New AdmGrid
    Set objGridFornecedores = New AdmGrid
    Set objGridItensRequisicoes = New AdmGrid
    Set objGridRequisicoes = New AdmGrid
    Set gcolFornecedorProdutoFF = New Collection

    objComprador.sCodUsuario = gsUsuario

    'Verifica se gsUsuario é comprador
    lErro = CF("Comprador_Le_Usuario", objComprador)
    If lErro <> SUCESSO And lErro <> 50059 Then gError 63409

    'Se gsUsuario nao é comprador==> erro
    If lErro = 50059 Then gError 63410
    
    giPodeAumentarQuant = objComprador.iAumentaQuant

    objUsuario.sCodUsuario = objComprador.sCodUsuario

    'Lê o usuário
    lErro = CF("Usuario_Le", objUsuario)
    If lErro <> SUCESSO And lErro <> 36347 Then gError 63411
    'Se não encontrou ==>erro
    If lErro = 36347 Then gError 63412

    'Coloca o Nome Reduzido do Comprador na tela
    Comprador.Caption = objUsuario.sNomeReduzido

    'Carrega a ListBox TipoProduto com Tipos de Produto que possam ser comprados
    lErro = Carrega_TipoProduto()
    If lErro <> SUCESSO Then gError 63414

    'Carrega a listbox CondPagtos
    'lErro = Carrega_CondicaoPagamento()
    lErro = CF("Carrega_CondicaoPagamento", CondPagtos, MODULO_CONTASAPAGAR)
    If lErro <> SUCESSO Then gError 63415

    'Inicializa o GridRequisicoes
    lErro = Inicializa_Grid_Requisicoes(objGridRequisicoes)
    If lErro <> SUCESSO Then gError 63417

    'Inicializa o GridItensRequisicoes
    lErro = Inicializa_Grid_ItensRequisicoes(objGridItensRequisicoes)
    If lErro <> SUCESSO Then gError 63418

    'Inicializa o GridProdutos
    lErro = Inicializa_Grid_Produtos(objGridProdutos)
    If lErro <> SUCESSO Then gError 63419

    'Inicializa o GridFornecedores
    lErro = Inicializa_Grid_Fornecedores(objGridFornecedores)
    If lErro <> SUCESSO Then gError 63420

    'Preenche a combo de OrdemRequisicao
    Call OrdemRequisicao_Carrega

    'Preenche a combo de OrdemFornecedor
    Call OrdemFornecedor_Carrega

    'Inicializa a máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 63533

    'Inicializa a máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoItem)
    If lErro <> SUCESSO Then gError 89159
    
    'Inicializa a máscara de ProdutoForn
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoForn)
    If lErro <> SUCESSO Then gError 63528

    'Inicializa mascara do Ccl
    lErro = Inicializa_MascaraCcl()
    If lErro <> SUCESSO Then gError 63579

    'Coloca as Quantidades da tela no formato de Estoque
    QuantCanceladaItem.Format = FORMATO_ESTOQUE
    QuantCotarItem.Format = FORMATO_ESTOQUE
    QuantidadeItem.Format = FORMATO_ESTOQUE
    QuantidadeProd.Format = FORMATO_ESTOQUE
    QuantPedidaForn.Format = FORMATO_ESTOQUE
    QuantPedidaItem.Format = FORMATO_ESTOQUE
    QuantRecebidaForn.Format = FORMATO_ESTOQUE
    QuantRecebidaItem.Format = FORMATO_ESTOQUE

    'Carrega a combo FilialEmpresa
    lErro = Carrega_FilialEmpresa()
    If lErro <> SUCESSO Then gError 63422

    'Coloca FiliaEmpresa Default na Tela
    iFilialEmpresa = giFilialEmpresa
    
    lErro = CF("FilialEmpresa_Customiza", iFilialEmpresa)
    If lErro <> SUCESSO Then gError 126944
    
    FilialEmpresa.Text = iFilialEmpresa
    Call FilialEmpresa_Validate(bSGECancelDummy)
    
    SelecionaDestino.Value = vbChecked

    TipoDestino(TIPO_DESTINO_EMPRESA).Value = True

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case gErr

        Case 63409, 63411, 63414, 63415, 63417, 63418, 126944
            'Erros tratados nas rotinas chamadas

        Case 63419, 63420, 63422, 63528, 63533, 63579
            'Erros tratados nas rotinas chamadas

        Case 63410
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_COMPRADOR", gErr, objUsuario.sCodUsuario)

        Case 634512
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", gErr, objUsuario.sCodUsuario)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161321)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_MascaraCcl() As Long
'Inicializa a mascara do centro de custo

Dim sMascaraCcl As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_mascaraccl

    sMascaraCcl = String(STRING_CCL, 0)

    'Lê a máscara dos centros de custo/lucro
    lErro = MascaraCcl(sMascaraCcl)
    If lErro <> SUCESSO Then gError 63584

    CclItem.Mask = sMascaraCcl
    CclReq.Mask = sMascaraCcl

    Inicializa_MascaraCcl = SUCESSO

    Exit Function

Erro_Inicializa_mascaraccl:

    Inicializa_MascaraCcl = gErr

    Select Case gErr

        Case 63584
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161322)

    End Select

    Exit Function

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
     Call Tela_QueryUnload(Me, giAlterado, Cancel, UnloadMode)
End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    'libera as variaveis globais
    Set objEventoFornecedor = Nothing

    Set objGridRequisicoes = Nothing
    Set objGridItensRequisicoes = Nothing
    Set objGridProdutos = Nothing
    Set objGridFornecedores = Nothing
    Set gobjGeracaoCotacao = Nothing

    Set gcolCotacaoProduto = Nothing
    Set gcolFornecedorProdutoFF = Nothing
    Set gobjCotacao = Nothing
    Set gcolItemGridFornecedores = Nothing
    Set gcolFornecedorProdutoFF = Nothing
    
    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161323)

    End Select

    Exit Sub

End Sub

Private Function Carrega_FilialEmpresa() As Long
'Carrega a combobox FilialEmpresa

Dim lErro As Long
Dim objCodigoNome As AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_FilialEmpresa

    'Lê o Código e o Nome de toda FilialEmpresa do BD
    lErro = CF("Cod_Nomes_Le_FilEmp", colCodigoNome)
    If lErro <> SUCESSO Then gError 63423

    'Carrega a combo de Filial Empresa com código e nome
    For Each objCodigoNome In colCodigoNome
        FilialEmpresa.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        FilialEmpresa.ItemData(FilialEmpresa.NewIndex) = objCodigoNome.iCodigo
    Next

    Carrega_FilialEmpresa = SUCESSO

    Exit Function

Erro_Carrega_FilialEmpresa:

    Carrega_FilialEmpresa = gErr

    Select Case gErr

        Case 63423
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161324)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Fornecedores(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Fornecedores

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Fornecedores

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Escolhido")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Exclusivo")
    objGridInt.colColuna.Add ("Última Cotação")
    objGridInt.colColuna.Add ("Valor Cotação")
    objGridInt.colColuna.Add ("Frete")
    objGridInt.colColuna.Add ("Última Compra")
    objGridInt.colColuna.Add ("Prazo Entrega")
    objGridInt.colColuna.Add ("Quant. Pedida")
    objGridInt.colColuna.Add ("Quant. Recebida")
    objGridInt.colColuna.Add ("Condição Pagto")
    objGridInt.colColuna.Add ("Saldo Tit. a Pagar")
    objGridInt.colColuna.Add ("Observação")

    'campos de edição do grid
    objGridInt.colCampo.Add (EscolhidoForn.Name)
    objGridInt.colCampo.Add (ProdutoForn.Name)
    objGridInt.colCampo.Add (DescProdutoForn.Name)
    objGridInt.colCampo.Add (FornecedorGrid.Name)
    objGridInt.colCampo.Add (FilialFornGrid.Name)
    objGridInt.colCampo.Add (Exclusivo.Name)
    objGridInt.colCampo.Add (DataUltimaCotacao.Name)
    objGridInt.colCampo.Add (UltimaCotacao.Name)
    objGridInt.colCampo.Add (TipoFrete.Name)
    objGridInt.colCampo.Add (DataUltimaCompra.Name)
    objGridInt.colCampo.Add (PrazoEntrega.Name)
    objGridInt.colCampo.Add (QuantPedidaForn.Name)
    objGridInt.colCampo.Add (QuantRecebidaForn.Name)
    objGridInt.colCampo.Add (CondicaoPagto.Name)
    objGridInt.colCampo.Add (SaldoTitulos.Name)
    objGridInt.colCampo.Add (ObservacaoForn.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_EscolhidoForn_Col = 1
    iGrid_ProdutoForn_Col = 2
    iGrid_DescProdutoForn_Col = 3
    iGrid_FornecedorGrid_Col = 4
    iGrid_FilialFornGrid_Col = 5
    iGrid_Exclusivo_Col = 6
    iGrid_UltimaCotacao_Col = 7
    iGrid_ValorCotacao_Col = 8
    iGrid_Frete_Col = 9
    iGrid_UltimaCompra_Col = 10
    iGrid_PrazoEntrega_Col = 11
    iGrid_QuantPedidaForn_Col = 12
    iGrid_QuantRecebidaForn_Col = 13
    iGrid_CondicaoPagto_Col = 14
    iGrid_SaldoTitulos_Col = 15
    iGrid_ObservacaoForn_Col = 16

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridFornecedores

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_FORNECEDORES_COTACAO + 1

    'Não permite incluir e excluir linhas do grid
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 20

    'Largura da primeira coluna
    GridFornecedores.ColWidth(0) = 400

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Fornecedores = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Fornecedores:

    Inicializa_Grid_Fornecedores = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161325)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Requisicoes(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Requisicoes

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Requisicoes

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Detalhe")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Filial Empresa")
    objGridInt.colColuna.Add ("P.V.")
    objGridInt.colColuna.Add ("Data Limite")
    objGridInt.colColuna.Add ("Data RC")
    objGridInt.colColuna.Add ("Urgente")
    objGridInt.colColuna.Add ("Requisitante")
    objGridInt.colColuna.Add ("Ccl")
    objGridInt.colColuna.Add ("Observação")

    'campos de edição do grid
    objGridInt.colCampo.Add (EscolhidoReq.Name)
    objGridInt.colCampo.Add (Requisicao.Name)
    objGridInt.colCampo.Add (FilialReq.Name)
    objGridInt.colCampo.Add (CodigoPV.Name)
    objGridInt.colCampo.Add (DataLimite.Name)
    objGridInt.colCampo.Add (DataReq.Name)
    objGridInt.colCampo.Add (Urgente.Name)
    objGridInt.colCampo.Add (Requisitante.Name)
    objGridInt.colCampo.Add (CclReq.Name)
    objGridInt.colCampo.Add (ObservacaoReq.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_EscolhidoReq_Col = 1
    iGrid_Numero_Col = 2
    iGrid_FilialReq_Col = 3
    iGrid_CodigoPV_Col = 4
    iGrid_DataLimite_Col = 5
    iGrid_DataRC_Col = 6
    iGrid_Urgente_Col = 7
    iGrid_Requisitante_Col = 8
    iGrid_CclReq_Col = 9
    iGrid_Observacao_Col = 10

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridRequisicoes

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_REQUISICOES + 1

    'Não permite incluir e excluir linhas do grid
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 20
    
    'Largura da primeira coluna
    GridRequisicoes.ColWidth(0) = 400

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161326)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_ItensRequisicoes(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid ItensRequisicoes

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_ItensRequisicoes

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Escolhido")
    objGridInt.colColuna.Add ("Requisição")
    objGridInt.colColuna.Add ("Filial Empresa")
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("UM")
    objGridInt.colColuna.Add ("A Cotar")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Quant. Pedida")
    objGridInt.colColuna.Add ("Quant. Recebida")
    objGridInt.colColuna.Add ("Quant. Cancelada")
    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("Ccl")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial Fornecedor")
    objGridInt.colColuna.Add ("Exclusividade")
    objGridInt.colColuna.Add ("Observação")

    'campos de edição do grid
    objGridInt.colCampo.Add (EscolhidoItem.Name)
    objGridInt.colCampo.Add (CodigoReq.Name)
    objGridInt.colCampo.Add (FilialEmpresaItemReq.Name)
    objGridInt.colCampo.Add (Item.Name)
    objGridInt.colCampo.Add (ProdutoItem.Name)
    objGridInt.colCampo.Add (DescProdutoItem.Name)
    objGridInt.colCampo.Add (UnidadeMedItem.Name)
    objGridInt.colCampo.Add (QuantCotarItem.Name)
    objGridInt.colCampo.Add (QuantidadeItem.Name)
    objGridInt.colCampo.Add (QuantPedidaItem.Name)
    objGridInt.colCampo.Add (QuantRecebidaItem.Name)
    objGridInt.colCampo.Add (QuantCanceladaItem.Name)
    objGridInt.colCampo.Add (Almoxarifado.Name)
    objGridInt.colCampo.Add (CclItem.Name)
    objGridInt.colCampo.Add (FornecedorItem.Name)
    objGridInt.colCampo.Add (FilialFornItem.Name)
    objGridInt.colCampo.Add (ExclusivoItem.Name)
    objGridInt.colCampo.Add (ObservacaoItem.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_EscolhidoItem_Col = 1
    iGrid_Requisicao_Col = 2
    iGrid_FilialEmpresaItemReq_Col = 3
    iGrid_Item_Col = 4
    iGrid_ProdutoItem_Col = 5
    iGrid_DescProdutoItem_Col = 6
    iGrid_UnidadeMedItem_Col = 7
    iGrid_QuantCotarItem_Col = 8
    iGrid_Quantidade_Col = 9
    iGrid_QuantPedida_Col = 10
    iGrid_QuantRecebida_Col = 11
    iGrid_QuantCancelada_Col = 12
    iGrid_Almoxarifado_Col = 13
    iGrid_CclItem_Col = 14
    iGrid_FornecedorItem_Col = 15
    iGrid_FilialFornItem_Col = 16
    iGrid_ExclusoItem_Col = 17
    iGrid_ObservacaoItem_Col = 18

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridItensRequisicoes

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_REQUISICAO + 1

    'Não permite incluir e excluir linhas do grid
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 20

    'Largura da primeira coluna
    GridItensRequisicoes.ColWidth(0) = 400

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
    
    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_ItensRequisicoes = SUCESSO

    Exit Function

Erro_Inicializa_Grid_ItensRequisicoes:

    Inicializa_Grid_ItensRequisicoes = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161327)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Produtos(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Produtos

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Produtos

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Escolhido")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M. de Compra")
    objGridInt.colColuna.Add ("A Cotar")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial Fornecedor")

    'campos de edição do grid
    objGridInt.colCampo.Add (EscolhidoProd.Name)
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescProduto.Name)
    objGridInt.colCampo.Add (UnidadeMedProd.Name)
    objGridInt.colCampo.Add (QuantidadeProd.Name)
    objGridInt.colCampo.Add (FornecedorProd.Name)
    objGridInt.colCampo.Add (FilialFornProd.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_EscolhidoProd_Col = 1
    iGrid_Produto_Col = 2
    iGrid_Descricao_Col = 3
    iGrid_UnidadeMed_Col = 4
    iGrid_QuantidadeProd_Col = 5
    iGrid_FornecedorProd_Col = 6
    iGrid_FilialFornProd_Col = 7

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridProdutos

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_PRODUTOS_COTACAO + 1

    'Não permite incluir e excluir linhas do grid
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 20
    
    'Largura da primeira coluna
    GridProdutos.ColWidth(0) = 400

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
    
    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Produtos = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Produtos:

    Inicializa_Grid_Produtos = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161328)

    End Select

    Exit Function

End Function

Private Sub OrdemRequisicao_Carrega()
'preenche a combo OrdemRequisicao e inicializa variaveis globais

Dim iIndice As Integer

    'Carregar os arrays de ordenação dos Bloqueios
    asOrdemRequisicao(0) = "RequisicaoCompra.DataLimite,RequisicaoCompra.Urgente"
    asOrdemRequisicao(1) = "RequisicaoCompra.Urgente,RequisicaoCompra.DataLimite"
    asOrdemRequisicao(2) = "RequisicaoCompra.Data,RequisicaoCompra.DataLimite"
    asOrdemRequisicao(3) = "RequisicaoCompra.FilialEmpresa,RequisicaoCompra.Codigo"
    asOrdemRequisicao(4) = "RequisicaoCompra.Ccl,RequisicaoCompra.DataLimite"


    asOrdemRequisicaoString(0) = "Data Limite"
    asOrdemRequisicaoString(1) = "Urgente"
    asOrdemRequisicaoString(2) = "Data"
    asOrdemRequisicaoString(3) = "Código"
    asOrdemRequisicaoString(4) = "Ccl"

    OrdemRequisicao.Clear

    'Carrega a Combobox OrdemRequisicao
    For iIndice = 0 To 4

        OrdemRequisicao.AddItem asOrdemRequisicaoString(iIndice)
        OrdemRequisicao.ItemData(OrdemRequisicao.NewIndex) = iIndice

    Next

    'Seleciona a opção Código de seleção
    OrdemRequisicao.ListIndex = 3

    Exit Sub

End Sub

Private Sub OrdemFornecedor_Carrega()
'preenche a combo OrdemFornecedor e inicializa variaveis globais

Dim iIndice As Integer

    'Carregar os arrays de ordenação dos Bloqueios
    asOrdemFornecedor(0) = "FornecedorProdutoFF.Produto AND FornecedorProdutoFF.Fornecedor"
    asOrdemFornecedor(1) = "FornecedorProdutoFF.Fornecedor AND FornecedorProdutoFF.Produto"

    asOrdemFornecedorString(0) = "Produto"
    asOrdemFornecedorString(1) = "Fornecedor"

    'Carrega a Combobox OrdemFornecedor
    For iIndice = 0 To 1

        OrdemFornecedor.AddItem asOrdemFornecedorString(iIndice)
        OrdemFornecedor.ItemData(OrdemFornecedor.NewIndex) = iIndice

    Next

    'Seleciona a opção Produto de seleção
    OrdemFornecedor.ListIndex = 0

    Exit Sub

End Sub

'Private Function Carrega_CondicaoPagamento() As Long
''Carrega Listbox CondPagtos com as condicoes de pagamento usadas em contas a pagar (EmPagamento=1)
'
'Dim lErro As Long
'Dim colCod_DescReduzida As New AdmColCodigoNome
'Dim objCod_DescReduzida As New AdmCodigoNome
'
'On Error GoTo Erro_Carrega_CondicaoPagamento
'
'    'Le todos os Codigos e DescReduzida da tabela CondicoesPagto com a condicao EmPagamento = 1 e coloca na colecao colCod_DescReduzida
'    lErro = CF("CondicoesPagto_Le_Pagamento", colCod_DescReduzida)
'    If lErro <> SUCESSO Then gError 63416
'
'    For Each objCod_DescReduzida In colCod_DescReduzida
'
'        'Verifica se a CondPagto é diferente de "À Vista"
'        If objCod_DescReduzida.iCodigo <> CONDPAGTO_VISTA Then
'
'            'Adiciona novo item na ListBox CondPagtos
'            CondPagtos.AddItem CInt(objCod_DescReduzida.iCodigo) & SEPARADOR & objCod_DescReduzida.sNome
'            CondPagtos.ItemData(CondPagtos.NewIndex) = objCod_DescReduzida.iCodigo
'
'        End If
'
'    Next
'
'    Carrega_CondicaoPagamento = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_CondicaoPagamento:
'
'    Carrega_CondicaoPagamento = gErr
'
'    Select Case gErr
'
'        Case 63416
'            'Erro tratado na rotina chamada
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161329)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Function Carrega_TipoProduto() As Long
'Carrega a ListBox TipoProduto com tipos de produtos que possam ser comprados (Compras=1)

Dim lErro As Long
Dim colCod_DescReduzida As New AdmColCodigoNome
Dim objCod_DescReduzida As New AdmCodigoNome
Dim iIndice As Integer

On Error GoTo Erro_Carrega_TipoProduto

    'Le todos os Codigos e DescReduzida de tipos de produtos
    lErro = CF("TiposProduto_Le_Todos", colCod_DescReduzida)
    If lErro <> SUCESSO Then gError 63413

    For Each objCod_DescReduzida In colCod_DescReduzida

        'Adiciona novo item na ListBox CondPagtos
        TipoProduto.AddItem CStr(objCod_DescReduzida.iCodigo) & SEPARADOR & objCod_DescReduzida.sNome
        TipoProduto.ItemData(TipoProduto.NewIndex) = objCod_DescReduzida.iCodigo

    Next

    'Marca todos os TipoProduto
    For iIndice = 0 To TipoProduto.ListCount - 1
        TipoProduto.Selected(iIndice) = True
    Next

    Carrega_TipoProduto = SUCESSO

    Exit Function

Erro_Carrega_TipoProduto:

    Carrega_TipoProduto = gErr

    Select Case gErr

        Case 63413
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161330)

    End Select

    Exit Function

End Function

Private Sub Almoxarifado_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Almoxarifado_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensRequisicoes)

End Sub

Private Sub Almoxarifado_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensRequisicoes)

End Sub

Private Sub Almoxarifado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensRequisicoes.objControle = Almoxarifado
    lErro = Grid_Campo_Libera_Foco(objGridItensRequisicoes)
    If lErro <> SUCESSO Then Cancel = True


End Sub
Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub Cotacao_Change()
    giAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Cotacao_GotFocus()
    Call MaskEdBox_TrataGotFocus(Cotacao, giAlterado)
End Sub

Private Sub FilialEmpresa_Click()
    giTabSelecao_Alterado = REGISTRO_ALTERADO
End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long
Dim iFrameAnterior As Integer

'teste de tempo
'Dim dtIni As Date
'dtIni = Now

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index = iFrameAtual Then Exit Sub

    If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub
        
    iFrameAnterior = iFrameAtual
    'Torna Frame correspondente ao Tab selecionado visivel
    Frame1(TabStrip1.SelectedItem.Index).Visible = True
    'Torna Frame atual invisivel
    Frame1(iFrameAtual).Visible = False
    'Armazena novo valor de iFrameAtual
    iFrameAtual = TabStrip1.SelectedItem.Index
    
    If giTabSelecao_Alterado = REGISTRO_ALTERADO And iFrameAtual <> TAB_Selecao Then
    
        'Limpa o GridRequisicoes
        Call Grid_Limpa(objGridRequisicoes) '0s

        'Limpa o GridItensRequisicoes
        Call Grid_Limpa(objGridItensRequisicoes) '0s
    
        'Limpa o GridProdutos
        Call Grid_Limpa(objGridProdutos) '0s
    
        'Limpa o GridFornecedores
        Call Grid_Limpa(objGridFornecedores) '0s
    
        'Lê Requisicoes p/ objGeracaoCotacao conforme filtros tab Selecao
        lErro = Geracao_Le_Requisicoes(gobjGeracaoCotacao) '1s
        If lErro <> SUCESSO Then gError 25969
    
        'Coloca quantidades a cotar default nos ítens requisicoes
        lErro = ItensReq_Calcula_QuantCotar(gobjGeracaoCotacao) '0s
        If lErro <> SUCESSO Then gError 77048
    
        'Preenche o Grid de Requisicoes
        lErro = GridRequisicao_Preenche(gobjGeracaoCotacao) '1s
        If lErro <> SUCESSO Then gError 63542
    
        'Preenche o GridItensRequisicoes
        lErro = GridItens_Preenche(gobjGeracaoCotacao) '2s
        If lErro <> SUCESSO Then gError 25970

        'Move os Itens de Requisicao para gcolCotacaoProduto
        lErro = Move_Itens_Produtos(gobjGeracaoCotacao, gcolCotacaoProduto) '1s
        If lErro <> SUCESSO Then gError 77019

        'Ordena gcolCotacaoProduto por ordem de Produto, Fornecedor, Filial
        lErro = Ordena_Produtos(gcolCotacaoProduto) '0s
        If lErro <> SUCESSO Then gError 77000

        'Preenche o Grid de Produtos
        lErro = GridProduto_Preenche(gcolCotacaoProduto) '1s
        If lErro <> SUCESSO Then gError 63494

        'Preenche dados dos Fornecedores dos Produtos a partir dos dados lidos da tabela FornecedorProdutoFF
        lErro = CF("CotacaoProdutoFornecedor_Le", gcolCotacaoProduto, gcolFornecedorProdutoFF) '3s
        If lErro <> SUCESSO Then gError 63585

        'Preenche coleção de Itens de GridFornecedores
        lErro = ColItemGridFornecedores_Preenche(gcolCotacaoProduto, gcolFornecedorProdutoFF, gcolItemGridFornecedores) '4s
        If lErro <> SUCESSO Then gError 77001

        'Ordena coleção de Itens de GridFornecedores
        lErro = Ordena_Fornecedor(gcolItemGridFornecedores) '0s
        If lErro <> SUCESSO Then gError 77002

        'Devolve os elementos ordenados para o GridFornecedores
        lErro = GridFornecedores_Devolve(gcolItemGridFornecedores) '0s
        If lErro <> SUCESSO Then gError 63491
        
        giTabSelecao_Alterado = 0
    
    'Se o frame de Requisiçoes foi alterado,
    ElseIf giTabRequisicao_Alterado = REGISTRO_ALTERADO Then
                    
        'Preenche o GridItensRequisicoes
        lErro = GridItens_Preenche(gobjGeracaoCotacao)
        If lErro <> SUCESSO Then gError 77003
    
        'Preenche o Grid de Produtos
        lErro = GridProduto_Preenche(gcolCotacaoProduto)
        If lErro <> SUCESSO Then gError 77005
    
        'Devolve os elementos ordenados para o GridFornecedores
        lErro = GridFornecedores_Devolve(gcolItemGridFornecedores)
        If lErro <> SUCESSO Then gError 77009
        
        giTabRequisicao_Alterado = 0
    
    ElseIf giTabItens_Alterado = REGISTRO_ALTERADO Then
    
        'Preenche o Grid de Produtos
        lErro = GridProduto_Preenche(gcolCotacaoProduto)
        If lErro <> SUCESSO Then gError 77011
            
        'Devolve os elementos ordenados para o GridFornecedores
        lErro = GridFornecedores_Devolve(gcolItemGridFornecedores)
        If lErro <> SUCESSO Then gError 77015
        
        giTabItens_Alterado = 0
    
    ElseIf giTabProdutos_Alterado = REGISTRO_ALTERADO Then
    
        'Devolve os elementos ordenados para o GridFornecedores
        lErro = GridFornecedores_Devolve(gcolItemGridFornecedores)
        If lErro <> SUCESSO Then gError 77018
        
        giTabProdutos_Alterado = 0
    
    End If
    
    'teste de tempo
    'MsgBox "Final:" & CStr(Format(Now, "nn:ss")) & "   Inicial:" & CStr(Format(dtIni, "nn:ss"))
    
    Exit Sub
    
Erro_TabStrip1_Click:

    Select Case gErr
    
        Case 25969, 25970, 63491, 63494, 63542, 63585, 77000 To 77003, 77005, 77009, 77011, 77015, 77018, 77048
            'Erros tratados nas rotinas chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161331)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoFornecedor_Click()
'Chama a tela de FilialFornecedor de acordo com a linha do GridFornecedores selecionada

Dim lErro As Long
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_BotaoFornecedor_Click

    'Verifica se a linha do GridFornecedores é uma linha existente
    If GridFornecedores.Row > 0 And GridFornecedores.Row <= objGridFornecedores.iLinhasExistentes Then

        'Coloca Codigo do Fornecedor e da Filial em objFilialFornecedor
        objFilialFornecedor.iCodFilial = Codigo_Extrai(GridFornecedores.TextMatrix(GridFornecedores.Row, iGrid_FilialFornGrid_Col))
        
        objFornecedor.sNomeReduzido = GridFornecedores.TextMatrix(GridFornecedores.Row, iGrid_FornecedorGrid_Col)
        
        If Len(Trim(objFornecedor.sNomeReduzido)) > 0 Then
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then gError 68344
        
            If lErro = 6681 Then gError 70515
        End If
        
        objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo

        'Chama a tela FilialFornecedor
        Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)

    'Se a linha do GridFornecedores não for uma linha existente==>erro
    ElseIf GridFornecedores.Row > objGridFornecedores.iLinhasExistentes Or GridFornecedores.Row = 0 Then gError 63435

    End If

    Exit Sub

Erro_BotaoFornecedor_Click:

    Select Case gErr

        Case 68344
            'Erro tratado na rotina chamada
            
        Case 63435
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_FORN_LINHA_NAO_SELECIONADA", gErr)
        
        Case 70515
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161332)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGeraPedidos_Click()
'Gera Pedidos de Cotação

Dim lErro As Long

On Error GoTo Erro_BotaoGeraPedidos_Click

    'Gera Pedido de Cotacao
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 63503

    Call Limpa_Tela_GeracaoCotacao
    
    giAlterado = 0
    
    Exit Sub

Erro_BotaoGeraPedidos_Click:

    Select Case gErr

        Case 63503
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161333)

    End Select

    Exit Sub

End Sub

Private Sub BotaoImprimePedidos_Click()
'Imprime os Pedidos de Cotacao

Dim lErro As Long
Dim objCotacao As New ClassCotacao
Dim colPedidoCotacao As New Collection
Dim objPedidoCotacao As ClassPedidoCotacao

On Error GoTo Erro_BotaoImprimePedidos_Click

    Set gobjCotacao = Nothing

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 63606
    
    If gobjCotacao Is Nothing Then gError 76077
    
   'Imprime os Pedidos de Cotacao
    lErro = PedidosCotacao_Imprime(gobjCotacao)
    If lErro <> SUCESSO Then gError 63149

    'Para cada Pedido de Cotação da coleção de pedidos
    For Each objPedidoCotacao In gcolPedidoCotacao

        'Atualiza data de emissao no BD para a data atual
        lErro = CF("PedidoCotacao_Atualiza_DataEmissao", objPedidoCotacao)
        If lErro <> SUCESSO And lErro <> 56348 Then gError 89861

    Next

    Call Limpa_Tela_GeracaoCotacao
    
    giAlterado = 0
    
    Exit Sub

Erro_BotaoImprimePedidos_Click:

    Select Case gErr

        Case 63148
            Call Rotina_Erro(vbOKOnly, "ERRO_PED_COTACAO_NAO_GERADO", gErr)

        Case 63149, 63606, 89861

        Case 76077
            Call Rotina_Erro(vbOKOnly, "ERRO_PED_COTACAO_NAO_GERADO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161334)

    End Select

    Exit Sub

End Sub

Function PedidosCotacao_Imprime(objCotacao As ClassCotacao) As Long
'Chama a impressao de pedidos de cotacao

Dim objRelatorio As New AdmRelatorio
Dim lErro As Long

On Error GoTo Erro_PedidosCotacao_Imprime

    lErro = objRelatorio.ExecutarDireto("Geracao de Pedido de Cotacao", "COTACAO.NumIntDoc=@NCOTACAO", 1, "COTACAO", "NCOTACAO", objCotacao.lNumIntDoc)
    If lErro <> SUCESSO Then gError 63200

    PedidosCotacao_Imprime = SUCESSO

    Exit Function

Erro_PedidosCotacao_Imprime:

    PedidosCotacao_Imprime = gErr

    Select Case gErr

        Case 63200
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161335)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()
'Limpa a tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, giAlterado)
    If lErro <> SUCESSO Then gError 63468

    'Limpa a tela
    Call Limpa_Tela_GeracaoCotacao

    giAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 63468
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161336)

    End Select

    Exit Sub

End Sub

Sub Limpa_Tela_GeracaoCotacao()
'Limpa os campos da tela

    Call Limpa_Tela(Me)

    'Marca todos os Tipos de Produto da list TipoProduto
    Call MarcaTodos_TipoProduto

    'Limpa o GridRequisicoes
    Call Grid_Limpa(objGridRequisicoes)

    'Limpa o GridItensRequisicoes
    Call Grid_Limpa(objGridItensRequisicoes)

    'Limpa o GridProdutos
    Call Grid_Limpa(objGridProdutos)

    'Limpa o GridFornecedores
    Call Grid_Limpa(objGridFornecedores)

    'Limpa os demais campos do TabSelecao
    SelecionaDestino.Value = vbChecked
    TipoDestino(TIPO_DESTINO_EMPRESA).Value = True
    FilialFornec.Clear
    ExibeCotadas.Value = vbUnchecked
    CondPagto.Caption = ""
    Cotacao.Caption = ""

End Sub

Private Sub BotaoMarcarTodos_Click(Index As Integer)
'Marca todos os itens de acordo com o índice determinado

    'De acordo com o índice do tab visivel, chama a função específica
    Select Case Index

        Case TAB_Selecao

            'Marca todos os Tipos de Produto da list TipoProduto
            Call MarcaTodos_TipoProduto

            giTabSelecao_Alterado = REGISTRO_ALTERADO

        Case TAB_REQUISICAO

            'Marca todas as Requisicoes do GridRequisicoes
            Call MarcaTodas_Requisicoes

            giTabRequisicao_Alterado = REGISTRO_ALTERADO

        Case TAB_ITENSREQ

            'Marca todos os itens do GridItensRequisicoes
            Call MarcaTodos_ItensRequisicoes

            giTabItens_Alterado = REGISTRO_ALTERADO

        Case TAB_Produtos

            'Marca todos os Produtos do GridProdutos
            Call MarcaTodos_Produtos

            giTabProdutos_Alterado = REGISTRO_ALTERADO

        Case TAB_FORNECEDOR

            'Marca todos os Fornecedores do GridFornecedores
            Call MarcaTodos_Fornecedores

            giTabFornecedor_Alterado = REGISTRO_ALTERADO

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoProduto_Click()
'Chama a tela Produtos

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objProduto As New ClassProduto

On Error GoTo Erro_BotaoProduto_Click

    'Verifica se existe alguma linha do GridProdutos selecionada
    If GridProdutos.Row = 0 Then gError 63600

    'Verifica se o produto da linha selecionada está preenchido
    If Len(Trim(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Produto_Col))) > 0 Then

        'Passa o codigo do produto para o formato do BD
        lErro = CF("Produto_Formata", GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 63434

        'Coloca o código formatado no objProduto
        objProduto.sCodigo = sProdutoFormatado

    End If

    'Chama a tela Produto
    Call Chama_Tela("Produto", objProduto)

    Exit Sub

Erro_BotaoProduto_Click:

    Select Case gErr

        Case 63434
            'Erro tratado na rotina chamada

        Case 63600
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_PRODUTOS_NAO_SELECIONADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161337)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCotacao As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Obtém o próximo código de Cotacao disponível
    lErro = CF("Cotacao_Automatica", lCotacao)
    If lErro <> SUCESSO Then gError 63421

    'Coloca o Código de Cotacao obtido na tela
    Cotacao.Caption = lCotacao

    Exit Sub
    
Erro_BotaoProxNum_Click:

    Select Case gErr
    
        Case 63421
        
        Case Else
           Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161338)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoRequisicao_Click()
'Chama a tela Requisicoes, de acordo com a RequisicaoCompra selecionada no grid

Dim lErro As Long
Dim objRequisicaoCompras As New ClassRequisicaoCompras

On Error GoTo Erro_BotaoRequisicao_Click

    'Verifica se existe alguma linha do GridRequisicoes selecionada
    If GridRequisicoes.Row = 0 Then gError 63599

    'Verifica se o número da requisicao da linha selecionada está preenchido
    If GridRequisicoes.TextMatrix(GridRequisicoes.Row, iGrid_Numero_Col) > 0 Then

        'Coloca o código da Requisicao de Compra no objRequisicaoCompras
        objRequisicaoCompras.lCodigo = StrParaLong(GridRequisicoes.TextMatrix(GridRequisicoes.Row, iGrid_Numero_Col))
        objRequisicaoCompras.iFilialEmpresa = Codigo_Extrai(GridRequisicoes.TextMatrix(GridRequisicoes.Row, iGrid_FilialReq_Col))
        
    End If

    'Chama a tela ReqCompras
    Call Chama_Tela("ReqComprasCons", objRequisicaoCompras)

    Exit Sub

Erro_BotaoRequisicao_Click:

    Select Case gErr

        Case 63599
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_REQUISICAO_NAO_SELECIONADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161339)

    End Select

    Exit Sub

End Sub

Private Sub CclItem_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CclItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensRequisicoes)

End Sub

Private Sub CclItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensRequisicoes)

End Sub

Private Sub CclItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensRequisicoes.objControle = CclItem
    lErro = Grid_Campo_Libera_Foco(objGridItensRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CclReq_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRequisicoes)

End Sub

Private Sub CclReq_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRequisicoes)

End Sub

Private Sub CclReq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRequisicoes.objControle = CclReq
    lErro = Grid_Campo_Libera_Foco(objGridRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CodigoAte_Change()

    giTabSelecao_Alterado = REGISTRO_ALTERADO

End Sub

Private Sub CodigoAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoAte, giAlterado)
    
End Sub

Private Sub CodigoDe_Change()

    giTabSelecao_Alterado = REGISTRO_ALTERADO

End Sub

Private Sub CodigoDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoDe, giAlterado)
    
End Sub

Private Sub CodigoReq_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodigoReq_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensRequisicoes)

End Sub

Private Sub CodigoReq_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensRequisicoes)

End Sub

Private Sub CodigoReq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensRequisicoes.objControle = CodigoReq
    lErro = Grid_Campo_Libera_Foco(objGridItensRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub


Private Sub CondicaoPagto_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CondicaoPagto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub CondicaoPagto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub CondicaoPagto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = CondicaoPagto
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CondPagtos_DblClick()

    'Preenche CondPagto com o texto de CondPagtos
    If gobjCRFAT.iCondPagtoSemCodigo = 0 Then
        CondPagto.Caption = CondPagtos.Text
    Else
        CondPagto.Caption = CStr(CondPagto_Extrai(CondPagtos)) & SEPARADOR & CondPagtos.Text
    End If

End Sub

Private Sub DataAte_Change()

    giAlterado = REGISTRO_ALTERADO
    giTabSelecao_Alterado = REGISTRO_ALTERADO

End Sub

Private Sub DataAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataAte, giAlterado)
    
End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'Verifica se  DataAte foi preenchida
    If Len(Trim(DataAte.Text)) = 0 Then Exit Sub

    'Critica DataAte
    lErro = Data_Critica(DataAte.Text)
    If lErro <> SUCESSO Then gError 63439

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 63439
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161340)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_Change()

    giAlterado = REGISTRO_ALTERADO
    giTabSelecao_Alterado = REGISTRO_ALTERADO

End Sub

Private Sub DataDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataDe, giAlterado)
    
End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'Verifica se  DataDe foi preenchida
    If Len(Trim(DataDe.Text)) = 0 Then Exit Sub

    'Critica DataDe
    lErro = Data_Critica(DataDe.Text)
    If lErro <> SUCESSO Then gError 63438

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 63438
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161341)

    End Select

    Exit Sub

End Sub

Private Sub DataLimite_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRequisicoes)

End Sub

Private Sub DataLimite_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRequisicoes)

End Sub

Private Sub DataLimite_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRequisicoes.objControle = DataLimite
    lErro = Grid_Campo_Libera_Foco(objGridRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DataLimiteAte_Change()

    giAlterado = REGISTRO_ALTERADO
    giTabSelecao_Alterado = REGISTRO_ALTERADO

End Sub

Private Sub DataLimiteAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataLimiteAte, giAlterado)
    
End Sub

Private Sub DataLimiteAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataLimiteAte_Validate

    'Verifica se  DataLimiteAte foi preenchida
    If Len(Trim(DataLimiteAte.Text)) = 0 Then Exit Sub

    'Critica DataLimiteAte
    lErro = Data_Critica(DataLimiteAte.Text)
    If lErro <> SUCESSO Then gError 63437

    Exit Sub

Erro_DataLimiteAte_Validate:

    Cancel = True

    Select Case gErr

        Case 63437
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161342)

    End Select

    Exit Sub

End Sub

Private Sub DataLimiteDe_Change()

    giAlterado = REGISTRO_ALTERADO
    giTabSelecao_Alterado = REGISTRO_ALTERADO

End Sub

Private Sub DataLimiteDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataLimiteDe, giAlterado)
    
End Sub

Private Sub DataLimiteDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataLimiteDe_Validate

    'Verifica se  DataLimiteDe foi preenchida
    If Len(Trim(DataLimiteDe.Text)) = 0 Then Exit Sub

    'Critica DataLimiteDe
    lErro = Data_Critica(DataLimiteDe.Text)
    If lErro <> SUCESSO Then gError 63436

    Exit Sub

Erro_DataLimiteDe_Validate:

    Cancel = True

    Select Case gErr

        Case 63436
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161343)

    End Select

    Exit Sub

End Sub

Private Sub DataReq_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRequisicoes)

End Sub

Private Sub DataReq_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRequisicoes)

End Sub

Private Sub DataReq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRequisicoes.objControle = DataReq
    lErro = Grid_Campo_Libera_Foco(objGridRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DataUltimaCompra_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataUltimaCompra_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub DataUltimaCompra_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub DataUltimaCompra_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = DataUltimaCompra
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DataUltimaCotacao_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataUltimaCotacao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub DataUltimaCotacao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub DataUltimaCotacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = DataUltimaCotacao
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DescProduto_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub DescProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub DescProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = DescProduto
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DescProdutoForn_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescProdutoForn_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub DescProdutoForn_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub DescProdutoForn_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = DescProdutoForn
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DescProdutoItem_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescProdutoItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensRequisicoes)

End Sub

Private Sub DescProdutoItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensRequisicoes)

End Sub

Private Sub DescProdutoItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensRequisicoes.objControle = DescProdutoItem
    lErro = Grid_Campo_Libera_Foco(objGridItensRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub EscolhidoForn_Click()

Dim lErro As Long
Dim iLinha As Integer
Dim objFornecedorProdutoFF As ClassFornecedorProdutoFF
Dim objItemGridFornecedor As ClassItemGridFornecedores

On Error GoTo Erro_EscolhidoForn_Click

    giTabFornecedor_Alterado = REGISTRO_ALTERADO
    iLinha = GridFornecedores.Row
        
    'Faz objItemGridFornecedor e objFornecedorProdutoFF apontarem para elementos correspondentes na colecao global
    lErro = ItemGridFornecedor_Escolhe(iLinha, objItemGridFornecedor, objFornecedorProdutoFF)
    If lErro <> SUCESSO Then gError 77074
        
    If StrParaInt(GridFornecedores.TextMatrix(GridFornecedores.Row, iGrid_EscolhidoForn_Col)) = vbChecked And objItemGridFornecedor.sEscolhido = CStr(NAO_SELECIONADO) Then
    
        objItemGridFornecedor.sEscolhido = CStr(Selecionado)
        objFornecedorProdutoFF.iEscolhido = Selecionado
        
    ElseIf StrParaInt(GridFornecedores.TextMatrix(GridFornecedores.Row, iGrid_EscolhidoForn_Col)) = vbUnchecked And objItemGridFornecedor.sEscolhido = CStr(Selecionado) Then
        
        objItemGridFornecedor.sEscolhido = CStr(NAO_SELECIONADO)
        objFornecedorProdutoFF.iEscolhido = NAO_SELECIONADO
        
    End If
    
    Exit Sub
    
Erro_EscolhidoForn_Click:

    Select Case gErr
    
        Case 77074
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161344)
            
    End Select
    
    Exit Sub

End Sub

Private Sub EscolhidoForn_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub EscolhidoForn_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub EscolhidoForn_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = EscolhidoForn
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub EscolhidoItem_Click()

Dim objItemReqCompras As ClassItemReqCompras
Dim iLinha As Integer
Dim lErro As Long

On Error GoTo Erro_EscolhidoItem_Click

    iLinha = GridItensRequisicoes.Row
   
    'Faz objItemReqCompras apontar para elemento correspondente na coleção global
    lErro = ItemReqCompras_Escolhe(iLinha, objItemReqCompras)
    If lErro <> SUCESSO Then gError 77044

    If StrParaInt(GridItensRequisicoes.TextMatrix(GridItensRequisicoes.Row, iGrid_EscolhidoItem_Col)) = vbChecked And objItemReqCompras.iSelecionado = NAO_SELECIONADO Then
    
        objItemReqCompras.iSelecionado = MARCADO
        
        'Atualiza colecoes globais
        lErro = Atualiza_Selecao_ItemReqCompras(objItemReqCompras, gcolCotacaoProduto, gcolFornecedorProdutoFF, gcolItemGridFornecedores)
        If lErro <> SUCESSO Then gError 77037
        
        giTabItens_Alterado = REGISTRO_ALTERADO

    ElseIf StrParaInt(GridItensRequisicoes.TextMatrix(GridItensRequisicoes.Row, iGrid_EscolhidoItem_Col)) = vbUnchecked And objItemReqCompras.iSelecionado = Selecionado Then
    
        objItemReqCompras.iSelecionado = DESMARCADO
        
        'Atualiza colecoes globais
        lErro = Atualiza_Selecao_ItemReqCompras(objItemReqCompras, gcolCotacaoProduto, gcolFornecedorProdutoFF, gcolItemGridFornecedores)
        If lErro <> SUCESSO Then gError 77038
    
        giTabItens_Alterado = REGISTRO_ALTERADO
    
    End If
    
    Exit Sub

Erro_EscolhidoItem_Click:

    Select Case gErr
    
        Case 77037, 77038, 77044
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161345)
        
    End Select
            
    Exit Sub
    
End Sub

Private Sub EscolhidoItem_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItensRequisicoes)
End Sub

Private Sub EscolhidoItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensRequisicoes)

End Sub

Private Sub EscolhidoItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensRequisicoes.objControle = EscolhidoItem
    lErro = Grid_Campo_Libera_Foco(objGridItensRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub EscolhidoReq_Change()
    giAlterado = REGISTRO_ALTERADO
End Sub

Private Sub EscolhidoProd_Click()

Dim objCotacaoProduto As ClassCotacaoProduto
Dim iLinha As Integer
Dim lErro As Long

On Error GoTo Erro_EscolhidoProd_Click
    
    iLinha = GridProdutos.Row
   
    'Faz objCotacaoProduto apontar para elemento correspondente na coleção global
    lErro = CotacaoProduto_Escolhe(iLinha, objCotacaoProduto)
    If lErro <> SUCESSO Then gError 77059

    If StrParaInt(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_EscolhidoProd_Col)) = vbChecked And objCotacaoProduto.iEscolhido = NAO_SELECIONADO Then
    
        objCotacaoProduto.iEscolhido = Selecionado
        
        'Atualiza colecoes globais
        lErro = Atualiza_Selecao_CotacaoProduto(gcolCotacaoProduto, gcolFornecedorProdutoFF, gcolItemGridFornecedores)
        If lErro <> SUCESSO Then gError 77060
        
        giTabProdutos_Alterado = REGISTRO_ALTERADO

    ElseIf StrParaInt(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_EscolhidoProd_Col)) = vbUnchecked And objCotacaoProduto.iEscolhido = Selecionado Then
    
        objCotacaoProduto.iEscolhido = NAO_SELECIONADO
        
        'Atualiza colecoes globais
        lErro = Atualiza_Selecao_CotacaoProduto(gcolCotacaoProduto, gcolFornecedorProdutoFF, gcolItemGridFornecedores)
        If lErro <> SUCESSO Then gError 77061
    
        giTabProdutos_Alterado = REGISTRO_ALTERADO
    
    End If
    
    Exit Sub
    
Erro_EscolhidoProd_Click:

    Select Case gErr
    
        Case 77059, 77060, 77061
        
        Case Else
           Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161346)
    
    End Select

    Exit Sub

End Sub

Private Sub EscolhidoReq_Click()

Dim lErro As Long
Dim objReqCompras As ClassRequisicaoCompras
Dim lCodReqCompras As Long
Dim iFilialEmpresa As Integer
Dim iLinha As Integer

On Error GoTo Erro_Saida_Celula_EscolhidoReq
    
    'Extrai o código da Req Compras da linha do grid
    iLinha = GridRequisicoes.Row
    lCodReqCompras = StrParaLong(GridRequisicoes.TextMatrix(iLinha, iGrid_Numero_Col))
    iFilialEmpresa = Codigo_Extrai(GridRequisicoes.TextMatrix(iLinha, iGrid_FilialReq_Col))
    
    'Seleciona na colecao o objReqCompras correspondente a linha do Grid
    For Each objReqCompras In gobjGeracaoCotacao.colReqCompra
        If (objReqCompras.lCodigo = lCodReqCompras) And (objReqCompras.iFilialEmpresa = iFilialEmpresa) Then Exit For
    Next
    
    If StrParaInt(GridRequisicoes.TextMatrix(iLinha, iGrid_EscolhidoReq_Col)) = vbChecked And objReqCompras.iSelecionado = NAO_SELECIONADO Then
        
        objReqCompras.iSelecionado = Selecionado
        
        lErro = Atualiza_SelecaoReqCompra(objReqCompras, gcolCotacaoProduto, gcolFornecedorProdutoFF, gcolItemGridFornecedores)
        If lErro <> SUCESSO Then gError 77023
        
        giTabRequisicao_Alterado = REGISTRO_ALTERADO
        
    ElseIf StrParaInt(GridRequisicoes.TextMatrix(iLinha, iGrid_EscolhidoReq_Col)) = vbUnchecked And objReqCompras.iSelecionado = Selecionado Then
    
        objReqCompras.iSelecionado = NAO_SELECIONADO
    
        lErro = Atualiza_SelecaoReqCompra(objReqCompras, gcolCotacaoProduto, gcolFornecedorProdutoFF, gcolItemGridFornecedores)
        If lErro <> SUCESSO Then gError 77024
    
        giTabRequisicao_Alterado = REGISTRO_ALTERADO
        
    End If
    
    Exit Sub

Erro_Saida_Celula_EscolhidoReq:

    Select Case gErr
        
        Case 77023 To 77024 'Tratados na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161347)

    End Select

    Exit Sub

End Sub

Private Sub EscolhidoReq_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRequisicoes)

End Sub

Private Sub EscolhidoReq_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRequisicoes)

End Sub

Private Sub EscolhidoReq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRequisicoes.objControle = EscolhidoReq
    lErro = Grid_Campo_Libera_Foco(objGridRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub
Private Sub EscolhidoProd_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub EscolhidoProd_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub EscolhidoProd_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = EscolhidoProd
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Exclusivo_Click()
    giAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Exclusivo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridFornecedores)
End Sub

Private Sub Exclusivo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub Exclusivo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = Exclusivo
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ExclusivoItem_Change()
    giAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ExclusivoItem_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItensRequisicoes)
End Sub

Private Sub ExclusivoItem_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensRequisicoes)
End Sub

Private Sub ExclusivoItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensRequisicoes.objControle = ExclusivoItem
    lErro = Grid_Campo_Libera_Foco(objGridItensRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ExibeCotadas_Click()

    giTabSelecao_Alterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialEmpresa_Change()

    giTabSelecao_Alterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialEmpresa_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_FilialEmpresa_Validate

    'Se a FilialEmpresa tiver sido selecionada ==> sai da rotina
    If FilialEmpresa.ListIndex <> -1 Then Exit Sub

        'Tenta selecionar a FilialEmpresa na combo FilialEmpresa
        lErro = Combo_Seleciona(FilialEmpresa, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 63464

        'Se nao encontra o ítem com o código informado
        If lErro = 6730 Then

            'preeenche objFilialEmpresa
            objFilialEmpresa.iCodFilial = iCodigo

            'Le a FilialEmpresa
            lErro = CF("FilialEmpresa_Le", objFilialEmpresa, True)
            If lErro <> SUCESSO And lErro <> 27378 Then gError 63465

            'Se nao encontrou => erro
            If lErro = 27378 Then gError 63466

            If lErro = SUCESSO Then

                'Coloca na tela o codigo e o nome da FilialEmpresa
                FilialEmpresa.Text = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome

            End If

        End If

        'Se nao encontrou e nao era codigo
        If lErro = 6731 Then gError 63467

    Exit Sub

Erro_FilialEmpresa_Validate:

    Cancel = True

    Select Case gErr

        Case 63464, 63465
            'Erros tratados nas rotinas chamadas

        Case 63466
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, iCodigo)

        Case 63467
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA1", gErr, FilialEmpresa.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161348)

    End Select

    Exit Sub

End Sub

Private Sub FilialFornec_Change()
    giTabSelecao_Alterado = REGISTRO_ALTERADO
End Sub

Private Sub FilialFornec_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_FilialFornec_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(FilialFornec.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If FilialFornec.ListIndex >= 0 Then Exit Sub

    'Tenta selecionar na combo de FilialFornec
    lErro = Combo_Seleciona(FilialFornec, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 63442

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        'Verifica de o fornecedor foi digitado
        If Len(Trim(Fornecedor.Text)) = 0 Then gError 63443

        sFornecedor = Fornecedor.Text

        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError 63444

        'Se nao existir
        If lErro = 18272 Then

            objFornecedor.sNomeReduzido = sFornecedor

            'Le o Código do Fornecedor --> Para Passar para a Tela de Filiais
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then gError 63445

            'Passa o Código do Fornecedor
            objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo

            'Sugere cadastrar nova Filial
            gError 63140

        End If

        'Coloca na tela o código e o nome da FilialForn
        FilialFornec.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 63446

    Exit Sub

Erro_FilialFornec_Validate:

    Cancel = True

    Select Case gErr

        Case 63442, 63444, 63445 'Tratados nas Rotinas chamadas

        Case 63140
            'Pergunta se deseja criar nova filial para o fornecedor em questao
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

            If vbMsgRes = vbYes Then
                'Chama a tela FiliaisFornecedores
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 63443
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 63446
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", gErr, FilialFornec.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161349)

    End Select

    Exit Sub

End Sub

Private Sub FilialFornGrid_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialFornGrid_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub FilialFornGrid_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub FilialFornGrid_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = FilialFornGrid
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub FilialFornItem_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub
Private Sub FilialFornItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensRequisicoes)

End Sub

Private Sub FilialFornItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensRequisicoes)

End Sub

Private Sub FilialFornItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensRequisicoes.objControle = FilialFornItem
    lErro = Grid_Campo_Libera_Foco(objGridItensRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub FilialFornProd_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub
Private Sub FilialFornProd_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub FilialFornProd_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub FilialFornProd_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = FilialFornProd
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub FilialReq_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridRequisicoes)
End Sub

Private Sub FilialReqReq_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRequisicoes)

End Sub

Private Sub FilialReq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRequisicoes.objControle = FilialReq
    lErro = Grid_Campo_Libera_Foco(objGridRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Fornecedor_Change()

    giFornecedorAlterado = REGISTRO_ALTERADO
    giTabSelecao_Alterado = REGISTRO_ALTERADO

End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Validate

    'Verifica se Fornecedor foi alterado
    If giFornecedorAlterado = 0 Then Exit Sub

    'Verifica se o Fornecedor esta preenchido
    If Len(Trim(Fornecedor.Text)) > 0 Then

        'Le o Fornecedor
        lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
        If lErro <> SUCESSO Then gError 63440

        'Le as Filiais do Fornecedor
        lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
        If lErro <> SUCESSO And lErro <> 6698 Then gError 63441

        'Preenche a combo FilialFornec
        Call CF("Filial_Preenche", FilialFornec, colCodigoNome)

        'Seleciona a filial na combo de FilialFornec
        Call CF("Filial_Seleciona", FilialFornec, iCodFilial)

    End If

    'Se o Fornecedor nao estiver preenchido
    If Len(Trim(Fornecedor.Text)) = 0 Then

        'Limpa a combo FilialForn
        FilialFornec.Clear

    End If

    giFornecedorAlterado = 0

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True

    Select Case gErr

        Case 63440, 63441
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161350)

    End Select

    Exit Sub

End Sub

Private Sub FornecedorGrid_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub
Private Sub FornecedorGrid_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub FornecedorGrid_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub FornecedorGrid_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = FornecedorGrid
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub FornecedorItem_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FornecedorItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensRequisicoes)

End Sub

Private Sub FornecedorItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensRequisicoes)

End Sub

Private Sub FornecedorItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensRequisicoes.objControle = FornecedorItem
    lErro = Grid_Campo_Libera_Foco(objGridItensRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub FornecedorLabel_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As New Collection

    'Preenche nome reduzido do Fornecedor em objFornecedor
    objFornecedor.sNomeReduzido = Fornecedor.Text

    'Chama a tela FornecedorLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

End Sub

Private Sub FornecedorProd_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FornecedorProd_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub FornecedorProd_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub FornecedorProd_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = FornecedorProd
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub GridItensRequisicoes_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItensRequisicoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItensRequisicoes, giAlterado)
    End If

End Sub

Private Sub GridItensRequisicoes_GotFocus()
    Call Grid_Recebe_Foco(objGridItensRequisicoes)
End Sub

Private Sub GridItensRequisicoes_EnterCell()
    Call Grid_Entrada_Celula(objGridItensRequisicoes, giAlterado)
End Sub

Private Sub GridItensRequisicoes_LeaveCell()
    Call Saida_Celula(objGridItensRequisicoes)
End Sub

Private Sub GridItensRequisicoes_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridItensRequisicoes)
End Sub

Private Sub GridItensRequisicoes_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItensRequisicoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItensRequisicoes, giAlterado)
    End If

End Sub

Private Sub GridItensRequisicoes_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridItensRequisicoes)
End Sub

Private Sub GridItensRequisicoes_RowColChange()
    Call Grid_RowColChange(objGridItensRequisicoes)
End Sub

Private Sub GridItensRequisicoes_Scroll()
    Call Grid_Scroll(objGridItensRequisicoes)
End Sub

Private Sub GridRequisicoes_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridRequisicoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRequisicoes, giAlterado)
    End If

End Sub

Private Sub GridRequisicoes_GotFocus()
    Call Grid_Recebe_Foco(objGridRequisicoes)
End Sub

Private Sub GridRequisicoes_EnterCell()
    Call Grid_Entrada_Celula(objGridRequisicoes, giAlterado)
End Sub

Private Sub GridRequisicoes_LeaveCell()
    Call Saida_Celula(objGridRequisicoes)
End Sub

Private Sub GridRequisicoes_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridRequisicoes)
End Sub

Private Sub GridRequisicoes_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridRequisicoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRequisicoes, giAlterado)
    End If

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

Private Sub Item_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Item_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensRequisicoes)

End Sub

Private Sub Item_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensRequisicoes)

End Sub

Private Sub Item_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensRequisicoes.objControle = Item
    lErro = Grid_Campo_Libera_Foco(objGridItensRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor
Dim bCancel As Boolean

    Set objFornecedor = obj1

    'Preenche o Fornecedor com o NomeReduzido
    Fornecedor.Text = objFornecedor.sNomeReduzido

    Fornecedor_Validate (bCancel)

    Me.Show

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name
        
            'Se for o GridItensRequisicoes
            Case GridRequisicoes.Name

                lErro = Saida_Celula_GridRequisicoes(objGridInt)
                If lErro <> SUCESSO Then gError 77020

            'Se for o GridItensRequisicoes
            Case GridItensRequisicoes.Name

                lErro = Saida_Celula_GridItensRequisicoes(objGridInt)
                If lErro <> SUCESSO Then gError 63470

            'Se for o GridProdutos
            Case GridProdutos.Name

                lErro = Saida_Celula_GridProdutos(objGridInt)
                If lErro <> SUCESSO Then gError 63471

            'Se for o GridFornecedores
            Case GridFornecedores.Name

                lErro = Saida_Celula_GridFornecedores(objGridInt)
                If lErro <> SUCESSO Then gError 63472

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 63473

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 63470 To 63473, 77020
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161351)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridProdutos(objGridInt As AdmGrid) As Long
'Faz a critica da celula do GridProdutos que esta deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridProdutos

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'EscolhidoProd
        Case iGrid_EscolhidoProd_Col
            lErro = Saida_Celula_EscolhidoProd(objGridInt)
            If lErro <> SUCESSO Then gError 77057
        
        'QuantidadeProd
        Case iGrid_QuantidadeProd_Col
            lErro = Saida_Celula_QuantidadeProd(objGridInt)
            If lErro <> SUCESSO Then gError 63474

    End Select

    Saida_Celula_GridProdutos = SUCESSO

    Exit Function

Erro_Saida_Celula_GridProdutos:

    Saida_Celula_GridProdutos = gErr

    Select Case gErr

        Case 63474, 77057

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161352)

    End Select

    Exit Function

End Function
Private Function Saida_Celula_GridRequisicoes(objGridInt As AdmGrid) As Long
'Faz a critica da celula do GridItensRequisicoes que esta deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridRequisicoes

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'EscolhidoReq
        Case iGrid_EscolhidoReq_Col
            lErro = Saida_Celula_EscolhidoReq(objGridInt)
            If lErro <> SUCESSO Then gError 77021

    End Select

    Saida_Celula_GridRequisicoes = SUCESSO

    Exit Function

Erro_Saida_Celula_GridRequisicoes:

    Saida_Celula_GridRequisicoes = gErr

    Select Case gErr

        Case 77021

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161353)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridItensRequisicoes(objGridInt As AdmGrid) As Long
'Faz a critica da celula do GridItensRequisicoes que esta deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridItensRequisicoes

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'EscolhidoItem
        Case iGrid_EscolhidoItem_Col
            lErro = Saida_Celula_EscolhidoItem(objGridInt)
            If lErro <> SUCESSO Then gError 63477

        'QuantCotarItem
        Case iGrid_QuantCotarItem_Col
            lErro = Saida_Celula_QuantCotarItem(objGridInt)
            If lErro <> SUCESSO Then gError 63478

    End Select

    Saida_Celula_GridItensRequisicoes = SUCESSO

    Exit Function

Erro_Saida_Celula_GridItensRequisicoes:

    Saida_Celula_GridItensRequisicoes = gErr

    Select Case gErr

        Case 63477, 63478

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161354)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridFornecedores(objGridInt As AdmGrid) As Long
'Faz a critica da celula do GridItensFornecedores que esta deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridFornecedores

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'EscolhidoForn
        Case iGrid_EscolhidoForn_Col
            lErro = Saida_Celula_EscolhidoForn(objGridInt)
            If lErro <> SUCESSO Then gError 63479

        'ObservacaoForn
        Case iGrid_ObservacaoForn_Col
            lErro = Saida_Celula_ObservacaoForn(objGridInt)
            If lErro <> SUCESSO Then gError 63480

    End Select

    Saida_Celula_GridFornecedores = SUCESSO

    Exit Function

Erro_Saida_Celula_GridFornecedores:

    Saida_Celula_GridFornecedores = gErr

    Select Case gErr

        Case 63479, 63480

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161355)

    End Select

    Exit Function

End Function
Private Function Saida_Celula_EscolhidoReq(objGridInt As AdmGrid) As Long
'Faz a saida de celula de EscolhidoReq

Dim lErro As Long
Dim objReqCompras As ClassRequisicaoCompras
Dim lCodReqCompras As Long
Dim iLinha As Integer

On Error GoTo Erro_Saida_Celula_EscolhidoReq

    Set objGridInt.objControle = EscolhidoReq

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 77022

    Saida_Celula_EscolhidoReq = SUCESSO

    Exit Function

Erro_Saida_Celula_EscolhidoReq:

    Saida_Celula_EscolhidoReq = gErr

    Select Case gErr
        
        Case 77022
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161356)

    End Select

    Exit Function

End Function
Private Function Saida_Celula_EscolhidoProd(objGridInt As AdmGrid) As Long
'Faz a saida de celula de EscolhidoProd

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_EscolhidoProd

    Set objGridInt.objControle = EscolhidoProd
    

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 77062

    Saida_Celula_EscolhidoProd = SUCESSO

    Exit Function

Erro_Saida_Celula_EscolhidoProd:

    Saida_Celula_EscolhidoProd = gErr

    Select Case gErr

        Case 77062
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161357)

    End Select

    Exit Function

End Function

Private Function CotacaoProduto_Escolhe(iLinha As Integer, objCotProduto As ClassCotacaoProduto) As Long
'Faz objCotProduto apontar para CotacaoProduto correspondente na gcolCotacaoProduto

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objCotProdutoGrid As New ClassCotacaoProduto

On Error GoTo Erro_CotacaoProduto_Escolhe
   
    'Coloca o produto no formato do BD
    lErro = CF("Produto_Formata", GridProdutos.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 77058

    'Produto
    objCotProdutoGrid.sProduto = sProdutoFormatado
   
    objFornecedor.sNomeReduzido = Trim(GridProdutos.TextMatrix(iLinha, iGrid_FornecedorProd_Col))
    
    If Len(Trim(objFornecedor.sNomeReduzido)) > 0 Then
    
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 77059
        
        If lErro = 6681 Then gError 77060
        
        'Fornecedor
        objCotProdutoGrid.lFornecedor = objFornecedor.lCodigo
            
    Else
        
        'Fornecedor
        objCotProdutoGrid.lFornecedor = 0
    
    End If

    'FilialFornecedor
    objCotProdutoGrid.iFilial = Codigo_Extrai(GridProdutos.TextMatrix(iLinha, iGrid_FilialFornProd_Col))
    
    'Faz objCotProduto apontar para elemento correspondente na coleção global
    For Each objCotProduto In gcolCotacaoProduto
        If objCotProduto.sProduto = objCotProdutoGrid.sProduto And objCotProduto.lFornecedor = objCotProdutoGrid.lFornecedor And objCotProduto.iFilial = objCotProdutoGrid.iFilial Then Exit For
    Next

    CotacaoProduto_Escolhe = SUCESSO

    Exit Function

Erro_CotacaoProduto_Escolhe:

    CotacaoProduto_Escolhe = gErr

    Select Case gErr

        Case 77058, 77059
            'Erro tratado na rotina chamada

        Case 77060
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161358)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_EscolhidoItem(objGridInt As AdmGrid) As Long
'Faz a saida de celula de EscolhidoItem

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_EscolhidoItem

    Set objGridInt.objControle = EscolhidoItem
    

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 63481

    Saida_Celula_EscolhidoItem = SUCESSO

    Exit Function

Erro_Saida_Celula_EscolhidoItem:

    Saida_Celula_EscolhidoItem = gErr

    Select Case gErr

        Case 63481, 77037, 77038, 77044
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161359)

    End Select

    Exit Function

End Function
Private Function ItemReqCompras_Escolhe(iLinha As Integer, objItemReqCompras As ClassItemReqCompras) As Long
'Faz objItemReqCompras apontar para elemento correspondente na coleção global

Dim lErro As Long
Dim objReqCompras As ClassRequisicaoCompras
Dim lCodReqCompras As Long
Dim iItem As Integer

On Error GoTo Erro_ItemReqCompras_Escolhe

    'Extrai o Codigo da Requisicao e o número de Item do Grid
    lCodReqCompras = StrParaLong(GridItensRequisicoes.TextMatrix(iLinha, iGrid_Requisicao_Col))
    iItem = StrParaInt(GridItensRequisicoes.TextMatrix(iLinha, iGrid_Item_Col))

    'Seleciona objReqCompras cujo código comparece na linha do GridItens
    For Each objReqCompras In gobjGeracaoCotacao.colReqCompra
        If objReqCompras.lCodigo = lCodReqCompras Then Exit For
    Next

    'Seleciona o ItemReqCompras correspondente a linha do Grid
    For Each objItemReqCompras In objReqCompras.colItens
        If objItemReqCompras.iItem = iItem Then Exit For
    Next

    ItemReqCompras_Escolhe = SUCESSO

    Exit Function

Erro_ItemReqCompras_Escolhe:

    ItemReqCompras_Escolhe = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161360)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ObservacaoForn(objGridInt As AdmGrid) As Long
'Faz a saida de celula de ObservacaoForn

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ObservacaoForn

    Set objGridInt.objControle = ObservacaoForn

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 63483

    Saida_Celula_ObservacaoForn = SUCESSO

    Exit Function

Erro_Saida_Celula_ObservacaoForn:

    Saida_Celula_ObservacaoForn = gErr

    Select Case gErr

        Case 63483
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161361)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_EscolhidoForn(objGridInt As AdmGrid) As Long
'Faz a saida de celula de EscolhidoForn

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_EscolhidoForn

    Set objGridInt.objControle = EscolhidoForn
    

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 77075

    Saida_Celula_EscolhidoForn = SUCESSO

    Exit Function

Erro_Saida_Celula_EscolhidoForn:

    Saida_Celula_EscolhidoForn = gErr

    Select Case gErr

        Case 77074, 77075
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161362)

    End Select

    Exit Function

End Function
Private Function ItemGridFornecedor_Escolhe(iLinha As Integer, objItemGridFornecedor As ClassItemGridFornecedores, objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'Faz objItemGridFornecedor e objFornecedorProdutoFF apontarem para elementos correspondentes na coleções globais

Dim lErro As Long
Dim objItemGridFornTela As New ClassItemGridFornecedores

On Error GoTo Erro_ItemGridFornecedor_Escolhe
        
    objItemGridFornTela.sProduto = GridFornecedores.TextMatrix(iLinha, iGrid_ProdutoForn_Col)
    objItemGridFornTela.sFornecedor = GridFornecedores.TextMatrix(iLinha, iGrid_FornecedorGrid_Col)
    objItemGridFornTela.sFilialForn = GridFornecedores.TextMatrix(iLinha, iGrid_FilialFornGrid_Col)
    objItemGridFornTela.sExclusivo = GridFornecedores.TextMatrix(iLinha, iGrid_Exclusivo_Col)

    'Faz objItemGridFornecedor apontar para elemento correspondente na coleção global
    For Each objItemGridFornecedor In gcolItemGridFornecedores
        
        If objItemGridFornTela.sProduto = objItemGridFornecedor.sProduto And objItemGridFornTela.sFornecedor = objItemGridFornecedor.sFornecedor And objItemGridFornTela.sFilialForn = objItemGridFornecedor.sFilialForn And objItemGridFornTela.sExclusivo = objItemGridFornecedor.sExclusivo Then
            Exit For
        End If
    
    Next
    
    'Faz objFornecedorProdutoFF apontar para elemento correspondente na coleção global
    For Each objFornecedorProdutoFF In gcolFornecedorProdutoFF
        
        If objFornecedorProdutoFF.lNumIntDoc = objItemGridFornecedor.lNumIntFornecedorProdutoFF Then
            Exit For
        End If
    
    Next

    ItemGridFornecedor_Escolhe = SUCESSO

    Exit Function

Erro_ItemGridFornecedor_Escolhe:

    ItemGridFornecedor_Escolhe = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161363)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantidadeProd(objGridInt As AdmGrid) As Long
'Faz a saida de celula de da coluna QuantidadeProd do GridProdutos

Dim lErro As Long
Dim dQuantidade As Double
Dim objCotacaoProduto As ClassCotacaoProduto
Dim iLinha As Integer
Dim objItemReqCompra As ClassItemReqCompras
Dim dQuantCotarItens As Double
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim dQuantProdutoAnterior As Double

On Error GoTo Erro_Saida_Celula_QuantidadeProd

     Set objGridInt.objControle = QuantidadeProd

    iLinha = GridProdutos.Row
    
    'Faz objCotacaoProduto apontar para elemento correspondente na coleção global
    lErro = CotacaoProduto_Escolhe(iLinha, objCotacaoProduto)
    If lErro <> SUCESSO Then gError 77069

    'Verifica se a quantidade esta preeenchida
    If Len(Trim(QuantidadeProd.ClipText)) > 0 Then

        'Critica a quantidade
        lErro = Valor_Positivo_Critica(QuantidadeProd.Text)
        If lErro <> SUCESSO Then gError 63475

        dQuantidade = StrParaDbl(QuantidadeProd.Text)

        'Coloca a quantidade com o formato de estoque da tela
        QuantidadeProd.Text = Formata_Estoque(dQuantidade)
         
        'Recolhe o Produto de objItemReqCompras
        objProduto.sCodigo = objCotacaoProduto.sProduto
        
        'Lê os dados do produto envolvido
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 77070
        If lErro <> SUCESSO Then gError 77071
        
        dQuantCotarItens = 0
        
        'Soma quantidades cotar dos ítens associados selecionados
        For Each objItemReqCompra In objCotacaoProduto.colItemReqCompras
            If objItemReqCompra.iSelecionado = Selecionado Then
                
                'Converte para a Unidade de Medida de Compras
                lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemReqCompra.sUM, objCotacaoProduto.sUM, dFator)
                If lErro <> SUCESSO Then gError 77072
                
                dQuantCotarItens = dQuantCotarItens + objItemReqCompra.dQuantCotar * dFator
            
            End If
        Next

        'Verifica se Quantidade é menor que a soma das QuantCotar dos ítensReqCompra
        If dQuantidade < dQuantCotarItens Then gError 77073

    End If

    dQuantProdutoAnterior = StrParaDbl(GridProdutos.TextMatrix(iLinha, iGrid_QuantidadeProd_Col))

    'Compara quantidade atual com anterior
    If dQuantidade <> dQuantProdutoAnterior Then
    
        'Atualiza coleção global
        objCotacaoProduto.dQuantidade = dQuantidade
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 63476

    Saida_Celula_QuantidadeProd = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantidadeProd:

    Saida_Celula_QuantidadeProd = gErr

    Select Case gErr

        Case 63475, 63476, 77069, 77070, 77072
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 77071
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 77073
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTPRODUTO_MENOR_QUANTREQUISITADA", gErr, GridProdutos.Row, dQuantCotarItens)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161364)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantCotarItem(objGridInt As AdmGrid) As Long
'Faz a saida de celula de da coluna QuantCotarItem do GridItensRequisicoes

Dim lErro As Long
Dim dQuantCotar As Double
Dim dQuantRequisitada As Double
Dim dQuantPedida As Double
Dim dQuantCancelada As Double
Dim dQuantRecebida As Double
Dim dQuantCotarAnterior As Double
Dim objItemReqCompras As ClassItemReqCompras
Dim iLinha As Integer
Dim dQuantFaltaCotar As Double

On Error GoTo Erro_Saida_Celula_QuantCotarItem

    Set objGridInt.objControle = QuantCotarItem

    dQuantCotar = 0

    'Verifica se a quantidade esta preeenchida
    If Len(Trim(QuantCotarItem.ClipText)) > 0 Then

        'Critica a quantidade
        lErro = Valor_Positivo_Critica(QuantCotarItem.Text)
        If lErro <> SUCESSO Then gError 63484

        dQuantCotar = StrParaDbl(QuantCotarItem.Text)

        'Coloca a quantidade com o formato de estoque da tela
        QuantCotarItem.Text = Formata_Estoque(dQuantCotar)

        dQuantRequisitada = StrParaDbl(GridItensRequisicoes.TextMatrix(GridItensRequisicoes.Row, iGrid_Quantidade_Col))
        dQuantPedida = StrParaDbl(GridItensRequisicoes.TextMatrix(GridItensRequisicoes.Row, iGrid_QuantPedida_Col))
        dQuantRecebida = StrParaDbl(GridItensRequisicoes.TextMatrix(GridItensRequisicoes.Row, iGrid_QuantRecebida_Col))
        dQuantCancelada = StrParaDbl(GridItensRequisicoes.TextMatrix(GridItensRequisicoes.Row, iGrid_QuantCancelada_Col))
        dQuantFaltaCotar = dQuantRequisitada - dQuantPedida - dQuantRecebida - dQuantCancelada
        
        'Se a quantidade à cotar é maior que a Quantidade Requisitada erro
        If dQuantCotar > dQuantFaltaCotar Then gError 70393
    
    'Se não está preenchida
    Else
        
        'Erro
        gError 79990

    End If
        
    dQuantCotarAnterior = StrParaDbl(GridItensRequisicoes.TextMatrix(GridItensRequisicoes.Row, iGrid_QuantCotarItem_Col))
    
    'Compara quantidade anterior com atual
    If dQuantCotar <> dQuantCotarAnterior Then
    
        iLinha = GridItensRequisicoes.Row
        
        'Faz objItemReqCompras apontar para elemento correspondente na coleção global
        lErro = ItemReqCompras_Escolhe(iLinha, objItemReqCompras)
        If lErro <> SUCESSO Then gError 77047
        
        'Passa quantidades cotar para objItemReqCompras
        objItemReqCompras.dQuantCotarAnterior = objItemReqCompras.dQuantCotar
        objItemReqCompras.dQuantCotar = dQuantCotar
        
        'Se ítem estiver selecionado
        If objItemReqCompras.iSelecionado = Selecionado Then
        
            'Atualiza colecoes globais
            lErro = Atualiza_QuantCotar_ItemReqCompras(objItemReqCompras, gcolCotacaoProduto, gcolFornecedorProdutoFF, gcolItemGridFornecedores)
            If lErro <> SUCESSO Then gError 77049
            
            'Só repercute para outros TABS se estiver selecionado
            giTabItens_Alterado = REGISTRO_ALTERADO
        
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 63485

    Saida_Celula_QuantCotarItem = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantCotarItem:

    Saida_Celula_QuantCotarItem = gErr

    Select Case gErr

        Case 63484, 63485, 77047, 77049
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 70393
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTCOTACAO_MAIOR_QUANTREQUISITADA", gErr, GridItensRequisicoes.Row, dQuantFaltaCotar)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 79990
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA", gErr, GridItensRequisicoes.Row)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161365)

    End Select

    Exit Function

End Function

Private Sub ObservacaoForn_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ObservacaoForn_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub ObservacaoForn_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub ObservacaoForn_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = ObservacaoForn
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ObservacaoItem_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub
Private Sub ObservacaoItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensRequisicoes)

End Sub

Private Sub ObservacaoItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensRequisicoes)

End Sub

Private Sub ObservacaoItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensRequisicoes.objControle = ObservacaoItem
    lErro = Grid_Campo_Libera_Foco(objGridItensRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub


Private Sub ObservacaoReq_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRequisicoes)

End Sub

Private Sub ObservacaoReq_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRequisicoes)

End Sub

Private Sub ObservacaoReq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRequisicoes.objControle = ObservacaoReq
    lErro = Grid_Campo_Libera_Foco(objGridRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Ordena_GridFornecedor()

Dim colItensGridFornecedoresEntrada As New Collection
Dim colItensGridFornecedoresSaida As New Collection
Dim colCampos As New Collection
Dim lErro As Long
Dim colExclusivo As New Collection

On Error GoTo Erro_Ordena_GridFornecedor

    'Recolhe os itens do GridFornecedores
    lErro = GridFornecedores_Recolhe(colItensGridFornecedoresEntrada, colExclusivo)
    If lErro <> SUCESSO Then gError 63489
    
    Call Monta_Colecao_Campos_Fornecedor(colCampos, OrdemFornecedor.ListIndex)
    
    lErro = Ordena_Colecao(colItensGridFornecedoresEntrada, colItensGridFornecedoresSaida, colCampos)
    If lErro <> SUCESSO Then gError 63490
    
    'Devolve os elementos ordenados para o  GridFornecedores
    lErro = GridFornecedores_Devolve(colItensGridFornecedoresSaida)
    If lErro <> SUCESSO Then gError 63491
        
    giTabFornecedor_Alterado = 0
    
    Ordena_GridFornecedor = SUCESSO
        
    Exit Function
    
Erro_Ordena_GridFornecedor:

    Ordena_GridFornecedor = gErr
    
    Select Case gErr
        
        Case 63489, 63490, 63491
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161366)

    End Select

    Exit Function
        
End Function

Private Sub OrdemFornecedor_Click()

Dim lErro As Long

On Error GoTo Erro_OrdemFornecedor_Click

    If Len(Trim(gsOrdemFornecedor)) = 0 Then Exit Sub

    'Verifica se Ordenacao da tela é diferente de gsOrdenacao
    If OrdemFornecedor.Text <> gsOrdemFornecedor Then

        lErro = Ordena_GridFornecedor()
        If lErro <> SUCESSO Then gError 64321
        
        gsOrdemFornecedor = OrdemFornecedor.Text

    End If

    Exit Sub

Erro_OrdemFornecedor_Click:

    Select Case gErr

        Case 64321

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161367)

    End Select

    Exit Sub

End Sub

Private Sub OrdemFornecedor_GotFocus()

    'Armazena a ordenacao em gsOrdemFornecedor
    gsOrdemFornecedor = OrdemFornecedor.Text

End Sub

Private Function GridFornecedores_Recolhe(colItensGridFornecedores As Collection, colExclusivo As Collection) As Long
'Recolhe os itens do GridFornecedores e adiciona em colItensGridFornecedores

Dim objItemGridFornecedores As New ClassItemGridFornecedores
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_GridFornecedores_Recolhe

    Set colItensGridFornecedores = New Collection

    'Percorre todas as linhas do GridFornecedores
    For iIndice = 1 To objGridFornecedores.iLinhasExistentes

        Set objItemGridFornecedores = New ClassItemGridFornecedores


        objItemGridFornecedores.sEscolhido = GridFornecedores.TextMatrix(iIndice, iGrid_EscolhidoForn_Col)
        objItemGridFornecedores.sProduto = GridFornecedores.TextMatrix(iIndice, iGrid_ProdutoForn_Col)
        objItemGridFornecedores.sDescProduto = GridFornecedores.TextMatrix(iIndice, iGrid_DescProdutoForn_Col)
        objItemGridFornecedores.sFornecedor = GridFornecedores.TextMatrix(iIndice, iGrid_FornecedorGrid_Col)
        objItemGridFornecedores.sFilialForn = GridFornecedores.TextMatrix(iIndice, iGrid_FilialFornGrid_Col)
        objItemGridFornecedores.sTipoFrete = GridFornecedores.TextMatrix(iIndice, iGrid_Frete_Col)
        objItemGridFornecedores.sUltimaCotacao = GridFornecedores.TextMatrix(iIndice, iGrid_UltimaCotacao_Col)
        objItemGridFornecedores.sDataUltimaCotacao = GridFornecedores.TextMatrix(iIndice, iGrid_ValorCotacao_Col)
        objItemGridFornecedores.sPrazoEntrega = GridFornecedores.TextMatrix(iIndice, iGrid_PrazoEntrega_Col)
        objItemGridFornecedores.sQuantPedida = GridFornecedores.TextMatrix(iIndice, iGrid_QuantPedida_Col)
        objItemGridFornecedores.sQuantRecebida = GridFornecedores.TextMatrix(iIndice, iGrid_QuantRecebida_Col)
        objItemGridFornecedores.sCondicaoPagto = GridFornecedores.TextMatrix(iIndice, iGrid_CondicaoPagto_Col)
        objItemGridFornecedores.sSaldoTitulos = GridFornecedores.TextMatrix(iIndice, iGrid_SaldoTitulos_Col)
        objItemGridFornecedores.sObservacao = GridFornecedores.TextMatrix(iIndice, iGrid_ObservacaoForn_Col)
        objItemGridFornecedores.sDataUltimaCompra = GridFornecedores.TextMatrix(iIndice, iGrid_UltimaCompra_Col)
        objItemGridFornecedores.sExclusivo = GridFornecedores.TextMatrix(iIndice, iGrid_Exclusivo_Col)
        objItemGridFornecedores.iSelecionado = MARCADO
        

        'Adiciona em colItensGridFornecedores
        colItensGridFornecedores.Add objItemGridFornecedores
        colExclusivo.Add GridFornecedores.TextMatrix(iIndice, iGrid_Exclusivo_Col)
    Next

    GridFornecedores_Recolhe = SUCESSO

    Exit Function

Erro_GridFornecedores_Recolhe:

    GridFornecedores_Recolhe = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161368)

    End Select

    Exit Function

End Function

Private Function GridRequisicoes_Recolhe(colRequisicao As Collection) As Long
'Recolhe os itens do GridRequisicao e adiciona em colRequisicao

Dim objReqCompras As New ClassRequisicaoCompras
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_GridRequisicoes_Recolhe

    Set colRequisicao = New Collection

    'Percorre todas as linhas do GridRequisicoes
    For iIndice = 1 To objGridRequisicoes.iLinhasExistentes

        Set objReqCompras = New ClassRequisicaoCompras

        If Len(Trim(GridRequisicoes.TextMatrix(iIndice, iGrid_EscolhidoReq_Col))) > 0 Then
            objReqCompras.iSelecionado = CInt(GridRequisicoes.TextMatrix(iIndice, iGrid_EscolhidoReq_Col))
        Else
            objReqCompras.iSelecionado = 0
        End If
        objReqCompras.iFilialCompra = Codigo_Extrai(GridRequisicoes.TextMatrix(iIndice, iGrid_FilialReq_Col))
        objReqCompras.lCodigo = GridRequisicoes.TextMatrix(iIndice, iGrid_Numero_Col)
        objReqCompras.dtDataLimite = StrParaDate(GridRequisicoes.TextMatrix(iIndice, iGrid_DataLimite_Col))
        objReqCompras.dtData = StrParaDate(GridRequisicoes.TextMatrix(iIndice, iGrid_DataRC_Col))
        objReqCompras.lUrgente = GridRequisicoes.TextMatrix(iIndice, iGrid_Urgente_Col)
        objReqCompras.lRequisitante = LCodigo_Extrai(GridRequisicoes.TextMatrix(iIndice, iGrid_Requisitante_Col))
        objReqCompras.sCcl = GridRequisicoes.TextMatrix(iIndice, iGrid_CclReq_Col)
        objReqCompras.sObservacao = GridRequisicoes.TextMatrix(iIndice, iGrid_Observacao_Col)

        'Adiciona em colRequisicao
        colRequisicao.Add objReqCompras

    Next

    GridRequisicoes_Recolhe = SUCESSO

    Exit Function

Erro_GridRequisicoes_Recolhe:

    GridRequisicoes_Recolhe = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161369)

    End Select

    Exit Function

End Function

Private Function ItensReq_Calcula_QuantCotar(objGeracaoCotacao As ClassGeracaoCotacao) As Long
'Calcula quantidade a cotar default para Itens de Requisicao

Dim lErro As Long
Dim objReqCompras As ClassRequisicaoCompras
Dim objItemReqCompras As ClassItemReqCompras

On Error GoTo Erro_ItensReq_Calcula_QuantCotar
    
    For Each objReqCompras In objGeracaoCotacao.colReqCompra
    
        For Each objItemReqCompras In objReqCompras.colItens
        
            With objItemReqCompras
                .dQuantCotar = .dQuantidade - .dQuantPedida - .dQuantRecebida - .dQuantCancelada
            End With
        
        Next
    
    Next

    ItensReq_Calcula_QuantCotar = SUCESSO

    Exit Function

Erro_ItensReq_Calcula_QuantCotar:

    ItensReq_Calcula_QuantCotar = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161370)

    End Select

    Exit Function

End Function
Private Function GridItens_Preenche(objGeracaoCotacao As ClassGeracaoCotacao) As Long
'Preenche o GridItensRequisicoes com os Itens das Requisicoes de objGeracaoCotacao

Dim lErro As Long
Dim objReqCompras As ClassRequisicaoCompras
Dim objItemReqCompras As ClassItemReqCompras
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objFornecedor As New ClassFornecedor
Dim objFilForn As New ClassFilialFornecedor
Dim sProdutoEnxuto As String
Dim sCclMascarado As String
Dim objObservacao As New ClassObservacao
Dim objFilialEmpresa As New AdmFiliais
Dim iLinha As Long
Dim sProdutoMascarado As String
Dim colAlmoxarifado As New Collection
Dim bAchou As Boolean
Dim iCount As Integer 'Inserido por Wagner

On Error GoTo Erro_GridItens_Preenche
    
    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridItensRequisicoes)

    '#######################################
    'Inserido por Wagner 14/10/2005
    'Grid Dinâmico
    iCount = 0
    For Each objReqCompras In objGeracaoCotacao.colReqCompra
    
        If objReqCompras.iSelecionado = Selecionado Then

            For Each objItemReqCompras In objReqCompras.colItens
                iCount = iCount + 1
            Next
    
        End If
    
    Next
    
    If iCount >= objGridItensRequisicoes.objGrid.Rows Then
        Call Refaz_Grid(objGridItensRequisicoes, iCount)
    End If
    '#######################################
    
    iLinha = 0

    For Each objReqCompras In objGeracaoCotacao.colReqCompra
    
        If objReqCompras.iSelecionado = Selecionado Then

            For Each objItemReqCompras In objReqCompras.colItens
        
                iLinha = iLinha + 1
        
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_EscolhidoItem_Col) = objItemReqCompras.iSelecionado
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_Item_Col) = objItemReqCompras.iItem
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_Requisicao_Col) = objReqCompras.lCodigo
        
                objFilialEmpresa.iCodFilial = objReqCompras.iFilialEmpresa
        
                'Lê a FilialEmpresa
                lErro = CF("FilialEmpresa_Le", objFilialEmpresa, True)
                If lErro <> SUCESSO And lErro <> 27378 Then gError 64307
        
                'Se não encontrou ==>erro
                If lErro = 27378 Then gError 64308
        
                'Preenche a Filial de Requisicao com código e nome reduzido
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_FilialEmpresaItemReq_Col) = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome
        
                'Mascara o Produto
                lErro = Mascara_RetornaProdutoEnxuto(objItemReqCompras.sProduto, sProdutoMascarado)
                If lErro <> SUCESSO Then gError 63572
        
                ProdutoItem.PromptInclude = False
                ProdutoItem.Text = sProdutoMascarado
                ProdutoItem.PromptInclude = True
        
                'Coloca o Produto com máscara no grid
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_ProdutoItem_Col) = ProdutoItem.Text
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_DescProdutoItem_Col) = objItemReqCompras.sDescProduto
        
                'Coloca UM no grid
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_UnidadeMedItem_Col) = objItemReqCompras.sUM
        
                'Coloca as quantidades no grid
                If objItemReqCompras.dQuantCotar > 0 Then
                    GridItensRequisicoes.TextMatrix(iLinha, iGrid_QuantCotarItem_Col) = Formata_Estoque(objItemReqCompras.dQuantCotar)
                Else
                    GridItensRequisicoes.TextMatrix(iLinha, iGrid_QuantCotarItem_Col) = ""
                End If
                
                If objItemReqCompras.dQuantidade > 0 Then
                    GridItensRequisicoes.TextMatrix(iLinha, iGrid_Quantidade_Col) = Formata_Estoque(objItemReqCompras.dQuantidade)
                Else
                    GridItensRequisicoes.TextMatrix(iLinha, iGrid_Quantidade_Col) = ""
                End If
                
                If objItemReqCompras.dQuantPedida > 0 Then
                    GridItensRequisicoes.TextMatrix(iLinha, iGrid_QuantPedida_Col) = Formata_Estoque(objItemReqCompras.dQuantPedida)
                Else
                    GridItensRequisicoes.TextMatrix(iLinha, iGrid_QuantPedida_Col) = ""
                End If

                If objItemReqCompras.dQuantRecebida > 0 Then
                    GridItensRequisicoes.TextMatrix(iLinha, iGrid_QuantRecebida_Col) = Formata_Estoque(objItemReqCompras.dQuantRecebida)
                Else
                    GridItensRequisicoes.TextMatrix(iLinha, iGrid_QuantRecebida_Col) = ""
                End If
                
                If objItemReqCompras.dQuantCancelada > 0 Then
                    GridItensRequisicoes.TextMatrix(iLinha, iGrid_QuantCancelada_Col) = Formata_Estoque(objItemReqCompras.dQuantCancelada)
                Else
                    GridItensRequisicoes.TextMatrix(iLinha, iGrid_QuantCancelada_Col) = ""
                End If
        
                'Verifica se
                If objItemReqCompras.iAlmoxarifado > 0 Then
        
                    Call Busca_Almoxarifado(objItemReqCompras.iAlmoxarifado, objAlmoxarifado, colAlmoxarifado, bAchou)
        
                    If Not bAchou Then
        
                        Set objAlmoxarifado = New ClassAlmoxarifado
            
                        objAlmoxarifado.iCodigo = objItemReqCompras.iAlmoxarifado
        
                        'Lê o Almoxarifado
                        lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                        If lErro <> SUCESSO And lErro <> 25056 Then gError 63538
        
                        'Se não encontrou o Almoxarifado==> erro
                        If lErro = 25056 Then gError 63546
        
                        colAlmoxarifado.Add objAlmoxarifado
                        
                    End If
        
                    'Preenche grid com nome reduzido do almoxarifado
                    GridItensRequisicoes.TextMatrix(iLinha, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido
        
                End If
        
        
                'Verifica se o Ccl está preeenchido
                If Len(Trim(objItemReqCompras.sCcl)) > 0 Then
        
                        sCclMascarado = String(STRING_CCL, 0)
        
                        'Mascara o Ccl
                        lErro = Mascara_MascararCcl(objItemReqCompras.sCcl, sCclMascarado)
                        If lErro <> SUCESSO Then gError 63573
        
                        'Coloca no grid o Ccl mascarado
                        GridItensRequisicoes.TextMatrix(iLinha, iGrid_CclItem_Col) = sCclMascarado
        
                End If
        
                'Verifica se o Fornecedor está preenchido
                If objItemReqCompras.lFornecedor > 0 Then
        
                    objFornecedor.lCodigo = objItemReqCompras.lFornecedor
        
                    'Lê o Fornecedor
                    lErro = CF("Fornecedor_Le", objFornecedor)
                    If lErro <> SUCESSO And lErro <> 12729 Then gError 63539
        
                    'Se não encontrou o Fornecedor ==> erro
                    If lErro = 12729 Then gError 63550
        
                    'Coloca código e nome reduzido do fornecedor no Grid
                    GridItensRequisicoes.TextMatrix(iLinha, iGrid_FornecedorItem_Col) = objFornecedor.sNomeReduzido
        
                End If
        
                'Verifica se a FilialFornecedor está preenchida
                If objItemReqCompras.iFilial > 0 Then
                    
                    objFilForn.lCodFornecedor = objItemReqCompras.lFornecedor
                    objFilForn.iCodFilial = objItemReqCompras.iFilial
        
                    'Lê a Filial do Fornecedor
                    lErro = CF("FilialFornecedor_Le", objFilForn)
                    If lErro <> SUCESSO And lErro <> 12929 Then gError 63540
        
                    'Se não encontrou a Filial do Fornecedor ==> erro
                    If lErro = 12929 Then gError 63548
        
                    'Coloca código e nome da filial do fornecedor no grid
                    GridItensRequisicoes.TextMatrix(iLinha, iGrid_FilialFornItem_Col) = objItemReqCompras.iFilial & SEPARADOR & objFilForn.sNome
        
                    'Verifica se o Item é preferencial
                    If objItemReqCompras.iExclusivo = ITEM_FILIALFORNECEDOR_PREFERENCIAL Then
            
                        GridItensRequisicoes.TextMatrix(iLinha, iGrid_ExclusoItem_Col) = "Preferencial"
            
                    'Verifica se o Item é exclusivo
                    ElseIf objItemReqCompras.iExclusivo = ITEM_FILIALFORNECEDOR_EXCLUSIVO Then
            
                        GridItensRequisicoes.TextMatrix(iLinha, iGrid_ExclusoItem_Col) = "Exclusivo"
            
                    End If
                
                End If
        
        
                'Verifica se Observacao está preenchida
                If Len(Trim(objItemReqCompras.sObservacao)) = 0 And objItemReqCompras.lObservacao > 0 Then
        
                    objObservacao.lNumInt = objItemReqCompras.lObservacao
        
                    'Lê a observacao
                    lErro = CF("Observacao_Le", objObservacao)
                    If lErro <> SUCESSO And lErro <> 53827 Then gError 63577
        
                    'Se não encontrou a Observacao ==> erro
                    If lErro = 53827 Then gError 63578
                    
                    objItemReqCompras.sObservacao = objObservacao.sObservacao
        
                End If
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_ObservacaoItem_Col) = objItemReqCompras.sObservacao
        
            Next

        End If

    Next

    'Atualiza o número de linhas existentes do grid
    objGridItensRequisicoes.iLinhasExistentes = iLinha

    Call Grid_Refresh_Checkbox(objGridItensRequisicoes)

    GridItens_Preenche = SUCESSO

    Exit Function

Erro_GridItens_Preenche:

    GridItens_Preenche = gErr

    Select Case gErr

        Case 63538, 63539, 63540, 63572, 63573, 63577, 64307
            'Erros tratados nas rotinas chamadas

        Case 63546
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", gErr, objAlmoxarifado.iCodigo)

        Case 63548
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objFilForn.iCodFilial, objFornecedor.lCodigo)

        Case 63550
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)

        Case 63578
            Call Rotina_Erro(vbOKOnly, "ERRO_OBSERVACAO_NAO_CADASTRADA", gErr, objObservacao.lNumInt)

        Case 64308
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161371)

    End Select

    Exit Function

End Function

Private Function GridFornecedores_Devolve(colItensGridFornecedores As Collection) As Long
'Devolve os elementos ordenados para o GridFornecedores

Dim lErro As Long
Dim objItemGridFornecedores As New ClassItemGridFornecedores
Dim iLinha As Long

On Error GoTo Erro_GridFornecedores_Devolve
    
    Call Grid_Limpa(objGridFornecedores)
    
    '#######################################
    'Inserido por Wagner 14/10/2005
    'Grid Dinâmico
    If colItensGridFornecedores.Count >= objGridFornecedores.objGrid.Rows Then
        Call Refaz_Grid(objGridFornecedores, colItensGridFornecedores.Count)
    End If
    '#######################################
    
    For Each objItemGridFornecedores In colItensGridFornecedores

        If objItemGridFornecedores.iSelecionado = Selecionado Then
        
            iLinha = iLinha + 1
    
            'Preenche o GridFornecedores
            GridFornecedores.TextMatrix(iLinha, iGrid_EscolhidoForn_Col) = objItemGridFornecedores.sEscolhido
            GridFornecedores.TextMatrix(iLinha, iGrid_ProdutoForn_Col) = objItemGridFornecedores.sProduto
            GridFornecedores.TextMatrix(iLinha, iGrid_DescProdutoForn_Col) = objItemGridFornecedores.sDescProduto
            GridFornecedores.TextMatrix(iLinha, iGrid_FornecedorGrid_Col) = objItemGridFornecedores.sFornecedor
            GridFornecedores.TextMatrix(iLinha, iGrid_FilialFornGrid_Col) = objItemGridFornecedores.sFilialForn
            GridFornecedores.TextMatrix(iLinha, iGrid_Exclusivo_Col) = objItemGridFornecedores.sExclusivo
            GridFornecedores.TextMatrix(iLinha, iGrid_UltimaCotacao_Col) = objItemGridFornecedores.sDataUltimaCotacao
            GridFornecedores.TextMatrix(iLinha, iGrid_ValorCotacao_Col) = objItemGridFornecedores.sUltimaCotacao
            GridFornecedores.TextMatrix(iLinha, iGrid_Frete_Col) = objItemGridFornecedores.sTipoFrete
            GridFornecedores.TextMatrix(iLinha, iGrid_UltimaCompra_Col) = objItemGridFornecedores.sDataUltimaCompra
            GridFornecedores.TextMatrix(iLinha, iGrid_PrazoEntrega_Col) = objItemGridFornecedores.sPrazoEntrega
            GridFornecedores.TextMatrix(iLinha, iGrid_QuantPedidaForn_Col) = objItemGridFornecedores.sQuantPedida
            GridFornecedores.TextMatrix(iLinha, iGrid_QuantRecebidaForn_Col) = objItemGridFornecedores.sQuantRecebida
            GridFornecedores.TextMatrix(iLinha, iGrid_CondicaoPagto_Col) = objItemGridFornecedores.sCondicaoPagto
            GridFornecedores.TextMatrix(iLinha, iGrid_SaldoTitulos_Col) = objItemGridFornecedores.sSaldoTitulos
            GridFornecedores.TextMatrix(iLinha, iGrid_ObservacaoForn_Col) = objItemGridFornecedores.sObservacao
    
        End If
    
    Next
    
    objGridFornecedores.iLinhasExistentes = iLinha
    
    Call Grid_Refresh_Checkbox(objGridFornecedores)
    
    giTabFornecedor_Alterado = 0
    
    GridFornecedores_Devolve = SUCESSO

    Exit Function

Erro_GridFornecedores_Devolve:

    GridFornecedores_Devolve = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161372)

    End Select

    Exit Function

End Function

Private Function GridRequisicoes_Devolve(colItensGridRequisicoes As Collection) As Long
'Devolve os elementos ordenados para o GridRequisicoes

Dim lErro As Long
Dim objRequisicao As New ClassRequisicaoCompras
Dim iLinha As Long
Dim objFilialEmpresa As New AdmFiliais
Dim objRequisitante As New ClassRequisitante

On Error GoTo Erro_GridRequisicoes_Devolve

    Call Grid_Limpa(objGridRequisicoes)

    '#######################################
    'Inserido por Wagner 14/10/2005
    'Grid Dinâmico
    If colItensGridRequisicoes.colReqCompra.Count >= objGridRequisicoes.objGrid.Rows Then
        Call Refaz_Grid(objGridRequisicoes, colItensGridRequisicoes.colReqCompra.Count)
    End If
    '#######################################
    
    For Each objRequisicao In colItensGridRequisicoes

        iLinha = iLinha + 1

        'Preenche o GridRequisicoes
        GridRequisicoes.TextMatrix(iLinha, iGrid_EscolhidoReq_Col) = objRequisicao.iSelecionado

        'Lê a FilialEmpresa
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa, True)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 64311

        'Se não encontrou ==>erro
        If lErro = 27378 Then gError 64312

        'Preenche a Filial de Requisicao com código e nome reduzido
        GridRequisicoes.TextMatrix(iLinha, iGrid_FilialReq_Col) = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome

        GridRequisicoes.TextMatrix(iLinha, iGrid_Numero_Col) = objRequisicao.lCodigo
        If objRequisicao.dtDataLimite <> DATA_NULA Then GridRequisicoes.TextMatrix(iLinha, iGrid_DataLimite_Col) = objRequisicao.dtDataLimite
        If objRequisicao.dtData <> DATA_NULA Then GridRequisicoes.TextMatrix(iLinha, iGrid_DataRC_Col) = objRequisicao.dtData
        GridRequisicoes.TextMatrix(iLinha, iGrid_Urgente_Col) = objRequisicao.lUrgente
        
        If objRequisicao.lRequisitante <> 0 Then
        
            objRequisitante.lCodigo = objRequisicao.lRequisitante
            
            'Lê o Requisitante
            lErro = CF("Requisitante_Le", objRequisitante)
            If lErro <> SUCESSO Then gError 63829
            
            GridRequisicoes.TextMatrix(iLinha, iGrid_Requisitante_Col) = objRequisicao.lRequisitante & SEPARADOR & objRequisitante.sNomeReduzido
        
        End If
        
        GridRequisicoes.TextMatrix(iLinha, iGrid_CclReq_Col) = objRequisicao.sCcl
        GridRequisicoes.TextMatrix(iLinha, iGrid_Observacao_Col) = objRequisicao.sObservacao

    Next

    Call Grid_Refresh_Checkbox(objGridRequisicoes)

    objGridRequisicoes.iLinhasExistentes = iLinha

    GridRequisicoes_Devolve = SUCESSO

    Exit Function

Erro_GridRequisicoes_Devolve:

    GridRequisicoes_Devolve = gErr

    Select Case gErr

        Case 64311, 63829
            'Erros tratados nas rotinas chamadas

        Case 64312
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161373)

    End Select

    Exit Function

End Function

Private Sub OrdemRequisicao_Click()

Dim lErro As Long
Dim colRequisicaoCompraSaida As New Collection
Dim colCampos As New Collection

On Error GoTo Erro_OrdemRequisicao_Click

    'Verifica se Ordenacao da tela é diferente de gsOrdemRequisicao
    If OrdemRequisicao.Text <> gsOrdemRequisicao Then

        Call Monta_Colecao_Campos_Requisicao(colCampos, OrdemRequisicao.ListIndex)

        lErro = Ordena_Colecao(gobjGeracaoCotacao.colReqCompra, colRequisicaoCompraSaida, colCampos)
        If lErro <> SUCESSO Then gError 63487

        Set gobjGeracaoCotacao.colReqCompra = colRequisicaoCompraSaida

        'Devolve os elementos ordenados para o GridRequisicoes
        lErro = GridRequisicao_Preenche(gobjGeracaoCotacao)
        If lErro <> SUCESSO Then gError 63488
        
        'Se houve alteração nas Reqs selecionadas no TabStrip-Click se rearruma todos os tabs
        'Se não houve alteração nas Requisições selecionadas
        If giTabRequisicao_Alterado = 0 Then
        
            'Preenche reordenado correspondentemente o GridItensRequisicoes
            lErro = GridItens_Preenche(gobjGeracaoCotacao)
            If lErro <> SUCESSO Then gError 77036
        
        End If
        
    End If

     gsOrdemRequisicao = OrdemRequisicao.Text

    Exit Sub

Erro_OrdemRequisicao_Click:

    Select Case gErr

        Case 63487, 63488, 77036
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161374)

    End Select

    Exit Sub

End Sub

Private Sub OrdemRequisicao_GotFocus()

    'Armazena a ordenacao em gsOrdemRequisicao
    gsOrdemRequisicao = OrdemRequisicao.Text

End Sub

Private Sub PrazoEntrega_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PrazoEntrega_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub PrazoEntrega_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub PrazoEntrega_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = PrazoEntrega
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Produto_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub
Private Sub Produto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ProdutoForn_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProdutoForn_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub ProdutoForn_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub ProdutoForn_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = ProdutoForn
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ProdutoItem_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub
Private Sub ProdutoItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensRequisicoes)

End Sub

Private Sub ProdutoItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensRequisicoes)

End Sub

Private Sub ProdutoItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensRequisicoes.objControle = ProdutoItem
    lErro = Grid_Campo_Libera_Foco(objGridItensRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantCanceladaItem_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub
Private Sub QuantCanceladaItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensRequisicoes)

End Sub

Private Sub QuantCanceladaItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensRequisicoes)

End Sub

Private Sub QuantCanceladaItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensRequisicoes.objControle = QuantCanceladaItem
    lErro = Grid_Campo_Libera_Foco(objGridItensRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantCotarItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensRequisicoes)

End Sub

Private Sub QuantCotarItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensRequisicoes)

End Sub

Private Sub QuantCotarItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensRequisicoes.objControle = QuantCotarItem
    lErro = Grid_Campo_Libera_Foco(objGridItensRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantidadeItem_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantidadeItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensRequisicoes)

End Sub

Private Sub QuantidadeItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensRequisicoes)

End Sub

Private Sub QuantidadeItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensRequisicoes.objControle = QuantidadeItem
    lErro = Grid_Campo_Libera_Foco(objGridItensRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantidadeProd_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub QuantidadeProd_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub QuantidadeProd_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = QuantidadeProd
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantPedidaForn_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantPedidaForn_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub QuantPedidaForn_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub QuantPedidaForn_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = QuantPedidaForn
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantPedidaItem_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantPedidaItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensRequisicoes)

End Sub

Private Sub QuantPedidaItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensRequisicoes)

End Sub

Private Sub QuantPedidaItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensRequisicoes.objControle = QuantPedidaItem
    lErro = Grid_Campo_Libera_Foco(objGridItensRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantRecebidaForn_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub
Private Sub QuantRecebidaForn_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub QuantRecebidaForn_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub QuantRecebidaForn_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = QuantRecebidaForn
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantRecebidaItem_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantRecebidaItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensRequisicoes)

End Sub

Private Sub QuantRecebidaItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensRequisicoes)

End Sub

Private Sub QuantRecebidaItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensRequisicoes.objControle = QuantRecebidaItem
    lErro = Grid_Campo_Libera_Foco(objGridItensRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Requisicao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRequisicoes)

End Sub

Private Sub Requisicao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRequisicoes)

End Sub

Private Sub Requisicao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRequisicoes.objControle = Requisicao
    lErro = Grid_Campo_Libera_Foco(objGridRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Requisitante_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRequisicoes)

End Sub

Private Sub Requisitante_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRequisicoes)

End Sub

Private Sub Requisitante_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRequisicoes.objControle = Requisitante
    lErro = Grid_Campo_Libera_Foco(objGridRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub SaldoTitulos_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub

Private Sub SaldoTitulos_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub SaldoTitulos_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub SaldoTitulos_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = SaldoTitulos
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub SelecionaDestino_Click()

Dim iIndice As Integer
Dim bCancel As Boolean

    giTabSelecao_Alterado = REGISTRO_ALTERADO

    'Verifica se SelecionaDestino estiver desmarcado
    If SelecionaDestino.Value = vbUnchecked Then

        For iIndice = 0 To 1

            'Desmarca todos os TipoDestino
            TipoDestino(iIndice).Enabled = False

        Next

        'Desabilita os componentes do FrameDestino()
        FilialEmpresa.Enabled = False
        Fornecedor.Enabled = False
        FilialFornec.Enabled = False
        FilialFornLabel.Enabled = False
    
        FornecedorLabel.Enabled = False
        'Limpa os campos do Frame Destino()
        FilialEmpresa.Text = ""
        FilialEmpresaLabel.Enabled = False
    
        Fornecedor.Text = ""
        FilialFornec.Clear
        FornecedorLabel.Enabled = False
        
    'Verifica se SelecionaDestino está marcado
    ElseIf SelecionaDestino.Value = vbChecked Then

         For iIndice = 0 To 1
            'Habilita todos os Tipos de Destino
            TipoDestino(iIndice).Enabled = True
        Next

        FilialEmpresa.Enabled = True
        FilialEmpresaLabel.Enabled = True

        'Se nenhuma FilialEmpresa estiver selecionada
        If FilialEmpresa.ListIndex = -1 Then
            FilialEmpresa_Validate (bCancel)
        End If
        'Habilita todos os campos do FrameDestino()
        Fornecedor.Enabled = True
        FornecedorLabel.Enabled = True

        FilialFornec.Enabled = True
        FilialFornLabel.Enabled = True
        
    End If

    Exit Sub

End Sub

Private Function Atualiza_QuantCotar_ItemReqCompras(objItemReqCompras As ClassItemReqCompras, colCotacaoProduto As Collection, colFornecedorProdutoFF As Collection, colItemGridFornecedores As Collection) As Long
'Atualiza coleções globais a partir da alteração de QuantCotar de ItemReqCompras

Dim lErro As Long

On Error GoTo Erro_Atualiza_QuantCotar_ItemReqCompras
    
    'Atualiza colCotacaoProduto
    lErro = CotacaoProduto_Atualiza_QuantCotar(objItemReqCompras, colCotacaoProduto)
    If lErro <> SUCESSO Then gError 77050

    'Atualiza colFornecedorProdutoFF
    lErro = FornecedoresProdutos_Atualiza(colCotacaoProduto, colFornecedorProdutoFF)
    If lErro <> SUCESSO Then gError 77051
    
    'Atualiza colItemGridFornecedores
    lErro = ItensGridFornecedores_Atualiza(colFornecedorProdutoFF, colItemGridFornecedores)
    If lErro <> SUCESSO Then gError 77052

    Atualiza_QuantCotar_ItemReqCompras = SUCESSO

    Exit Function

Erro_Atualiza_QuantCotar_ItemReqCompras:

    Atualiza_QuantCotar_ItemReqCompras = gErr

    Select Case gErr

        Case 77050, 77051, 77052
        'Erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161375)

    End Select

    Exit Function

End Function

Private Function Atualiza_Selecao_CotacaoProduto(colCotacaoProduto As Collection, colFornecedorProdutoFF As Collection, colItemGridFornecedores As Collection) As Long
'Atualiza coleções globais a partir de selecao/desselecao de ItemReqCompras

Dim lErro As Long

On Error GoTo Erro_Atualiza_Selecao_CotacaoProduto
    
    'Atualiza colFornecedorProdutoFF
    lErro = FornecedoresProdutos_Atualiza(colCotacaoProduto, colFornecedorProdutoFF)
    If lErro <> SUCESSO Then gError 77063
    
    'Atualiza colItemGridFornecedores
    lErro = ItensGridFornecedores_Atualiza(colFornecedorProdutoFF, colItemGridFornecedores)
    If lErro <> SUCESSO Then gError 77064

    Atualiza_Selecao_CotacaoProduto = SUCESSO

    Exit Function

Erro_Atualiza_Selecao_CotacaoProduto:

    Atualiza_Selecao_CotacaoProduto = gErr

    Select Case gErr

        Case 77063, 77064
        'Erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161376)

    End Select

    Exit Function

End Function


Private Function Atualiza_Selecao_ItemReqCompras(objItemReqCompras As ClassItemReqCompras, colCotacaoProduto As Collection, colFornecedorProdutoFF As Collection, colItemGridFornecedores As Collection) As Long
'Atualiza coleções globais a partir de selecao/desselecao de ItemReqCompras

Dim lErro As Long

On Error GoTo Erro_Atualiza_Selecao_ItemReqCompras
    
    'Atualiza colCotacaoProduto
    lErro = CotacaoProduto_Atualiza(objItemReqCompras, colCotacaoProduto)
    If lErro <> SUCESSO Then gError 77033

    'Atualiza colFornecedorProdutoFF
    lErro = FornecedoresProdutos_Atualiza(colCotacaoProduto, colFornecedorProdutoFF)
    If lErro <> SUCESSO Then gError 77039
    
    'Atualiza colItemGridFornecedores
    lErro = ItensGridFornecedores_Atualiza(colFornecedorProdutoFF, colItemGridFornecedores)
    If lErro <> SUCESSO Then gError 77040

    Atualiza_Selecao_ItemReqCompras = SUCESSO

    Exit Function

Erro_Atualiza_Selecao_ItemReqCompras:

    Atualiza_Selecao_ItemReqCompras = gErr

    Select Case gErr

        Case 77033, 77039, 77040
        'Erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161377)

    End Select

    Exit Function

End Function

Private Function Atualiza_SelecaoReqCompra(objReqCompras As ClassRequisicaoCompras, colCotacaoProduto As Collection, colFornecedorProdutoFF As Collection, colItemGridFornecedores As Collection) As Long
'Atualiza coleções globais a partir de selecao/desselecao de ReqCompras

Dim lErro As Long

On Error GoTo Erro_Atualiza_SelecaoReqCompra
    
    'Atualiza colCotacaoProduto
    lErro = CotacoesProduto_Atualiza(objReqCompras, colCotacaoProduto)
    If lErro <> SUCESSO Then gError 77033

    'Atualiza colFornecedorProdutoFF
    lErro = FornecedoresProdutos_Atualiza(colCotacaoProduto, colFornecedorProdutoFF)
    If lErro <> SUCESSO Then gError 77039
    
    'Atualiza colItemGridFornecedores
    lErro = ItensGridFornecedores_Atualiza(colFornecedorProdutoFF, colItemGridFornecedores)
    If lErro <> SUCESSO Then gError 77040

    Atualiza_SelecaoReqCompra = SUCESSO

    Exit Function

Erro_Atualiza_SelecaoReqCompra:

    Atualiza_SelecaoReqCompra = gErr

    Select Case gErr

        Case 77033, 77039, 77040
        'Erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161378)

    End Select

    Exit Function

End Function
Private Function FornecedoresProdutos_Atualiza(colCotacaoProduto As Collection, colFornecedorProdutoFF As Collection) As Long
'Atualiza coleções globais a partir de selecao/desselecao de ReqCompras

Dim lErro As Long
Dim objCotProduto As ClassCotacaoProduto
Dim objFornecedorProdutoFF As ClassFornecedorProdutoFF

On Error GoTo Erro_FornecedoresProdutos_Atualiza
    
    For Each objFornecedorProdutoFF In colFornecedorProdutoFF
        For Each objCotProduto In colCotacaoProduto
            If objCotProduto.lNumIntDoc = objFornecedorProdutoFF.lNumIntCotacaoProduto Then
                If objCotProduto.iSelecionado = Selecionado And objCotProduto.iEscolhido = Selecionado Then
                    objFornecedorProdutoFF.iSelecionado = Selecionado
                Else
                    objFornecedorProdutoFF.iSelecionado = NAO_SELECIONADO
                End If
                
                Exit For
            End If
        Next
    Next
    
    FornecedoresProdutos_Atualiza = SUCESSO

    Exit Function

Erro_FornecedoresProdutos_Atualiza:

    FornecedoresProdutos_Atualiza = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161379)

    End Select

    Exit Function

End Function
Private Function ItensGridFornecedores_Atualiza(colFornecedorProdutoFF As Collection, colItemGridFornecedores As Collection) As Long
'Atualiza coleções globais a partir de selecao/desselecao de ReqCompras

Dim lErro As Long
Dim objFornecedorProdutoFF As ClassFornecedorProdutoFF
Dim objItemGridFornecedores As ClassItemGridFornecedores

On Error GoTo Erro_ItensGridFornecedores_Atualiza
    
    For Each objItemGridFornecedores In colItemGridFornecedores
        For Each objFornecedorProdutoFF In colFornecedorProdutoFF
            If objFornecedorProdutoFF.lNumIntDoc = objItemGridFornecedores.lNumIntFornecedorProdutoFF Then
                objItemGridFornecedores.iSelecionado = objFornecedorProdutoFF.iSelecionado
                Exit For
            End If
        Next
    Next

    ItensGridFornecedores_Atualiza = SUCESSO

    Exit Function

Erro_ItensGridFornecedores_Atualiza:

    ItensGridFornecedores_Atualiza = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161380)

    End Select

    Exit Function

End Function

Private Function CotacaoProduto_Atualiza_QuantCotar(objItemReqCompras As ClassItemReqCompras, colCotacaoProduto As Collection) As Long
'Atualiza colCotacaoProduto a partir de objItemReqCompras

Dim lErro As Long
Dim objItemReqComprasPesq As ClassItemReqCompras
Dim objCotProduto As ClassCotacaoProduto
Dim bEncontrado As Boolean

On Error GoTo Erro_CotacaoProduto_Atualiza_QuantCotar
    
    'Flag para encontro de CotacaoProduto correspondente
    bEncontrado = False
    
    'Localiza objCotProduto correspondente
    For Each objCotProduto In colCotacaoProduto
    
        For Each objItemReqComprasPesq In objCotProduto.colItemReqCompras
        
            If objItemReqComprasPesq.lNumIntDoc = objItemReqCompras.lNumIntDoc Then
            
                'Atualiza CotacaoProduto a partir de ItemReqCompras
                lErro = CotacaoProduto_Atualiza_Alterando(objItemReqCompras, objCotProduto)
                If lErro <> SUCESSO Then gError 77053
            
                'Encontrou e atualizou CotacaoProduto
                bEncontrado = True
                Exit For
            
            End If

        Next

        'Encontrou e atualizou CotacaoProduto
        If bEncontrado = True Then Exit For
    
    Next
        
    CotacaoProduto_Atualiza_QuantCotar = SUCESSO

    Exit Function

Erro_CotacaoProduto_Atualiza_QuantCotar:

    CotacaoProduto_Atualiza_QuantCotar = gErr

    Select Case gErr

        Case 77053
        'Erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161381)

    End Select

    Exit Function

End Function

Private Function CotacaoProduto_Atualiza(objItemReqCompras As ClassItemReqCompras, colCotacaoProduto As Collection) As Long
'Atualiza colCotacaoProduto a partir de objItemReqCompras

Dim lErro As Long
Dim objItemReqComprasPesq As ClassItemReqCompras
Dim objCotProduto As ClassCotacaoProduto
Dim bEncontrado As Boolean

On Error GoTo Erro_CotacaoProduto_Atualiza
    
    'Flag para encontro de CotacaoProduto correspondente
    bEncontrado = False

    'Localiza objCotProduto correspondente
    For Each objCotProduto In colCotacaoProduto
    
        For Each objItemReqComprasPesq In objCotProduto.colItemReqCompras
        
            If objItemReqComprasPesq.lNumIntDoc = objItemReqCompras.lNumIntDoc Then
            
                If objItemReqCompras.iSelecionado = Selecionado Then
                
                    'Atualiza CotacaoProduto com ItemReqCompras que foi selecionado
                    lErro = CotacaoProduto_Atualiza_Adicionando(objItemReqCompras, objCotProduto)
                    If lErro <> SUCESSO Then gError 77028
                    
                ElseIf objItemReqCompras.iSelecionado = NAO_SELECIONADO Then
                
                    'Atualiza CotacaoProduto a partir de ItemReqCompras
                    lErro = CotacaoProduto_Atualiza_Subtraindo(objItemReqCompras, objCotProduto)
                    If lErro <> SUCESSO Then gError 77029
                
                End If
            
                'Encontrou e atualizou CotacaoProduto
                bEncontrado = True
                Exit For
            
            End If

        Next

        'Encontrou e atualizou CotacaoProduto
        If bEncontrado = True Then Exit For
    
    Next
        
    CotacaoProduto_Atualiza = SUCESSO

    Exit Function

Erro_CotacaoProduto_Atualiza:

    CotacaoProduto_Atualiza = gErr

    Select Case gErr

        Case 77028, 77029
        'Erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161382)

    End Select

    Exit Function

End Function

Private Function CotacoesProduto_Atualiza(objReqCompras As ClassRequisicaoCompras, colCotacaoProduto As Collection) As Long
'Atualiza colCotacaoProduto

Dim lErro As Long
Dim objItemReqCompras As ClassItemReqCompras
Dim objItemReqComprasPesq As ClassItemReqCompras
Dim objCotProduto As ClassCotacaoProduto
Dim bEncontrado As Boolean

On Error GoTo Erro_CotacoesProduto_Atualiza
    
    'Para cada objItemReqCompras
    For Each objItemReqCompras In objReqCompras.colItens
    
        If objItemReqCompras.iSelecionado = Selecionado Then
        
            'Flag para encontro de CotacaoProduto correspondente
            bEncontrado = False
            
            'Localiza objCotProduto correspondente
            For Each objCotProduto In colCotacaoProduto
            
                For Each objItemReqComprasPesq In objCotProduto.colItemReqCompras
                
                    If objItemReqComprasPesq.lNumIntDoc = objItemReqCompras.lNumIntDoc Then
                    
                        If objReqCompras.iSelecionado = Selecionado Then
                        
                            'Atualiza CotacaoProduto com ItemReqCompras que foi selecionado
                            lErro = CotacaoProduto_Atualiza_Adicionando(objItemReqCompras, objCotProduto)
                            If lErro <> SUCESSO Then gError 77041
                            
                        ElseIf objReqCompras.iSelecionado = NAO_SELECIONADO Then
                        
                            'Atualiza CotacaoProduto a partir de ItemReqCompras
                            lErro = CotacaoProduto_Atualiza_Subtraindo(objItemReqCompras, objCotProduto)
                            If lErro <> SUCESSO Then gError 77042
                        
                        End If
                    
                        'Encontrou e atualizou CotacaoProduto
                        bEncontrado = True
                        Exit For
                    
                    End If
        
                Next
        
                'Encontrou e atualizou CotacaoProduto
                If bEncontrado = True Then Exit For
            
            Next
        End If
    Next

    CotacoesProduto_Atualiza = SUCESSO

    Exit Function

Erro_CotacoesProduto_Atualiza:

    CotacoesProduto_Atualiza = gErr

    Select Case gErr

        Case 77041, 77042
        'Erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161383)

    End Select

    Exit Function

End Function
Private Function CotacaoProduto_Atualiza_Alterando(objItemReqCompras As ClassItemReqCompras, objCotProduto As ClassCotacaoProduto) As Long
'Altera quantidade de objCotacaoProduto a partir de objItemReqCompras

Dim lErro As Long
Dim objCotProdutoAuxiliar As ClassCotacaoProduto
Dim objProduto As New ClassProduto
Dim dFator As Double

On Error GoTo Erro_CotacaoProduto_Atualiza_Alterando
                        
    'Cria objeto para adicionar quantidade do ítem no Produto
    Set objCotProdutoAuxiliar = New ClassCotacaoProduto
    
    'Recolhe o Produto de objItemReqCompras
    objProduto.sCodigo = objItemReqCompras.sProduto
    
    'Lê os dados do produto envolvido
    lErro = CF("Produto_Le", objProduto)
    
    If lErro <> SUCESSO And lErro <> 28030 Then gError 77054
    If lErro <> SUCESSO Then gError 77055
    
    objCotProdutoAuxiliar.sProduto = objProduto.sCodigo
    
    'Recolhe a UM de objItemReqCompras
    objCotProdutoAuxiliar.sUM = objItemReqCompras.sUM
        
    'Converte para a Unidade de Medida de Compras
    lErro = CF("UM_Conversao", objProduto.iClasseUM, objCotProdutoAuxiliar.sUM, objCotProduto.sUM, dFator)
    If lErro <> SUCESSO Then gError 77056
                             
    'Recolhe a Quantidade Anterior A Cotar de objItemReqCompras
    objCotProdutoAuxiliar.dQuantidade = objItemReqCompras.dQuantCotarAnterior
    
    'Transforma quantidade para UM de objCotProduto (UMCompra)
    objCotProdutoAuxiliar.dQuantidade = objCotProdutoAuxiliar.dQuantidade * dFator
    
    'Subtrai quantidade de objCotProduto
    objCotProduto.dQuantidade = objCotProduto.dQuantidade - objCotProdutoAuxiliar.dQuantidade

    'Recolhe a Quantidade Atual A Cotar de objItemReqCompras
    objCotProdutoAuxiliar.dQuantidade = objItemReqCompras.dQuantCotar
    
    'Transforma quantidade para UM de objCotProduto (UMCompra)
    objCotProdutoAuxiliar.dQuantidade = objCotProdutoAuxiliar.dQuantidade * dFator
    
    'Adiciona quantidade a objCotProduto
    objCotProduto.dQuantidade = objCotProduto.dQuantidade + objCotProdutoAuxiliar.dQuantidade

    'Seleciona objCotProduto se houver quantidade positiva
    If objCotProduto.dQuantidade > 0 Then
        objCotProduto.iSelecionado = Selecionado
    'Se não, desseleciona objCotProduto
    Else
        objCotProduto.iSelecionado = NAO_SELECIONADO
    End If

    CotacaoProduto_Atualiza_Alterando = SUCESSO

    Exit Function

Erro_CotacaoProduto_Atualiza_Alterando:

    CotacaoProduto_Atualiza_Alterando = gErr

    Select Case gErr

        Case 77054, 77056
        'Erro tratado na rotina chamada
        
        Case 77055
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161384)

    End Select

    Exit Function

End Function

Private Function CotacaoProduto_Atualiza_Adicionando(objItemReqCompras As ClassItemReqCompras, objCotProduto As ClassCotacaoProduto) As Long
'Atualiza objCotacaoProduto a partir de objItemReqCompras selecionado

Dim lErro As Long
Dim objCotProdutoAuxiliar As ClassCotacaoProduto
Dim objProduto As New ClassProduto
Dim dFator As Double

On Error GoTo Erro_CotacaoProduto_Atualiza_Adicionando
                        
    'Cria objeto para adicionar quantidade do ítem no Produto
    Set objCotProdutoAuxiliar = New ClassCotacaoProduto
    
    'Recolhe o Produto de objItemReqCompras
    objProduto.sCodigo = objItemReqCompras.sProduto
    
    'Lê os dados do produto envolvido
    lErro = CF("Produto_Le", objProduto)
    
    If lErro <> SUCESSO And lErro <> 28030 Then gError 77025
    If lErro <> SUCESSO Then gError 77026
    
    objCotProdutoAuxiliar.sProduto = objProduto.sCodigo
    
    'Recolhe a UM de objItemReqCompras
    objCotProdutoAuxiliar.sUM = objItemReqCompras.sUM
        
    'Recolhe a Quantidade a cotar de objItemReqCompras
    objCotProdutoAuxiliar.dQuantidade = objItemReqCompras.dQuantCotar
    
    'Converte para a Unidade de Medida de Compras
    lErro = CF("UM_Conversao", objProduto.iClasseUM, objCotProdutoAuxiliar.sUM, objCotProduto.sUM, dFator)
    If lErro <> SUCESSO Then gError 77027
                             
    'Transforma quantidade para UM de objCotProduto (UMCompra)
    objCotProdutoAuxiliar.sUM = objCotProduto.sUM
    objCotProdutoAuxiliar.dQuantidade = objCotProdutoAuxiliar.dQuantidade * dFator
    
    'Adiciona quantidade a objCotProduto
    objCotProduto.dQuantidade = objCotProduto.dQuantidade + objCotProdutoAuxiliar.dQuantidade

    'Seleciona objCotProduto se houver quantidade positiva
    If objCotProduto.dQuantidade > 0 Then
        objCotProduto.iSelecionado = Selecionado
    End If

    CotacaoProduto_Atualiza_Adicionando = SUCESSO

    Exit Function

Erro_CotacaoProduto_Atualiza_Adicionando:

    CotacaoProduto_Atualiza_Adicionando = gErr

    Select Case gErr

        Case 77025, 77027
        'Erro tratado na rotina chamada
        
        Case 77026
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161385)

    End Select

    Exit Function

End Function
Private Function CotacaoProduto_Atualiza_Subtraindo(objItemReqCompras As ClassItemReqCompras, objCotProduto As ClassCotacaoProduto) As Long
'Atualiza objCotacaoProduto a partir de objItemReqCompras selecionado

Dim lErro As Long
Dim objCotProdutoAuxiliar As ClassCotacaoProduto
Dim objProduto As New ClassProduto
Dim dFator As Double

On Error GoTo Erro_CotacaoProduto_Atualiza_Subtraindo
                        
    'Cria objeto para adicionar quantidade do ítem no Produto
    Set objCotProdutoAuxiliar = New ClassCotacaoProduto
    
    'Recolhe o Produto de objItemReqCompras
    objProduto.sCodigo = objItemReqCompras.sProduto
    
    'Lê os dados do produto envolvido
    lErro = CF("Produto_Le", objProduto)
    
    If lErro <> SUCESSO And lErro <> 28030 Then gError 77030
    If lErro <> SUCESSO Then gError 77031
    
    objCotProdutoAuxiliar.sProduto = objProduto.sCodigo
    
    'Recolhe a UM de objItemReqCompras
    objCotProdutoAuxiliar.sUM = objItemReqCompras.sUM
        
    'Recolhe a Quantidade a cotar de objItemReqCompras
    objCotProdutoAuxiliar.dQuantidade = objItemReqCompras.dQuantCotar
    
    'Converte para a Unidade de Medida de Compras
    lErro = CF("UM_Conversao", objProduto.iClasseUM, objCotProdutoAuxiliar.sUM, objCotProduto.sUM, dFator)
    If lErro <> SUCESSO Then gError 77032
                             
    'Transforma quantidade para UM de objCotProduto (UMCompra)
    objCotProdutoAuxiliar.sUM = objCotProduto.sUM
    objCotProdutoAuxiliar.dQuantidade = objCotProdutoAuxiliar.dQuantidade * dFator
    
    'Subtrai quantidade de objCotProduto
    objCotProduto.dQuantidade = objCotProduto.dQuantidade - objCotProdutoAuxiliar.dQuantidade

    'Se quantidade associada aos ítens zera, desseleciona objCotProduto
    If objCotProduto.dQuantidade < QTDE_ESTOQUE_DELTA Then
        objCotProduto.iSelecionado = NAO_SELECIONADO
    End If
    
    CotacaoProduto_Atualiza_Subtraindo = SUCESSO

    Exit Function

Erro_CotacaoProduto_Atualiza_Subtraindo:

    CotacaoProduto_Atualiza_Subtraindo = gErr

    Select Case gErr

        Case 77030, 77032
        'Erro tratado na rotina chamada
        
        Case 77031
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161386)

    End Select

    Exit Function

End Function

Private Function Ordena_Fornecedor(colItemGridFornecedores As Collection) As Long
'Ordena coleção de Itens de GridFornecedores

Dim colItemGridFornecedoresSaida As New Collection
Dim colCampos As New Collection
Dim lErro As Long

On Error GoTo Erro_Ordena_Fornecedor
   
    Call Monta_Colecao_Campos_Fornecedor(colCampos, OrdemFornecedor.ListIndex)
    
    lErro = Ordena_Colecao(colItemGridFornecedores, colItemGridFornecedoresSaida, colCampos)
    If lErro <> SUCESSO Then gError 63490
    
    Set colItemGridFornecedores = colItemGridFornecedoresSaida
    
    Ordena_Fornecedor = SUCESSO
        
    Exit Function
    
Erro_Ordena_Fornecedor:

    Ordena_Fornecedor = gErr
    
    Select Case gErr
        
        Case 63490
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161387)

    End Select

    Exit Function
        
End Function


Private Function ColItemGridFornecedores_Preenche(colCotacaoProduto As Collection, colFornecedorProdutoFF As Collection, colItemGridFornecedores As Collection) As Long
'Preenche ColItemGridFornecedores

Dim lErro As Long
Dim sProdutoMascarado As String
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim objItemGridFornecedores As ClassItemGridFornecedores
Dim objProduto As New ClassProduto
Dim objFornecedor As New ClassFornecedor
Dim objFilFornecedor As New ClassFilialFornecedor
Dim objFilFornEstendida As New ClassFilialFornecedorEst
Dim objCondPagto As New ClassCondicaoPagto
Dim colFornecedor As New Collection
Dim colProdutos As New Collection
Dim bAchou As Boolean

On Error GoTo Erro_ColItemGridFornecedores_Preenche

    Set colItemGridFornecedores = New Collection

    For Each objFornecedorProdutoFF In colFornecedorProdutoFF

        Set objItemGridFornecedores = New ClassItemGridFornecedores

        objItemGridFornecedores.sEscolhido = CStr(objFornecedorProdutoFF.iEscolhido)
        objItemGridFornecedores.iSelecionado = objFornecedorProdutoFF.iSelecionado
                    
        'Guarda referencia para ítem que o originou na coleção de entrada
        objItemGridFornecedores.lNumIntFornecedorProdutoFF = objFornecedorProdutoFF.lNumIntDoc

        objItemGridFornecedores.sUMCompra = objFornecedorProdutoFF.sUMQuantPedida
        
        'Mascara o Produto
        lErro = Mascara_RetornaProdutoEnxuto(objFornecedorProdutoFF.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 63582

        ProdutoForn.PromptInclude = False
        ProdutoForn.Text = sProdutoMascarado
        ProdutoForn.PromptInclude = True
        
        'Coloca o produto mascarado
        objItemGridFornecedores.sProduto = ProdutoForn.Text
                
        Call Busca_Produto(objFornecedorProdutoFF.sProduto, objProduto, colProdutos, bAchou)
        
        If Not bAchou Then
        
            Set objProduto = New ClassProduto
            
            objProduto.sCodigo = objFornecedorProdutoFF.sProduto
    
            'Lê o Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 63543
    
            'Se não encontrou o Produto ==> erro
            If lErro = 28030 Then gError 63551
            
            colProdutos.Add objProduto
            
        End If
        
        'Preenche a Descricao do produto
        objItemGridFornecedores.sDescProduto = objProduto.sDescricao

        Call Busca_Fornecedor(objFornecedorProdutoFF.lFornecedor, objFornecedor, colFornecedor, bAchou)

        If Not bAchou Then
            
            Set objFornecedor = New ClassFornecedor
            
            'Preenchimento de Fornecedor
            objFornecedor.lCodigo = objFornecedorProdutoFF.lFornecedor
    
            'Lê  o Fornecedor
            lErro = CF("Fornecedor_Le_Basico", objFornecedor)
            If lErro <> SUCESSO And lErro <> 12729 Then gError 63544
    
            'Se não encontrou o Fornecedor ==> erro
            If lErro = 12729 Then gError 63549
            
            colFornecedor.Add objFornecedor
                
        End If
        
        'Preenche nome reduzido do Fornecedor e observação
        objItemGridFornecedores.sFornecedor = objFornecedor.sNomeReduzido
        objItemGridFornecedores.sObservacao = objFornecedor.sObservacao
        
        'Verifica se existe CondicaoPagto para o Fornecedor
        If objFornecedor.iCondicaoPagto <> 0 Then

            objCondPagto.iCodigo = objFornecedor.iCondicaoPagto

            lErro = CF("CondicaoPagto_Le", objCondPagto)
            If lErro <> SUCESSO And lErro <> 19205 Then gError 68342
            If lErro <> SUCESSO Then gError 62862
            

            objItemGridFornecedores.sCondicaoPagto = CStr(objFornecedor.iCondicaoPagto) & SEPARADOR & objCondPagto.sDescReduzida
        
        End If

        'Preenchimento de Filial Fornecedor
        objFilFornecedor.iCodFilial = objFornecedorProdutoFF.iFilialForn
        objFilFornecedor.lCodFornecedor = objFornecedorProdutoFF.lFornecedor
        
        'Lê os dados e Estatística da Filial do Fornecedor
        lErro = CF("FilialFornecedor_Le_Estendida", objFilFornecedor, objFilFornEstendida, False)
        If lErro <> SUCESSO And lErro <> 12929 Then gError 63545

        'Se não encontrou a Filial do Fornecedor ==> erro
        If lErro = 12929 Then gError 63547

        'Preenche o grid com código e nome da Filial do Fornecedor
        objItemGridFornecedores.sFilialForn = CStr(objFornecedorProdutoFF.iFilialForn) & SEPARADOR & objFilFornecedor.sNome
        objItemGridFornecedores.sObservacao = objFilFornecedor.sObservacao

        'Verifica se DataUltimaCotacao é diferente de Data nula
        If objFornecedorProdutoFF.dtDataUltimaCotacao <> DATA_NULA Then
             objItemGridFornecedores.sDataUltimaCotacao = Format(objFornecedorProdutoFF.dtDataUltimaCotacao, "dd/mm/yyyy")
        End If
        
        If objFornecedorProdutoFF.dUltimaCotacao > 0 Then
            objItemGridFornecedores.sUltimaCotacao = Format(objFornecedorProdutoFF.dUltimaCotacao, gobjCOM.sFormatoPrecoUnitario) ' "STANDARD") 'Alterado por Wagner
        End If

        If objFornecedorProdutoFF.iTipoFreteUltimaCotacao = TIPO_FOB Then
            objItemGridFornecedores.sTipoFrete = "FOB"
        ElseIf objFornecedorProdutoFF.iTipoFreteUltimaCotacao = TIPO_CIF Then
            objItemGridFornecedores.sTipoFrete = "CIF"
        End If

        If objFornecedorProdutoFF.dtDataUltimaCompra <> DATA_NULA And objFornecedorProdutoFF.dtDataReceb <> DATA_NULA Then
            objItemGridFornecedores.sPrazoEntrega = CStr(objFornecedorProdutoFF.dtDataReceb - objFornecedorProdutoFF.dtDataUltimaCompra)
        End If
        
            
        objItemGridFornecedores.sQuantPedida = Formata_Estoque(objFornecedorProdutoFF.dQuantPedida)
        objItemGridFornecedores.sQuantRecebida = Formata_Estoque(objFornecedorProdutoFF.dQuantRecebida)

        objFornecedor.lCodigo = objFornecedorProdutoFF.lFornecedor

        objItemGridFornecedores.sSaldoTitulos = Format(objFilFornEstendida.dSaldoTitulos, "Standard")

        'Verifica se DataUltimaCompra é diferente de Data Nula
        If objFilFornEstendida.dtDataUltimaCompra <> DATA_NULA Then
            objItemGridFornecedores.sDataUltimaCompra = Format(objFilFornEstendida.dtDataUltimaCompra, "dd/mm/yyyy")
        End If

        'Coloca exclusividade
        If colCotacaoProduto(objFornecedorProdutoFF.lNumIntCotacaoProduto).lFornecedor > 0 Then
            objItemGridFornecedores.sExclusivo = MARCADO
        Else
            objItemGridFornecedores.sExclusivo = DESMARCADO
        End If

        'Adiciona na coleção
        colItemGridFornecedores.Add objItemGridFornecedores

    Next

    ColItemGridFornecedores_Preenche = SUCESSO

    Exit Function

Erro_ColItemGridFornecedores_Preenche:

    ColItemGridFornecedores_Preenche = gErr

    Select Case gErr
    
        Case 62862
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondPagto.iCodigo)

        Case 63543, 63544, 63545, 63582, 68342, 72343, 72344
            'Erros tratados nas rotinas chamadas

        Case 63547
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objFilFornecedor.iCodFilial, objFornecedor.lCodigo)

        Case 63549
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)

        Case 63551
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case 70519
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161388)

    End Select

    Exit Function

End Function

Private Function Ordena_Produtos(colCotacaoProduto As Collection) As Long

Dim colCampos As New Collection
Dim lErro As Long
Dim colCotacaoProdutoSaida As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Ordena_Produtos
    
    colCampos.Add "sProduto"
    colCampos.Add "lFornecedor"
    colCampos.Add "iFilial"
    
    lErro = Ordena_Colecao(colCotacaoProduto, colCotacaoProdutoSaida, colCampos)
    If lErro <> SUCESSO Then gError 25999
    
    
    Set colCotacaoProduto = colCotacaoProdutoSaida
    
    For iIndice = 1 To colCotacaoProduto.Count
        colCotacaoProduto(iIndice).lNumIntDoc = iIndice
    Next
    
    Ordena_Produtos = SUCESSO
        
    Exit Function
    
Erro_Ordena_Produtos:

    Ordena_Produtos = gErr
    
    Select Case gErr
        
        Case 25999
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161389)

    End Select

    Exit Function
        
End Function
Function Trata_TabFornecedores() As Long
'Traz os dados do TabFornecedores para a tela

Dim lErro As Long
Dim colCotacaoProduto As New Collection
Dim colCotacaoProdutoFornecedor As New Collection
Dim objCotacao As New ClassCotacao
Dim iIndice As Integer
Dim iLinha As Integer

On Error GoTo Erro_Trata_TabFornecedores

    'Preenche o GridFornecedores
    lErro = Preenche_GridFornecedores(colCotacaoProdutoFornecedor)
    If lErro <> SUCESSO Then gError 63502
    
    'Marca Fornecedores Exclusivos
    For iIndice = 1 To objGridItensRequisicoes.iLinhasExistentes
        For iLinha = 1 To objGridFornecedores.iLinhasExistentes
        
            'Se encontrou ItemRC com o mesmo Fornecedor e Filial no Grid de Fornecedores
            If GridItensRequisicoes.TextMatrix(iIndice, iGrid_ProdutoItem_Col) = GridFornecedores.TextMatrix(iLinha, iGrid_ProdutoForn_Col) And GridItensRequisicoes.TextMatrix(iIndice, iGrid_EscolhidoItem_Col) = GRID_CHECKBOX_ATIVO And _
               GridItensRequisicoes.TextMatrix(iIndice, iGrid_FornecedorItem_Col) = GridFornecedores.TextMatrix(iLinha, iGrid_FornecedorGrid_Col) And _
               GridItensRequisicoes.TextMatrix(iIndice, iGrid_FilialFornItem_Col) = GridFornecedores.TextMatrix(iLinha, iGrid_FilialFornGrid_Col) Then
                
                'Verifica se ele é exclusivo
                If GridItensRequisicoes.TextMatrix(iIndice, iGrid_ExclusoItem_Col) = "Exclusivo" Then
                                                    
                    'Se for, marca o Fornecedor e a Filial exclusivos no GridFornecedores
                    GridFornecedores.TextMatrix(iLinha, iGrid_EscolhidoForn_Col) = GRID_CHECKBOX_ATIVO
                    Exit For
                
                End If
            End If
        Next
    Next
    
    Call Grid_Refresh_Checkbox(objGridFornecedores)
    
    Trata_TabFornecedores = SUCESSO

    Exit Function

Erro_Trata_TabFornecedores:

    Trata_TabFornecedores = gErr

    Select Case gErr

        Case 63502
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161390)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objCotacao As New ClassCotacao
Dim colPedidoCotacao As Collection
Dim iIndice2 As Integer
Dim iAchou As Integer
Dim iItemMarcado As Integer
Dim iReqMarcado As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Cotacao.Caption)) = 0 Then gError 62705
    
    'Verifica se existe alguma linha preenchida no GridRequisicoes
    If objGridRequisicoes.iLinhasExistentes = 0 Then gError 63504

    For iIndice = 1 To objGridRequisicoes.iLinhasExistentes

        If StrParaInt(GridRequisicoes.TextMatrix(iIndice, iGrid_EscolhidoReq_Col)) = MARCADO Then
            iReqMarcado = iReqMarcado + 1
        End If

    Next

    'Verifica se existe alguma linha selecionada no GridRequisicoes
    If iReqMarcado = 0 Then gError 63505

    'Verifica se existe alguma linha selecionada no GridItensRequisicoes
    For iIndice = 1 To objGridItensRequisicoes.iLinhasExistentes

        If StrParaInt(GridItensRequisicoes.TextMatrix(iIndice, iGrid_EscolhidoItem_Col)) = MARCADO Then
            iItemMarcado = iItemMarcado + 1
        
            'Verifica se a Quantidade a Cotar está preenchida
            If StrParaDbl(GridItensRequisicoes.TextMatrix(iIndice, iGrid_QuantCotarItem_Col)) = 0 Then gError 63507

        End If

    Next

    'Verifica se nenhuma linha do GridItensRequisicoes está marcada
    If iItemMarcado = 0 Then gError 63506

    'Para cada linha do Grid de Produtos
    For iIndice = 1 To objGridProdutos.iLinhasExistentes

        If StrParaInt(GridProdutos.TextMatrix(iIndice, iGrid_EscolhidoProd_Col)) = MARCADO Then
        
            iAchou = 0
            
            'Verifica se o Produto não é exclusivo
            If Len(Trim(GridProdutos.TextMatrix(iIndice, iGrid_FornecedorProd_Col))) > 0 Then
    
                'Se for exclusivo, procura pelo mesmo Produto marcado no Grid de Fornecedores
                For iIndice2 = 1 To objGridFornecedores.iLinhasExistentes
    
                    'Verifica se para o produto em questao existe pelo menos um Fornecedor escolhido no GridFornecedores
                    If StrParaInt(GridFornecedores.TextMatrix(iIndice2, iGrid_EscolhidoForn_Col)) = MARCADO And GridFornecedores.TextMatrix(iIndice2, iGrid_ProdutoForn_Col) = GridProdutos.TextMatrix(iIndice, iGrid_Produto_Col) And GridFornecedores.TextMatrix(iIndice2, iGrid_Exclusivo_Col) = "1" Then
                        iAchou = 1
                        Exit For
                    End If
                Next
            
            Else
                'Se não for exclusivo, procura pelo mesmo Produto marcado no Grid de Fornecedores
                For iIndice2 = 1 To objGridFornecedores.iLinhasExistentes
    
                    'Verifica se para o produto em questao existe pelo menos um Fornecedor escolhido no GridFornecedores
                    If StrParaInt(GridFornecedores.TextMatrix(iIndice2, iGrid_EscolhidoForn_Col)) = MARCADO And GridFornecedores.TextMatrix(iIndice2, iGrid_ProdutoForn_Col) = GridProdutos.TextMatrix(iIndice, iGrid_Produto_Col) And GridFornecedores.TextMatrix(iIndice2, iGrid_Exclusivo_Col) = "0" Then
                        iAchou = 1
                        Exit For
                    End If
                Next
            
            End If
    
            'Se não existe fornecedor escolhido no GridFornecedores para o produto do GridProdutos ==>erro
            If iAchou = 0 Then gError 63553

        End If

    Next

    'Verifica se SelecionaDestino foi selecionado
    If SelecionaDestino.Value = vbChecked Then

        If TipoDestino(TIPO_DESTINO_EMPRESA) = True Then

            'Se a FilialEmpresa não foi preenchida==>erro
            If Len(Trim(FilialEmpresa.Text)) = 0 Then gError 63508

        ElseIf TipoDestino(TIPO_DESTINO_FORNECEDOR) = True Then

            'Verifica se Fornecedor e FilialFornec estao preenchidos
            If Len(Trim(Fornecedor.Text)) = 0 Then gError 63509
            If Len(Trim(FilialFornec.Text)) = 0 Then gError 63510

        End If

    End If

    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objCotacao, colPedidoCotacao)
    If lErro <> SUCESSO Then gError 63513

    'Gera Cotacao a partir dos Itens de Requisicao
    lErro = CF("Cotacao_Grava_Pedidos", objCotacao, colPedidoCotacao)
    If lErro <> SUCESSO Then gError 63514

    Set gobjCotacao = objCotacao
    Set gcolPedidoCotacao = colPedidoCotacao
    
    Call Limpa_Tela_GeracaoCotacao
    
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = gErr

    Select Case gErr

        Case 62705
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_GERACAO_NAO_PREENCHIDO", gErr)
        
        Case 63504
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_REQUISICOES_VAZIO", gErr)

        Case 63505
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_REQUISICAO_NAO_SELECIONADO", gErr)

        Case 63506
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_ITENS_REQUISICAO_NAO_SELECIONADO", gErr)

        Case 63507
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_COTAR_ITEM_NAO_PREENCHIDA", gErr)

        Case 63508
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_PREENCHIDA", gErr)

        Case 63509
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 63510
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 63513, 63514
            'Erros tratados nas rotinas chamadas

        Case 63553
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_FORNECEDOR_ESCOLHIDO", gErr, GridProdutos.TextMatrix(iIndice, iGrid_Produto_Col))

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161391)

    End Select

    Exit Function

End Function

Private Function Move_Cotacao_Memoria(objCotacao As ClassCotacao) As Long
'Recolhe os dados da tela que não pertencem aos grids para a memória

Dim lErro As Long
Dim objComprador As New ClassComprador
Dim objUsuario As New ClassUsuario
Dim iIndice As Integer
Dim objFornecedor As New ClassFornecedor
Dim iCodigo  As Integer

On Error GoTo Erro_Move_Cotacao_Memoria

    objUsuario.sNomeReduzido = Comprador.Caption

    'Lê o usuario a partir do nome reduzido
    lErro = CF("Usuario_Le_NomeRed", objUsuario)
    If lErro <> SUCESSO And lErro <> 57269 Then gError 63529
    If lErro = 57269 Then gError 63530

    objComprador.sCodUsuario = objUsuario.sCodUsuario

    'Lê o comprador a partir do codUsuario
    lErro = CF("Comprador_Le_Usuario", objComprador)
    If lErro <> SUCESSO And lErro <> 50059 Then gError 63531

    'Se não encontrou o comprador==>erro
    If lErro = 50059 Then gError 63532

    objCotacao.iComprador = objComprador.iCodigo
    objCotacao.iFilialEmpresa = giFilialEmpresa
    objCotacao.sDescricao = Descricao.Text
    objCotacao.dtData = gdtDataAtual
    objCotacao.lCodigo = StrParaLong(Cotacao.Caption)
    iCodigo = COD_A_VISTA
    objCotacao.colCondPagtos.Add (iCodigo)
    
    If Len(Trim(CondPagto.Caption)) > 0 Then
        iCodigo = Codigo_Extrai(CondPagto.Caption)
        objCotacao.colCondPagtos.Add (iCodigo)
    End If
    
    'Se foi selecionado um Tipo de destino
    If SelecionaDestino.Value = vbChecked Then
    
        'Verifica o TipoDestino selecionado
        'Filial Empresa
        If TipoDestino(TIPO_DESTINO_EMPRESA) = True Then
    
            objCotacao.iFilialDestino = Codigo_Extrai(FilialEmpresa.Text)
            objCotacao.iTipoDestino = TIPO_DESTINO_EMPRESA
        
        'Fornecedor
        ElseIf TipoDestino(TIPO_DESTINO_FORNECEDOR) = True Then
    
            objCotacao.iFilialDestino = Codigo_Extrai(FilialFornec.Text)
            objFornecedor.sNomeReduzido = Fornecedor.Text
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then gError 68502
            If lErro = 6681 Then gError 70516
            
            objCotacao.lFornCliDestino = objFornecedor.lCodigo
            
            objCotacao.iTipoDestino = TIPO_DESTINO_FORNECEDOR
        End If
    
    'Se não foi selecionado um Tipo de destino
    Else
        objCotacao.iTipoDestino = TIPO_DESTINO_AUSENTE
    End If

    Move_Cotacao_Memoria = SUCESSO

    Exit Function

Erro_Move_Cotacao_Memoria:

    Move_Cotacao_Memoria = gErr

    Select Case gErr

        Case 63529, 63531, 68502
            'Erros tratados nas rotinas chamadas

        Case 63530
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO2", gErr, objUsuario.sNomeReduzido)

        Case 63532
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_COMPRADOR", gErr, objComprador.sCodUsuario)
                    
        Case 70516
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161392)

    End Select

    Exit Function
    
End Function

Private Function Move_Tela_Memoria(objCotacao As ClassCotacao, colPedidoCotacao As Collection) As Long
'Recolhe os dados da tela

Dim lErro As Long
Dim colCotacaoProduto As New Collection

On Error GoTo Erro_Move_Tela_Memoria

    'Move os dados do TabSelecao
    lErro = Move_Cotacao_Memoria(objCotacao)
    If lErro <> SUCESSO Then gError 63515

    'Recolhe os dados do GridProdutos
    lErro = Move_GridProdutos_Memoria(objCotacao)
    If lErro <> SUCESSO Then gError 63516

    'Recolhe os dados do GridFornecedores
    lErro = Move_GridFornecedores_Memoria(objCotacao, colPedidoCotacao)
    If lErro <> SUCESSO Then gError 63517

    'Recolhe os dados do GridItensRequisicoes
    lErro = Move_GridItens_Memoria(objCotacao)
    If lErro <> SUCESSO Then gError 63518

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 63515 To 63518
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161393)

    End Select

    Exit Function

End Function

Function Move_GridItens_Memoria(objCotacao As ClassCotacao) As Long
'Recolhe os dados do GridItensRequisicoes

Dim lErro As Long
Dim objCotacaoProduto As ClassCotacaoProduto
Dim objItemReqCompra As ClassItemReqCompras
Dim objFornecedor As New ClassFornecedor
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim iLinha As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sCclFormata As String
Dim iCclPreenchida As Integer
Dim sProdutoEnxuto As String
Dim iFilialEmpresa As Integer

On Error GoTo Erro_Move_GridItens_Memoria

    'Para cada Cotação Produto da Cotação
    For Each objCotacaoProduto In objCotacao.colCotacaoProduto

        'Para cada linha do Grid de Itens de Requisição
        For iLinha = 1 To objGridItensRequisicoes.iLinhasExistentes

            'Verifica se a linha foi selecionada
            If GridItensRequisicoes.TextMatrix(iLinha, iGrid_EscolhidoItem_Col) = GRID_CHECKBOX_ATIVO Then
                
                'Formata o Produto
                lErro = CF("Produto_Formata", GridItensRequisicoes.TextMatrix(iLinha, iGrid_ProdutoItem_Col), sProdutoFormatado, iProdutoPreenchido)
                If lErro <> SUCESSO Then gError 63831
                            
                'Se o Fornecedor foi preenchido
                If Len(Trim(GridItensRequisicoes.TextMatrix(iLinha, iGrid_FornecedorItem_Col))) > 0 Then
                    
                    'Lê o Fornecedor
                    objFornecedor.sNomeReduzido = GridItensRequisicoes.TextMatrix(iLinha, iGrid_FornecedorItem_Col)
                    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                    If lErro <> SUCESSO And lErro <> 6681 Then gError 66968
                    
                    'Se não encontrou o Fornecedor, erro
                    If lErro = 6681 Then gError 66969
                
                Else
                    objFornecedor.lCodigo = 0
                End If
                
                'Verifica se o Produto de Cotação Produto é igual ao Produto do Grid amarrando por Fornecedor e Filial, se houver exclusividade
                If objCotacaoProduto.sProduto = sProdutoFormatado And ((objCotacaoProduto.lFornecedor <> 0 And GridItensRequisicoes.TextMatrix(iLinha, iGrid_ExclusoItem_Col) = "Exclusivo" And objFornecedor.lCodigo = objCotacaoProduto.lFornecedor And Codigo_Extrai(GridItensRequisicoes.TextMatrix(iLinha, iGrid_FilialFornItem_Col)) = objCotacaoProduto.iFilial) Or (objCotacaoProduto.lFornecedor = 0 And GridItensRequisicoes.TextMatrix(iLinha, iGrid_ExclusoItem_Col) <> "Exclusivo")) Then
                    
                    Set objItemReqCompra = New ClassItemReqCompras
                    
                    objItemReqCompra.sProduto = sProdutoFormatado
                    objItemReqCompra.sDescProduto = GridItensRequisicoes.TextMatrix(iLinha, iGrid_DescProdutoItem_Col)
                    objItemReqCompra.lFornecedor = objFornecedor.lCodigo
                    objItemReqCompra.iFilial = Codigo_Extrai(GridItensRequisicoes.TextMatrix(iLinha, iGrid_FilialFornItem_Col))
                    objItemReqCompra.sUM = GridItensRequisicoes.TextMatrix(iLinha, iGrid_UnidadeMedItem_Col)
                    objItemReqCompra.dQuantPedida = StrParaDbl(GridItensRequisicoes.TextMatrix(iLinha, iGrid_QuantPedida_Col))
                    objItemReqCompra.dQuantRecebida = StrParaDbl(GridItensRequisicoes.TextMatrix(iLinha, iGrid_QuantRecebida_Col))
                    objItemReqCompra.dQuantCancelada = StrParaDbl(GridItensRequisicoes.TextMatrix(iLinha, iGrid_QuantCancelada_Col))
                    objItemReqCompra.dQuantidade = StrParaDbl(GridItensRequisicoes.TextMatrix(iLinha, iGrid_Quantidade_Col))
                    objItemReqCompra.dQuantCotar = StrParaDbl(GridItensRequisicoes.TextMatrix(iLinha, iGrid_QuantCotarItem_Col))
                                    
                    If Len(Trim(GridItensRequisicoes.TextMatrix(iLinha, iGrid_Almoxarifado_Col))) > 0 Then
                    
                        objAlmoxarifado.sNomeReduzido = GridItensRequisicoes.TextMatrix(iLinha, iGrid_Almoxarifado_Col)
                
                        'Lê dados do almoxarifado a partir do Nome Reduzido
                        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                        If lErro <> SUCESSO And lErro <> 25060 Then gError 70393
                
                        'Se não econtrou o almoxarifado, erro
                        If lErro = 25060 Then gError 70394
                
                        objItemReqCompra.iAlmoxarifado = objAlmoxarifado.iCodigo
                    
                    End If
                    
                    'Verifica se o Ccl foi preenchido
                    If Len(Trim(GridItensRequisicoes.TextMatrix(iLinha, iGrid_CclItem_Col))) > 0 Then
        
                        'Formata o Ccl
                        lErro = CF("Ccl_Formata", GridItensRequisicoes.TextMatrix(iLinha, iGrid_CclItem_Col), sCclFormata, iCclPreenchida)
                        If lErro <> SUCESSO Then gError 63596
        
                        objItemReqCompra.sCcl = sCclFormata
        
                    End If
        
                    'Exclusividade
                    If GridItensRequisicoes.TextMatrix(iLinha, iGrid_ExclusoItem_Col) = "Exclusivo" Then
                        objItemReqCompra.iExclusivo = ITEM_FILIALFORNECEDOR_EXCLUSIVO
                    ElseIf GridItensRequisicoes.TextMatrix(iLinha, iGrid_ExclusoItem_Col) = "Preferencial" Then
                        objItemReqCompra.iExclusivo = ITEM_FILIALFORNECEDOR_PREFERENCIAL
                    End If
        
                    objItemReqCompra.sObservacao = GridItensRequisicoes.TextMatrix(iLinha, iGrid_ObservacaoItem_Col)
                    objItemReqCompra.lReqCompra = StrParaLong(GridItensRequisicoes.TextMatrix(iLinha, iGrid_Requisicao_Col))
                    iFilialEmpresa = Codigo_Extrai(GridItensRequisicoes.TextMatrix(iLinha, iGrid_FilialEmpresaItemReq_Col))
                    
                    'Lê o NumIntDoc do Item da Requisição a partir do Código da Requisição, FilialEmpresa, Produto, Fornecedor e Filial
                    lErro = CF("ItemRC_NumInt_Le", objItemReqCompra, iFilialEmpresa)
                    If lErro <> SUCESSO And lErro <> 70398 Then gError 70399
                    
                    'Se não encontrou o ItemRC, erro
                    If lErro = 70398 Then gError 70400
                    
                    'Adiciona em ColItensRequisicao
                    objCotacaoProduto.colItemReqCompras.Add objItemReqCompra
    
                End If
    
            End If
        
        Next

    Next

    Move_GridItens_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItens_Memoria:

    Move_GridItens_Memoria = gErr

    Select Case gErr

        Case 63596, 63831, 66968, 70393, 70399
        
        Case 66969
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
        
        Case 70394
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO1", gErr, objAlmoxarifado.sNomeReduzido)
        
        Case 70400
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMREQCOMPRA_NAO_CADASTRADO2", gErr, objItemReqCompra.sProduto, objItemReqCompra.lReqCompra)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161394)

    End Select

    Exit Function

End Function

Private Function Geracao_Le_Requisicoes(objGeracaoCotacao As ClassGeracaoCotacao) As Long
'Lê para objGeracaoCotacao as Requisicoes com seus ítens, de acordo com a selecao feita em TabSelecao

Dim lErro As Long
Dim iIndice As Integer
Dim colReqCompras As Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Geracao_Le_Requisicoes
    
    Set objGeracaoCotacao = New ClassGeracaoCotacao
  
    If Len(Trim(DataDe.ClipText)) > 0 And Len(Trim(DataAte.ClipText)) > 0 Then
    
        If StrParaDate(DataDe.Text) > StrParaDate(DataAte.Text) Then gError 63605
        
    End If
    
    'Preenche objGeracaoCotacao com DataDe e DataAte
    objGeracaoCotacao.dtDataDe = StrParaDate(DataDe.Text)
    objGeracaoCotacao.dtDataAte = StrParaDate(DataAte.Text)

    objGeracaoCotacao.sOrdenacaoReq = asOrdemRequisicao(OrdemRequisicao.ListIndex)

    'Verifica se foi selecionado um local de entrega
    If SelecionaDestino.Value = vbChecked Then

        For iIndice = 0 To 1

            'Verifica qual o TipoDestino selecionado
            If TipoDestino(iIndice).Value = True Then objGeracaoCotacao.iTipoDestino = iIndice

        Next

        'Verifica o TipoDestino da Requisicao de Compras
        If objGeracaoCotacao.iTipoDestino = TIPO_DESTINO_EMPRESA Then

            'Preenche FilialDestino com a FilialEmpresa informada na tela
            If Len(Trim(FilialEmpresa.Text)) > 0 Then objGeracaoCotacao.iFilialDestino = Codigo_Extrai(FilialEmpresa.Text)

        ElseIf objGeracaoCotacao.iTipoDestino = TIPO_DESTINO_FORNECEDOR Then

            'Preenche objGeracaoCotacao com Fornecedor e FilialFornecedor informados na tela
            If Len(Trim(Fornecedor.Text)) > 0 Then
            
                objFornecedor.sNomeReduzido = Fornecedor.Text
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO And lErro <> 6681 Then gError 68503
                
                If lErro = 6681 Then gError 70517
                
                objGeracaoCotacao.lFornCliDestino = objFornecedor.lCodigo
            End If

            If Len(Trim(FilialFornec.Text)) > 0 Then objGeracaoCotacao.iFilialDestino = Codigo_Extrai(FilialFornec.Text)
        
        'Se não foi selecionado nenhum tipo destino
        Else
            objGeracaoCotacao.iTipoDestino = TIPO_DESTINO_AUSENTE
        End If

    Else
        objGeracaoCotacao.iTipoDestino = -1
    End If

    If Len(Trim(CodigoDe.Text)) > 0 And Len(Trim(CodigoAte.Text)) > 0 Then
        If StrParaLong(CodigoDe.Text) > StrParaLong(CodigoAte.Text) Then gError 63603
    End If

    'Preenche objGeracaoCotacao com CodigoDe e CodigoAte
    objGeracaoCotacao.lCodigoDe = StrParaLong(CodigoDe.Text)
    objGeracaoCotacao.lCodigoAte = StrParaLong(CodigoAte.Text)

    If Len(Trim(DataLimiteDe.ClipText)) > 0 And Len(Trim(DataLimiteAte.ClipText)) > 0 Then
    
        If StrParaDate(DataLimiteDe.ClipText) > StrParaDate(DataLimiteAte.ClipText) Then gError 63604
    
    End If
    'Preenche objGeracaoCotacao com DataLimiteDe e DataLimiteAte
    objGeracaoCotacao.dtDataLimiteDe = StrParaDate(DataLimiteDe.Text)
    objGeracaoCotacao.dtDataLimiteAte = StrParaDate(DataLimiteAte.Text)

    'Verifica se quer exibir as Requisicoes Cotadas
    If ExibeCotadas.Value = vbChecked Then

        objGeracaoCotacao.iExibeReqCotadas = EXIBE_REQUISICOES_COTADAS

    Else

        objGeracaoCotacao.iExibeReqCotadas = NAO_EXIBE_REQUISICOES_COTADAS

    End If

    'Preenche colecao de tipos de produtos com os Tipos de Produtos selecionados
    Set objGeracaoCotacao.colTipoProduto = New Collection

    For iIndice = 0 To TipoProduto.ListCount - 1
         If TipoProduto.Selected(iIndice) = True Then
            objGeracaoCotacao.colTipoProduto.Add (Codigo_Extrai(TipoProduto.List(iIndice)))
        End If
    Next

    'Verifica se nenhum tipo de produto foi selecionado
    If objGeracaoCotacao.colTipoProduto.Count = 0 Then gError 74870
    
    'Preenche a colecao de Requisicoes com seus ítens
    lErro = CF("ReqCompras_Le_GeracaoCotacao", objGeracaoCotacao)
    If lErro <> SUCESSO Then gError 63541

    Geracao_Le_Requisicoes = SUCESSO

    Exit Function

Erro_Geracao_Le_Requisicoes:

    Geracao_Le_Requisicoes = gErr

    Select Case gErr

        Case 63541, 68503
            'Erros tratados nas rotinas chamadas

        Case 63603
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOINICIAL_MAIOR_REQUISICAOFINAL", gErr)

        Case 63604
            Call Rotina_Erro(vbOKOnly, "ERRO_DATADE_MAIOR_DATAATE", gErr)

        Case 63605
            Call Rotina_Erro(vbOKOnly, "ERRO_DATADE_MAIOR_DATAATE", gErr)

        Case 70517
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
            
        Case 74870
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_TIPOPRODUTO_SELECIONADO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161395)

    End Select

    Exit Function

End Function

Private Function GridRequisicao_Preenche(objGeracaoCotacao As ClassGeracaoCotacao) As Long
'Preenche o GridRequisicoes com as Requisicoes que estão em colReqCompra

Dim lErro As Long
Dim objReqCompras As New ClassRequisicaoCompras
Dim iLinha As Integer
Dim objRequisitante As New ClassRequisitante
Dim sCclMascarado As String
Dim objObservacao As New ClassObservacao
Dim objFilialEmpresa As New AdmFiliais
Dim lCodigoPV As Long

On Error GoTo Erro_GridRequisicao_Preenche

'    'Verifica se existem mais Requisicoes de compra em colReqCompra do que o permitido no grid
'    If objGeracaoCotacao.colReqCompra.Count > NUM_MAX_REQUISICOES_GRID Then
'
'        objGridRequisicoes.objGrid.Rows = objGeracaoCotacao.colReqCompra.Count
'        objGridRequisicoes.objGrid.Refresh
'
'    End If

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridRequisicoes)

    '#######################################
    'Inserido por Wagner 14/10/2005
    'Grid Dinâmico
    If objGeracaoCotacao.colReqCompra.Count >= objGridRequisicoes.objGrid.Rows Then
        Call Refaz_Grid(objGridRequisicoes, objGeracaoCotacao.colReqCompra.Count)
    End If
    '#######################################

    For Each objReqCompras In objGeracaoCotacao.colReqCompra

        iLinha = iLinha + 1
        
        GridRequisicoes.TextMatrix(iLinha, iGrid_EscolhidoReq_Col) = objReqCompras.iSelecionado
        
        objFilialEmpresa.iCodFilial = objReqCompras.iFilialEmpresa

        'Lê a FilialEmpresa
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa, True)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 63580

        'Se não encontrou ==>erro
        If lErro = 27378 Then gError 63581

        'Preenche a Filial de Requisicao com código e nome reduzido
        GridRequisicoes.TextMatrix(iLinha, iGrid_FilialReq_Col) = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome
        GridRequisicoes.TextMatrix(iLinha, iGrid_Numero_Col) = objReqCompras.lCodigo

        'Verifica se Data Limite é diferente de data nula
        If objReqCompras.dtDataLimite <> DATA_NULA Then

            'Coloca a Data Limite no formato da tela
            GridRequisicoes.TextMatrix(iLinha, iGrid_DataLimite_Col) = Format(objReqCompras.dtDataLimite, "dd/mm/yy")
        End If

        'Verifica se Data é diferente de Data Nula
        If objReqCompras.dtData <> DATA_NULA Then

            'Coloca a data no formato da tela
            GridRequisicoes.TextMatrix(iLinha, iGrid_DataRC_Col) = Format(objReqCompras.dtData, "dd/mm/yy")

        End If

        GridRequisicoes.TextMatrix(iLinha, iGrid_Urgente_Col) = objReqCompras.lUrgente

        'Verifica se Requisitante está prenchido
        If objReqCompras.lRequisitante > 0 Then
            objRequisitante.lCodigo = objReqCompras.lRequisitante

            'Lê o Requisitante a partir do código fornecido
            lErro = CF("Requisitante_Le", objRequisitante)
            If lErro <> SUCESSO And lErro <> 49084 Then gError 63565
            If lErro = 49084 Then gError 63566

            'Coloca código e nome reduzido do Requisitante no Grid
            GridRequisicoes.TextMatrix(iLinha, iGrid_Requisitante_Col) = objReqCompras.lRequisitante & SEPARADOR & objRequisitante.sNomeReduzido

        End If

        'Verifica se Ccl foi preenchido
        If Len(Trim(objReqCompras.sCcl)) > 0 Then

            sCclMascarado = String(STRING_CCL, 0)

            'Mascara o Ccl
            lErro = Mascara_MascararCcl(objReqCompras.sCcl, sCclMascarado)
            If lErro <> SUCESSO Then gError 63574

            'Coloca o Ccl mascarado no Grid
            GridRequisicoes.TextMatrix(iLinha, iGrid_CclReq_Col) = sCclMascarado

        End If

        'Verifica se Observacao foi preenchida
        If objReqCompras.lObservacao > 0 Then

            objObservacao.lNumInt = objReqCompras.lObservacao

            'Lê a Observacao
            lErro = CF("Observacao_Le", objObservacao)
            If lErro <> SUCESSO And lErro <> 53827 Then gError 63575
            If lErro = 53827 Then gError 63576

            GridRequisicoes.TextMatrix(iLinha, iGrid_Observacao_Col) = objObservacao.sObservacao

        End If

        If Len(Trim(objReqCompras.sOPCodigo)) > 0 Then

            lErro = Preenche_CodigoPV(objReqCompras, lCodigoPV)
            If lErro <> SUCESSO Then gError 178860
    
            If lCodigoPV <> 0 Then
                GridRequisicoes.TextMatrix(iLinha, iGrid_CodigoPV_Col) = lCodigoPV
            End If

        End If

    Next

    'Atualiza o número de linhas existentes do GridRequisicoes
    objGridRequisicoes.iLinhasExistentes = iLinha

    Call Grid_Refresh_Checkbox(objGridRequisicoes)

    GridRequisicao_Preenche = SUCESSO

    Exit Function

Erro_GridRequisicao_Preenche:

    GridRequisicao_Preenche = gErr

    Select Case gErr

        Case 63565, 63574, 63575, 63580, 178860

        Case 63566
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_NAO_CADASTRADO", gErr, objRequisitante.lCodigo)

        Case 63576
            Call Rotina_Erro(vbOKOnly, "ERRO_OBSERVACAO_NAO_CADASTRADA", gErr, objObservacao.lNumInt)

        Case 63581
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161396)

    End Select

    Exit Function

End Function

Private Sub GridFornecedores_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridFornecedores, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFornecedores, giAlterado)
    End If

End Sub

Private Sub GridFornecedores_GotFocus()
    Call Grid_Recebe_Foco(objGridFornecedores)
End Sub

Private Sub GridFornecedores_EnterCell()
    Call Grid_Entrada_Celula(objGridFornecedores, giAlterado)
End Sub

Private Sub GridFornecedores_LeaveCell()
    Call Saida_Celula(objGridFornecedores)
End Sub

Private Sub GridFornecedores_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridFornecedores)
End Sub

Private Sub GridFornecedores_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridFornecedores, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFornecedores, giAlterado)
    End If

End Sub

Private Sub GridFornecedores_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridFornecedores)
End Sub

Private Sub GridFornecedores_RowColChange()
    Call Grid_RowColChange(objGridFornecedores)
End Sub

Private Sub GridFornecedores_Scroll()
    Call Grid_Scroll(objGridFornecedores)
End Sub

Private Sub GridProdutos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridProdutos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridProdutos, giAlterado)
    End If

End Sub

Private Sub GridProdutos_GotFocus()
    Call Grid_Recebe_Foco(objGridProdutos)
End Sub

Private Sub GridProdutos_EnterCell()
    Call Grid_Entrada_Celula(objGridProdutos, giAlterado)
End Sub

Private Sub GridProdutos_LeaveCell()
    Call Saida_Celula(objGridProdutos)
End Sub

Private Sub GridProdutos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridProdutos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridProdutos, giAlterado)
    End If

End Sub

Private Sub GridProdutos_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridProdutos)
End Sub

Private Sub GridProdutos_RowColChange()
    Call Grid_RowColChange(objGridProdutos)
End Sub

Private Sub GridProdutos_Scroll()
    Call Grid_Scroll(objGridProdutos)
End Sub

Function GridProduto_Preenche(colCotacaoProduto As Collection) As Long
'Preenche o GridProdutos com todos os Produtos da colecao colCotacaoProduto

Dim lErro As Long
Dim objCotProduto As New ClassCotacaoProduto
Dim objProduto As New ClassProduto
Dim objFornecedor As New ClassFornecedor
Dim objFilialForn As New ClassFilialFornecedor
Dim iLinha As Integer
Dim sProdutoEnxuto As String

On Error GoTo Erro_GridProduto_Preenche

    'Limpa o GridProdutos
    Call Grid_Limpa(objGridProdutos)

    '#######################################
    'Inserido por Wagner 14/10/2005
    'Grid Dinâmico
    If colCotacaoProduto.Count >= objGridProdutos.objGrid.Rows Then
        Call Refaz_Grid(objGridProdutos, colCotacaoProduto.Count)
    End If
    '#######################################
    
    For Each objCotProduto In colCotacaoProduto

        If objCotProduto.iSelecionado = Selecionado Then
        
            iLinha = iLinha + 1
    
            GridProdutos.TextMatrix(iLinha, iGrid_EscolhidoProd_Col) = objCotProduto.iEscolhido
    
            'Formata o Produto
            lErro = Mascara_RetornaProdutoEnxuto(objCotProduto.sProduto, sProdutoEnxuto)
            If lErro <> SUCESSO Then gError 63583
    
            Produto.PromptInclude = False
            Produto.Text = sProdutoEnxuto
            Produto.PromptInclude = True
    
            'Coloca o Produto mascarado no Grid
            GridProdutos.TextMatrix(iLinha, iGrid_Produto_Col) = Produto.Text
    
            objProduto.sCodigo = objCotProduto.sProduto
    
            'Lê o Produto a partir do código do produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 63495
    
            'Se não encontrou o Produto ==> erro
            If lErro = 28030 Then gError 63498
    
            GridProdutos.TextMatrix(iLinha, iGrid_Descricao_Col) = objProduto.sDescricao
            GridProdutos.TextMatrix(iLinha, iGrid_UnidadeMed_Col) = objCotProduto.sUM
            GridProdutos.TextMatrix(iLinha, iGrid_QuantidadeProd_Col) = Formata_Estoque(objCotProduto.dQuantidade)
    
            'Verifica se Fornecedor está preenchida
            If objCotProduto.lFornecedor > 0 Then
    
                objFornecedor.lCodigo = objCotProduto.lFornecedor
    
                'Lê o Fornecedor a partir do código do Fornecedor
                lErro = CF("Fornecedor_Le", objFornecedor)
                If lErro <> SUCESSO And lErro <> 12729 Then gError 63496
    
                'Se não encontrou o Fornecedor ==> erro
                If lErro = 12729 Then gError 63500
    
                'Coloca código e nome reduzido do fornecedor no grid
                GridProdutos.TextMatrix(iLinha, iGrid_FornecedorProd_Col) = objFornecedor.sNomeReduzido
    
            End If
    
            'Verifica se a Filial está preenchida
            If objCotProduto.iFilial > 0 Then
    
                objFilialForn.iCodFilial = objCotProduto.iFilial
                objFilialForn.lCodFornecedor = objCotProduto.lFornecedor
    
                'Lê a FilialFornecedor a partir do código da Filial
                lErro = CF("FilialFornecedor_Le", objFilialForn)
                If lErro <> SUCESSO And lErro <> 12929 Then gError 63497
    
                'Se não encontrou a Filial==> erro
                If lErro = 12929 Then gError 63499
    
                'Coloca código e nome da FilialFornecedor no grid
                GridProdutos.TextMatrix(iLinha, iGrid_FilialFornProd_Col) = objFilialForn.iCodFilial & SEPARADOR & objFilialForn.sNome
    
            End If
            
        End If

    Next

    'Atualiza o número de Linhas existentes do grid
    objGridProdutos.iLinhasExistentes = iLinha

    Call Grid_Refresh_Checkbox(objGridProdutos)
    
    GridProduto_Preenche = SUCESSO

    Exit Function

Erro_GridProduto_Preenche:

    Select Case gErr

        Case 63495, 63496, 63497, 63583
            'Erros tratados nas rotinas chamadas

        Case 63498
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case 63499
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objFilialForn.iCodFilial, objFilialForn.lCodFornecedor)

        Case 63500
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161397)

    End Select

    Exit Function

End Function

Function Move_GridFornecedores_Memoria(objCotacao As ClassCotacao, colPedidoCotacao As Collection) As Long
'Recolhe os dados do GridFornecedores

Dim lErro As Long
Dim iIndice As Integer
Dim iLinha As Integer
Dim iLinha2 As Integer
Dim objPedidoCotacao As ClassPedidoCotacao
Dim objCotacaoProduto As ClassCotacaoProduto
Dim objItemPedCotacao As ClassItemPedCotacao
Dim objFornecedor As New ClassFornecedor
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Move_GridFornecedores_Memoria

    'Inicializa a coleção de Pedidos de Cotação
    Set colPedidoCotacao = New Collection
    
    'Para cada linha do Grid de Fornecedores
    For iLinha = 1 To objGridFornecedores.iLinhasExistentes
                
        'Se a linha foi marcada
        If GridFornecedores.TextMatrix(iLinha, iGrid_EscolhidoForn_Col) = GRID_CHECKBOX_ATIVO Then
        
            'Lê o Fornecedor do Grid de Fornecedores
            objFornecedor.sNomeReduzido = GridFornecedores.TextMatrix(iLinha, iGrid_FornecedorGrid_Col)
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then gError 66782
            
            'Se não encontrou o Fornecedor, erro
            If lErro = 6681 Then gError 66783
            
            'Verifica se já existe um Pedido de Cotação para o Fornecedor e Filial da linha do Grid de Fornecedores em questão
            For iIndice = 1 To colPedidoCotacao.Count
                        
                'Se encontrou um Pedido de cotação com o mesmo Fornecedor e Filial
                If colPedidoCotacao(iIndice).lFornecedor = objFornecedor.lCodigo And colPedidoCotacao(iIndice).iFilial = Codigo_Extrai(GridFornecedores.TextMatrix(iLinha, iGrid_FilialFornGrid_Col)) Then
                    Exit For
                End If
            
            Next
                
            'Se não encontrou o Pedido de Cotação
            If iIndice > colPedidoCotacao.Count Then
            
                'Cria um novo Pedido de Cotação
                Set objPedidoCotacao = New ClassPedidoCotacao
                
                objPedidoCotacao.lFornecedor = objFornecedor.lCodigo
                objPedidoCotacao.iFilial = Codigo_Extrai(GridFornecedores.TextMatrix(iLinha, iGrid_FilialFornGrid_Col))
                                                                                                
                'Tipo de Frete
                If GridFornecedores.TextMatrix(iLinha, iGrid_Frete_Col) = "CIF" Then
                    objPedidoCotacao.iTipoFrete = TIPO_CIF
                ElseIf GridFornecedores.TextMatrix(iLinha, iGrid_Frete_Col) = "FOB" Then
                    objPedidoCotacao.iTipoFrete = TIPO_FOB
                End If
                
                'Se existe Condicao de Pagamento à prazo selecionada
                If objCotacao.colCondPagtos.Count > 1 Then
                    objPedidoCotacao.iCondPagtoPrazo = objCotacao.colCondPagtos.Item(2)
                End If
                
                objPedidoCotacao.iStatus = STATUS_GERADO_NAO_ATUALIZADO
                objPedidoCotacao.iFilialEmpresa = giFilialEmpresa
                objPedidoCotacao.dtData = gdtDataHoje
                objPedidoCotacao.dtDataEmissao = DATA_NULA
                objPedidoCotacao.dtDataValidade = DATA_NULA
                                              
                'Procura no Grid de Fornecedores as linhas que possuem o mesmo Fornecedor e Filial do Pedido de Cotação que acaba de ser criado
                For iLinha2 = 1 To objGridFornecedores.iLinhasExistentes
                    
                    'Se encontrou e a linha está marcada
                    If GridFornecedores.TextMatrix(iLinha2, iGrid_FornecedorGrid_Col) = objFornecedor.sNomeReduzido And Codigo_Extrai(GridFornecedores.TextMatrix(iLinha2, iGrid_FilialFornGrid_Col)) = objPedidoCotacao.iFilial And GridFornecedores.TextMatrix(iLinha2, iGrid_EscolhidoForn_Col) = GRID_CHECKBOX_ATIVO Then
                
                        'Formata o Produto da linha
                        lErro = CF("Produto_Formata", GridFornecedores.TextMatrix(iLinha2, iGrid_ProdutoForn_Col), sProdutoFormatado, iProdutoPreenchido)
                        If lErro <> SUCESSO Then gError 66793
                
                        'Lê o Fornecedor do Grid de Fornecedores
                        objFornecedor.sNomeReduzido = GridFornecedores.TextMatrix(iLinha2, iGrid_FornecedorGrid_Col)
                        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                        If lErro <> SUCESSO And lErro <> 6681 Then gError 70506
                        
                        'Se não encontrou o Fornecedor, erro
                        If lErro = 6681 Then gError 70507
                
                        If GridFornecedores.TextMatrix(iLinha2, iGrid_Exclusivo_Col) = "1" Then
                        
                            'Procura o Produto na coleção de Cotação Produto
                            For Each objCotacaoProduto In objCotacao.colCotacaoProduto
                                                            
                                'Se encontrou o mesmo Produto, Fornecedor, Filial
                                If sProdutoFormatado = objCotacaoProduto.sProduto And objFornecedor.lCodigo = objCotacaoProduto.lFornecedor And Codigo_Extrai(GridFornecedores.TextMatrix(iLinha2, iGrid_FilialFornGrid_Col)) = objCotacaoProduto.iFilial Then
                                                                    
                                    'Cria Item de Pedido de Cotação
                                    Set objItemPedCotacao = New ClassItemPedCotacao
                                    
                                    objItemPedCotacao.lCotacaoProduto = objCotacaoProduto.lNumIntDoc
                                    objItemPedCotacao.sProduto = sProdutoFormatado
                                    objItemPedCotacao.dQuantidade = objCotacaoProduto.dQuantidade
                                    objItemPedCotacao.sUM = objCotacaoProduto.sUM
                                    objItemPedCotacao.lCotacaoProduto = objCotacaoProduto.lNumIntDoc
                                
                                    'Adiciona o Item na coleção de Pedido de Cotação
                                    objPedidoCotacao.colItens.Add objItemPedCotacao
                                    
                                    Exit For
                                                                                                
                                End If
                                
                            Next
                        
                        'Fornecedor não é exclusivo
                        Else
                        
                            'Procura o Produto na coleção de Cotação Produto
                            For Each objCotacaoProduto In objCotacao.colCotacaoProduto
                                                            
                                'Se encontrou o mesmo Produto e não tem identificação de Fornecedor
                                If sProdutoFormatado = objCotacaoProduto.sProduto And objCotacaoProduto.lFornecedor = 0 Then
                                                                    
                                    'Cria Item de Pedido de Cotação
                                    Set objItemPedCotacao = New ClassItemPedCotacao
                                    
                                    objItemPedCotacao.lCotacaoProduto = objCotacaoProduto.lNumIntDoc
                                    objItemPedCotacao.sProduto = sProdutoFormatado
                                    objItemPedCotacao.dQuantidade = objCotacaoProduto.dQuantidade
                                    objItemPedCotacao.sUM = objCotacaoProduto.sUM
                                    objItemPedCotacao.lCotacaoProduto = objCotacaoProduto.lNumIntDoc
                                
                                    'Adiciona o Item na coleção de Pedido de Cotação
                                    objPedidoCotacao.colItens.Add objItemPedCotacao
                                    
                                    Exit For
                                                                                                
                                End If
                                
                            Next
                        
                        End If
                    
                    End If
                    
                Next
                
                'Adiciona na coleção de Pedidos de cotação
                colPedidoCotacao.Add objPedidoCotacao
            End If
        End If
    Next
                                    
    Move_GridFornecedores_Memoria = SUCESSO

    Exit Function

Erro_Move_GridFornecedores_Memoria:

    Move_GridFornecedores_Memoria = gErr

    Select Case gErr

        Case 66782, 66793, 70506
                
        Case 66783, 70507
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161398)

    End Select

    Exit Function

End Function

Function Move_GridProdutos_Memoria(objCotacao As ClassCotacao) As Long
'Recolhe os dados do GridProdutos e guarda-os em colCotacaoProduto

Dim lErro As Long
Dim objCotProduto As ClassCotacaoProduto
Dim iIndice As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objFornecedor As New ClassFornecedor
Dim lNumIntProvisorio As Long

On Error GoTo Erro_Move_GridProdutos_Memoria

    lNumIntProvisorio = 1
    
    'Para cada linha do Grid de Produtos
    For iIndice = 1 To objGridProdutos.iLinhasExistentes

        'Verifica se a linha foi selecionada
        If GridProdutos.TextMatrix(iIndice, iGrid_EscolhidoProd_Col) = GRID_CHECKBOX_ATIVO Then

            Set objCotProduto = New ClassCotacaoProduto
            
            'Preenche objCotProduto com os dados do GridProdutos
            objCotProduto.dQuantidade = StrParaDbl(GridProdutos.TextMatrix(iIndice, iGrid_QuantidadeProd_Col))
            objCotProduto.iFilial = Codigo_Extrai(GridProdutos.TextMatrix(iIndice, iGrid_FilialFornProd_Col))
            objFornecedor.sNomeReduzido = GridProdutos.TextMatrix(iIndice, iGrid_FornecedorProd_Col)
            
            If Len(Trim(objFornecedor.sNomeReduzido)) > 0 Then
            
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO And lErro <> 6681 Then gError 68347
                
                If lErro = 6681 Then gError 70518
                
                objCotProduto.lFornecedor = objFornecedor.lCodigo
                    
            Else
                objCotProduto.lFornecedor = 0
            End If
            

            'Verifica se o produto está preenchido
            If Len(Trim(GridProdutos.TextMatrix(iIndice, iGrid_Produto_Col))) > 0 Then

                'Coloca o produto no formato do BD
                lErro = CF("Produto_Formata", GridProdutos.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
                If lErro <> SUCESSO Then gError 63586

                objCotProduto.sProduto = sProdutoFormatado

            End If

            objCotProduto.sUM = GridProdutos.TextMatrix(iIndice, iGrid_UnidadeMed_Col)

            'Coloca número provisório em NumIntDoc para linkar com ItensPedCOtacao de PedidosCotacao
            objCotProduto.lNumIntDoc = lNumIntProvisorio
                        
            'Adiciona em colCotacaoProduto
            objCotacao.colCotacaoProduto.Add objCotProduto
            
            'Incrementa o número provisório
            lNumIntProvisorio = lNumIntProvisorio + 1

        End If
    Next

    Move_GridProdutos_Memoria = SUCESSO

    Exit Function

Erro_Move_GridProdutos_Memoria:

    Move_GridProdutos_Memoria = gErr

    Select Case gErr

        Case 63586, 68347
            'Erro tratado na rotina chamada

        Case 70518
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161399)

    End Select

    Exit Function

End Function

Function Preenche_GridFornecedores(colCotacaoProdutoFornecedor As Collection) As Long
'Preenche o grid de fornecedores

Dim lErro As Long
Dim objCotProdutoFornecedor As New ClassCotacaoProdutoForn
Dim iLinha As Integer
Dim iLinha2 As Integer
Dim objProduto As New ClassProduto
Dim objFornecedor As New ClassFornecedor
Dim objFilFornecedor As New ClassFilialFornecedor
Dim sProdutoMascarado As String
Dim objFornecedorEstatistica As New ClassFilialFornecedorEst
Dim objCondPagto As New ClassCondicaoPagto
Dim iIndice As Integer
Dim sProdFormatado As String
Dim iProdPreenchido As Integer

On Error GoTo Erro_Preenche_GridFornecedores

    'Limpa o Grid de Fornecedores
    Call Grid_Limpa(objGridFornecedores)

    Set objCotProdutoFornecedor.objCotacaoProduto = New ClassCotacaoProduto
    Set objCotProdutoFornecedor.objFornecedorProdutoFF = New ClassFornecedorProdutoFF

    For iLinha = colCotacaoProdutoFornecedor.Count To 1 Step -1
        
        Set objCotProdutoFornecedor = colCotacaoProdutoFornecedor.Item(iLinha)
        
        iLinha2 = iLinha2 + 1
        
        'Mascara o Produto
        lErro = Mascara_RetornaProdutoEnxuto(objCotProdutoFornecedor.objFornecedorProdutoFF.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 63582

        ProdutoForn.PromptInclude = False
        ProdutoForn.Text = sProdutoMascarado
        ProdutoForn.PromptInclude = True

        'Coloca o produto mascarado no grid
        GridFornecedores.TextMatrix(iLinha2, iGrid_ProdutoForn_Col) = ProdutoForn.Text
        
        objProduto.sCodigo = objCotProdutoFornecedor.objFornecedorProdutoFF.sProduto

        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 63543

        'Se não encontrou o Produto ==> erro
        If lErro = 28030 Then gError 63551

        'Preenche a Descricao do produto
        GridFornecedores.TextMatrix(iLinha2, iGrid_DescProdutoForn_Col) = objProduto.sDescricao

        'Verifica se o Fornecedor está preenchido
        If objCotProdutoFornecedor.objCotacaoProduto.lFornecedor > 0 Then

            objFornecedor.lCodigo = objCotProdutoFornecedor.objCotacaoProduto.lFornecedor

            'Lê  o Fornecedor
            lErro = CF("Fornecedor_Le", objFornecedor)
            If lErro <> SUCESSO And lErro <> 12729 Then gError 63544

            'Se não encontrou o Fornecedor ==> erro
            If lErro = 12729 Then gError 63549
            
            'Coloca o nome reduzido do Fornecedor no Grid
            GridFornecedores.TextMatrix(iLinha2, iGrid_FornecedorGrid_Col) = objFornecedor.sNomeReduzido
            
        End If

        'Verifica se a Filial está preeenchida
        If objCotProdutoFornecedor.objCotacaoProduto.iFilial > 0 Then

            objFilFornecedor.iCodFilial = objCotProdutoFornecedor.objCotacaoProduto.iFilial
            objFilFornecedor.lCodFornecedor = objCotProdutoFornecedor.objCotacaoProduto.lFornecedor
            'Lê a Filial do Fornecedor
            lErro = CF("FilialFornecedor_Le", objFilFornecedor)
            If lErro <> SUCESSO And lErro <> 12929 Then gError 63545

            'Se não encontrou a Filial do Fornecedor ==> erro
            If lErro = 12929 Then gError 63547

            'Preenche o grid com código e nome da Filial do Fornecedor
            GridFornecedores.TextMatrix(iLinha2, iGrid_FilialFornGrid_Col) = objCotProdutoFornecedor.objCotacaoProduto.iFilial & SEPARADOR & objFilFornecedor.sNome
            
        End If

        'Verifica se DataUltimaCotacao é diferente de Data nula
        If objCotProdutoFornecedor.objFornecedorProdutoFF.dtDataUltimaCotacao <> DATA_NULA Then

            GridFornecedores.TextMatrix(iLinha2, iGrid_UltimaCotacao_Col) = Format(objCotProdutoFornecedor.objFornecedorProdutoFF.dtDataUltimaCotacao, "dd/mm/yy")

        End If
        GridFornecedores.TextMatrix(iLinha2, iGrid_ValorCotacao_Col) = Formata_Estoque(objCotProdutoFornecedor.objFornecedorProdutoFF.dQuantUltimaCotacao)
        
        If objCotProdutoFornecedor.objFornecedorProdutoFF.iTipoFreteUltimaCotacao = TIPO_FOB Then
        
            GridFornecedores.TextMatrix(iLinha2, iGrid_Frete_Col) = "FOB"
            
        ElseIf objCotProdutoFornecedor.objFornecedorProdutoFF.iTipoFreteUltimaCotacao = TIPO_CIF Then
        
            GridFornecedores.TextMatrix(iLinha2, iGrid_Frete_Col) = "CIF"
            
        End If

        If objCotProdutoFornecedor.objFornecedorProdutoFF.dtDataUltimaCompra - objCotProdutoFornecedor.objFornecedorProdutoFF.dtDataReceb <> 0 Then
            GridFornecedores.TextMatrix(iLinha2, iGrid_PrazoEntrega_Col) = objCotProdutoFornecedor.objFornecedorProdutoFF.dtDataUltimaCompra - objCotProdutoFornecedor.objFornecedorProdutoFF.dtDataReceb
        End If
        
        GridFornecedores.TextMatrix(iLinha2, iGrid_QuantPedidaForn_Col) = Format(objCotProdutoFornecedor.objFornecedorProdutoFF.dQuantPedida)
        GridFornecedores.TextMatrix(iLinha2, iGrid_QuantRecebidaForn_Col) = Format(objCotProdutoFornecedor.objFornecedorProdutoFF.dQuantRecebida)


        objFornecedor.lCodigo = objCotProdutoFornecedor.objCotacaoProduto.lFornecedor

        'Lê os dados de Estatística do Fornecedor
        lErro = CF("Fornecedor_Le_Estendida", objFornecedor, objFornecedorEstatistica)
        If lErro <> SUCESSO And lErro <> 52701 Then gError 63601

        'Se não encontrou o Fornecedor ==> erro
        If lErro = 52701 Then gError 63602

        GridFornecedores.TextMatrix(iLinha2, iGrid_SaldoTitulos_Col) = objFornecedorEstatistica.dSaldoTitulos
        GridFornecedores.TextMatrix(iLinha2, iGrid_ObservacaoForn_Col) = objFornecedor.sObservacao
        
        'Verifica se existe CondicaoPagto para o Fornecedor
        If objFornecedor.iCondicaoPagto <> 0 Then
        
            objCondPagto.iCodigo = objFornecedor.iCondicaoPagto
            
            lErro = CF("CondicaoPagto_Le", objCondPagto)
            If lErro <> SUCESSO And lErro <> 19205 Then gError 68342
            If lErro <> SUCESSO Then gError 62863
            
            GridFornecedores.TextMatrix(iLinha2, iGrid_CondicaoPagto_Col) = objFornecedor.iCondicaoPagto & SEPARADOR & objCondPagto.sDescReduzida
        End If
        
        'Verifica se DataUltimaCompra é diferente de Data Nula
        If objFornecedorEstatistica.dtDataUltimaCompra <> DATA_NULA Then

            GridFornecedores.TextMatrix(iLinha2, iGrid_UltimaCompra_Col) = Format(objFornecedorEstatistica.dtDataUltimaCompra, "dd/mm/yy")

        End If
        
        If objCotProdutoFornecedor.objCotacaoProduto.lFornecedor > 0 Then
            GridFornecedores.TextMatrix(iLinha2, iGrid_Exclusivo_Col) = MARCADO
        Else
            GridFornecedores.TextMatrix(iLinha2, iGrid_Exclusivo_Col) = DESMARCADO
        End If
        
        objGridFornecedores.iLinhasExistentes = iLinha2
        
    Next

    Call Grid_Refresh_Checkbox(objGridFornecedores)

    Preenche_GridFornecedores = SUCESSO

    Exit Function

Erro_Preenche_GridFornecedores:

    Preenche_GridFornecedores = gErr

    Select Case gErr
        
        Case 62863
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondPagto.iCodigo)

        Case 63543, 63544, 63545, 63582, 63601, 68342, 72343, 72344
            'Erros tratados nas rotinas chamadas

        Case 63547
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objFilFornecedor.iCodFilial, objFornecedor.lCodigo)

        Case 63549
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)

        Case 63551
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case 63602
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)

        Case 70519
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161400)

    End Select

    Exit Function
    
End Function

Function Move_Itens_Produtos(objGeracaoCotacao As ClassGeracaoCotacao, colCotProduto As Collection) As Long
'Move em memória Itens de objGeracaoCotacao para colCotProduto

Dim lErro As Long
Dim objReqCompras As ClassRequisicaoCompras
Dim objItemReqCompras As ClassItemReqCompras
Dim objCotProduto As ClassCotacaoProduto
Dim objCotProdutoPesq As ClassCotacaoProduto
Dim iEncontrou As Integer
Dim objProduto As New ClassProduto
Dim dFator As Double

On Error GoTo Erro_Move_Itens_Produtos

    Set colCotProduto = New Collection

    For Each objReqCompras In objGeracaoCotacao.colReqCompra

        For Each objItemReqCompras In objReqCompras.colItens
    
            Set objCotProduto = New ClassCotacaoProduto

            objCotProduto.iSelecionado = MARCADO
            objCotProduto.iEscolhido = Selecionado

            'Recolhe o Produto de objItemReqCompras
            objProduto.sCodigo = objItemReqCompras.sProduto

            'Lê os dados do produto envolvido
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 62706
            If lErro <> SUCESSO Then gError 62707

            objCotProduto.sProduto = objProduto.sCodigo

            'Recolhe a UM de objItemReqCompras
            objCotProduto.sUM = objItemReqCompras.sUM
                
            'Recolhe a Quantidade a cotar de objItemReqCompras
            objCotProduto.dQuantidade = objItemReqCompras.dQuantCotar
            
            'Define selecao
            If objCotProduto.dQuantidade > 0 Then
                objCotProduto.iSelecionado = Selecionado
            Else
                objCotProduto.iSelecionado = NAO_SELECIONADO
            End If
            
            'Converte para a Unidade de Medida de Compras
            lErro = CF("UM_Conversao", objProduto.iClasseUM, objCotProduto.sUM, objProduto.sSiglaUMCompra, dFator)
            If lErro <> SUCESSO Then gError 62708
                                     
            objCotProduto.sUM = objProduto.sSiglaUMCompra
            objCotProduto.dQuantidade = objCotProduto.dQuantidade * dFator

            'Se for um item com fornecedor Exclusivo então
            If objItemReqCompras.iExclusivo = FORNECEDOR_EXCLUSIVO Then

                objCotProduto.lFornecedor = objItemReqCompras.lFornecedor
                objCotProduto.iFilial = objItemReqCompras.iFilial
            
            Else

                objCotProduto.lFornecedor = 0
                objCotProduto.iFilial = 0
            
            End If
            
            iEncontrou = PRODUTO_NAO_ENCONTRADO
            
            'Verifica se está na coleção com o Trio igual(Produto, Fornecedor, Filial)
            For Each objCotProdutoPesq In colCotProduto
                If (objCotProdutoPesq.sProduto = objCotProduto.sProduto) And (objCotProdutoPesq.lFornecedor = objCotProduto.lFornecedor) And (objCotProdutoPesq.iFilial = objCotProduto.iFilial) Then
                    objCotProdutoPesq.dQuantidade = objCotProdutoPesq.dQuantidade + objCotProduto.dQuantidade
                    
                    If objCotProdutoPesq.dQuantidade > 0 Then
                        objCotProdutoPesq.iSelecionado = Selecionado
                    End If
                    
                    objCotProdutoPesq.colItemReqCompras.Add objItemReqCompras
                    iEncontrou = PRODUTO_ENCONTRADO
                    Exit For
                End If
            Next
            
            If iEncontrou = PRODUTO_NAO_ENCONTRADO Then
                'Coloca um NumIntDoc provisório que corresponde ao índice do ítem
                objCotProduto.lNumIntDoc = colCotProduto.Count + 1
                objCotProduto.colItemReqCompras.Add objItemReqCompras
                colCotProduto.Add objCotProduto
            End If
            
        Next

    Next

    Move_Itens_Produtos = SUCESSO

    Exit Function

Erro_Move_Itens_Produtos:

    Move_Itens_Produtos = gErr

    Select Case gErr

        Case 62707
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            
        Case 62706, 62708
            'Erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161401)

    End Select

    Exit Function

End Function

Function Move_ItemReq_Memoria(objGeracaoCotacao As ClassGeracaoCotacao, colCotProduto As Collection) As Long
'Move para memória Item de Grid de Itens que foi selecionado/desselecionado ou teve quantidade alterada

Dim lErro As Long
Dim objReqCompras As ClassRequisicaoCompras
Dim objItemReqCompras As ClassItemReqCompras
Dim objCotProduto As New ClassCotacaoProduto
Dim objCotProdutoPesq As New ClassCotacaoProduto
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iIndice As Integer
Dim iEncontrou As Integer
Dim objFornecedor As New ClassFornecedor
Dim objProduto As New ClassProduto
Dim dFator As Double

On Error GoTo Erro_Move_ItemReq_Memoria

    Set colCotProduto = New Collection

    For Each objReqCompras In objGeracaoCotacao
        
        If objReqCompras.iSelecionado = Selecionado Then
    
            For Each objItemReqCompras In objReqCompras.colItens
                
                If objItemReqCompras.iSelecionado = Selecionado Then
    
                    Set objCotProduto = New ClassCotacaoProduto
        
                    'Recolhe o Produto do Grid
                    sProduto = GridItensRequisicoes.TextMatrix(iIndice, iGrid_ProdutoItem_Col)
        
                    'Critica o formato do Produto
                    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
                    If lErro <> SUCESSO Then gError 63493
        
                    objProduto.sCodigo = sProdutoFormatado
                    'Lê os dados do produto envolvido
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 62706
                    If lErro <> SUCESSO Then gError 62707
        
                    objCotProduto.sProduto = sProdutoFormatado
        
                    'Recolhe a UM do grid
                    objCotProduto.sUM = GridItensRequisicoes.TextMatrix(iIndice, iGrid_UnidadeMedItem_Col)
                        
                    'Recolhe a Quantidade do Grid
                    objCotProduto.dQuantidade = StrParaDbl(GridItensRequisicoes.TextMatrix(iIndice, iGrid_Quantidade_Col))
                    
                    'Converte para a Unidade de Medida de Compras
                    lErro = CF("UM_Conversao", objProduto.iClasseUM, objCotProduto.sUM, objProduto.sSiglaUMCompra, dFator)
                    If lErro <> SUCESSO Then gError 62708
                                             
                    objCotProduto.sUM = objProduto.sSiglaUMCompra
                    objCotProduto.dQuantidade = objCotProduto.dQuantidade * dFator
        
                    'Se for um item com fornecedor Exclusivo então
                    If GridItensRequisicoes.TextMatrix(iIndice, iGrid_ExclusoItem_Col) = "Exclusivo" Then
        
                        objFornecedor.sNomeReduzido = GridItensRequisicoes.TextMatrix(iIndice, iGrid_FornecedorItem_Col)
                        
                        If Len(Trim(objFornecedor.sNomeReduzido)) > 0 Then
                            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                            If lErro <> SUCESSO And lErro <> 6681 Then gError 68343
                            
                            If lErro = 6681 Then gError 70520
                        
                        End If
                        
                        objCotProduto.lFornecedor = objFornecedor.lCodigo
                        objCotProduto.iFilial = Codigo_Extrai(GridItensRequisicoes.TextMatrix(iIndice, iGrid_FilialFornItem_Col))
        
                        iEncontrou = PRODUTO_NAO_ENCONTRADO
                        
                        'Verifica se esta na colecao com o Trio igual(Produto, Fornecedor, Filial)
                        For Each objCotProdutoPesq In colCotProduto
                            If (objCotProdutoPesq.sProduto = objCotProduto.sProduto) And (objCotProdutoPesq.lFornecedor = objCotProduto.lFornecedor) And (objCotProdutoPesq.iFilial = objCotProduto.iFilial) Then
                                objCotProdutoPesq.dQuantidade = objCotProdutoPesq.dQuantidade + objCotProduto.dQuantidade
                                iEncontrou = PRODUTO_ENCONTRADO
                                Exit For
                            End If
                        Next
                    Else
        
                        objCotProduto.lFornecedor = 0
                        objCotProduto.iFilial = 0
                        iEncontrou = PRODUTO_NAO_ENCONTRADO
                        
                        'Se não for exclusivo então Verifica se o Produto está na colecao
                        For Each objCotProdutoPesq In colCotProduto
                            If (objCotProdutoPesq.sProduto = objCotProduto.sProduto) And (objCotProdutoPesq.lFornecedor = 0) And (objCotProdutoPesq.iFilial = 0) Then
                                               
                            
                                objCotProdutoPesq.dQuantidade = objCotProdutoPesq.dQuantidade + objCotProduto.dQuantidade
                                iEncontrou = PRODUTO_ENCONTRADO
                                Exit For
                            End If
                        Next
                    End If
        
                    If iEncontrou = PRODUTO_NAO_ENCONTRADO Then
                        colCotProduto.Add objCotProduto
                    End If

                End If
            
            Next

        End If

    Next

    Move_ItemReq_Memoria = SUCESSO

    Exit Function

Erro_Move_ItemReq_Memoria:

    Move_ItemReq_Memoria = gErr

    Select Case gErr

        Case 62707
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            
        Case 62706, 62707, 63493, 68343
            'Erro tratado na rotina chamada

        
        Case 70520
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161402)

    End Select

    Exit Function

End Function

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Geração de Pedidos de Cotação"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "GeracaoPedCotacao"

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

Private Sub TipoDestino_Click(Index As Integer)

    giTabSelecao_Alterado = REGISTRO_ALTERADO
    
    'Torna invisivel o FrameDestino com índice igual a iFrameDestinoAtual
    FrameDestino(giFrameDestinoAtual).Visible = False

    'Torna visível o FrameDestino com índice igual a Index
    FrameDestino(Index).Visible = True

    'Armazena novo valor de giFrameDestinoAtual
    giFrameDestinoAtual = Index

End Sub

Private Sub TipoFrete_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub
Private Sub TipoFrete_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub TipoFrete_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub TipoFrete_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = TipoFrete
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub TipoProduto_Click()

    giTabSelecao_Alterado = REGISTRO_ALTERADO

End Sub

Private Sub UltimaCotacao_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UltimaCotacao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFornecedores)

End Sub

Private Sub UltimaCotacao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFornecedores)

End Sub

Private Sub UltimaCotacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFornecedores.objControle = UltimaCotacao
    lErro = Grid_Campo_Libera_Foco(objGridFornecedores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UnidadeMedItem_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UnidadeMedItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensRequisicoes)

End Sub

Private Sub UnidadeMedItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensRequisicoes)

End Sub

Private Sub UnidadeMedItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensRequisicoes.objControle = UnidadeMedItem
    lErro = Grid_Campo_Libera_Foco(objGridItensRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UnidadeMedProd_Change()

    giAlterado = REGISTRO_ALTERADO

End Sub
Private Sub UnidadeMedProd_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos)

End Sub

Private Sub UnidadeMedProd_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)

End Sub

Private Sub UnidadeMedProd_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = UnidadeMedProd
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 63457

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 63457
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161403)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Aumenta um dia em DataAte
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 63458

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 63458
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161404)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 63456

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 63456
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161405)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Aumenta um dia em DataDe
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 63459

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 63459
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161406)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataLimAte_DownClick()
Dim lErro As Long

On Error GoTo Erro_UpDownDataLimAte_DownClick

    'Diminui um dia em DataLimiteDe
    lErro = Data_Up_Down_Click(DataLimiteAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 63455

    Exit Sub

Erro_UpDownDataLimAte_DownClick:

    Select Case gErr

        Case 63455
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161407)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataLimAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataLimAte_UpClick

    'Aumenta um dia em DataLimiteAte
    lErro = Data_Up_Down_Click(DataLimiteAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 63460

    Exit Sub

Erro_UpDownDataLimAte_UpClick:

    Select Case gErr

        Case 63460
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161408)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataLimDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataLimDe_DownClick

    'Diminui um dia em DataLimiteDe
    lErro = Data_Up_Down_Click(DataLimiteDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 63454

    Exit Sub

Erro_UpDownDataLimDe_DownClick:

    Select Case gErr

        Case 63454
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161409)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataLimDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataLimDe_UpClick

    'Aumenta um dia em DataLimiteDe
    lErro = Data_Up_Down_Click(DataLimiteDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 63461

    Exit Sub

Erro_UpDownDataLimDe_UpClick:

    Select Case gErr

        Case 63461
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161410)

    End Select

    Exit Sub

End Sub

Private Sub Urgente_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRequisicoes)

End Sub

Private Sub Urgente_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRequisicoes)

End Sub

Private Sub Urgente_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRequisicoes.objControle = Urgente
    lErro = Grid_Campo_Libera_Foco(objGridRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

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
    m_Caption = New_Caption
End Property

'**** fim do trecho a ser copiado *****

Private Sub FilialFornLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialFornLabel, Source, X, Y)
End Sub

Private Sub FilialFornLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialFornLabel, Button, Shift, X, Y)
End Sub

Private Sub FilialEmpresaLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialEmpresaLabel, Source, X, Y)
End Sub

Private Sub FilialEmpresaLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialEmpresaLabel, Button, Shift, X, Y)
End Sub

Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub Label17_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label17, Source, X, Y)
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label17, Button, Shift, X, Y)
End Sub


Private Sub FornecedorLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecedorLabel, Source, X, Y)
End Sub

Private Sub FornecedorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel, Button, Shift, X, Y)
End Sub

Private Sub Comprador_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Comprador, Source, X, Y)
End Sub

Private Sub Comprador_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Comprador, Button, Shift, X, Y)
End Sub

Private Sub Label28_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label28, Source, X, Y)
End Sub

Private Sub Label28_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label28, Button, Shift, X, Y)
End Sub

Private Sub Label37_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label37, Source, X, Y)
End Sub

Private Sub Label37_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label37, Button, Shift, X, Y)
End Sub

Private Sub CondPagto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CondPagto, Source, X, Y)
End Sub

Private Sub CondPagto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CondPagto, Button, Shift, X, Y)
End Sub

Private Sub Label55_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label55, Source, X, Y)
End Sub

Private Sub Label55_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label55, Button, Shift, X, Y)
End Sub

Private Sub Label54_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label54, Source, X, Y)
End Sub

Private Sub Label54_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label54, Button, Shift, X, Y)
End Sub

Private Sub Label31_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label31, Source, X, Y)
End Sub

Private Sub Label31_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label31, Button, Shift, X, Y)
End Sub

Private Sub Label40_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label40, Source, X, Y)
End Sub

Private Sub Label40_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label40, Button, Shift, X, Y)
End Sub

Sub Monta_Colecao_Campos_Requisicao(colCampos As Collection, iOrdenacao As Integer)

    Select Case iOrdenacao

        Case 0

            colCampos.Add "dtDataLimite"
            colCampos.Add "lUrgente"
            colCampos.Add "lCodigo"
            colCampos.Add "iFilialEmpresa"

        Case 1

            colCampos.Add "lUrgente"
            colCampos.Add "dtDataLimite"
            colCampos.Add "lCodigo"
            colCampos.Add "iFilialEmpresa"

        Case 2

            colCampos.Add "dtData"
            colCampos.Add "dtDataLimite"
            colCampos.Add "lCodigo"
            colCampos.Add "iFilialEmpresa"

        Case 3

            colCampos.Add "lCodigo"
            colCampos.Add "iFilialEmpresa"

        Case 4

            colCampos.Add "sCcl"
            colCampos.Add "dtDataLimite"
            colCampos.Add "lCodigo"
            colCampos.Add "iFilialEmpresa"

    End Select
    
End Sub

Sub Monta_Colecao_Campos_Fornecedor(colCampos As Collection, iOrdenacao As Integer)

    Select Case iOrdenacao

        Case 0
            colCampos.Add "sProduto"
            colCampos.Add "sFornecedor"
            colCampos.Add "sFilialForn"
            colCampos.Add "sExclusivo"

        Case 1
            colCampos.Add "sFornecedor"
            colCampos.Add "sFilialForn"
            colCampos.Add "sProduto"
            colCampos.Add "sExclusivo"

    End Select
    
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label45_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label45, Source, X, Y)
End Sub

Private Sub Label45_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label45, Button, Shift, X, Y)
End Sub

Private Sub Label57_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label57, Source, X, Y)
End Sub

Private Sub Label57_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label57, Button, Shift, X, Y)
End Sub


Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim lErro As Long
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Rotina_Grid_Enable
   
    'Só executa para entrada de célula
    'Pesquisa controle da coluna em questão
    Select Case objControl.Name
        
        'QuantCotarItem
        Case QuantCotarItem.Name
            If iLinha = 0 Then
                objControl.Enabled = False
                Exit Sub
            End If
        
            'Verifica se a linha está selecionada está preenchido
            If StrParaInt(GridItensRequisicoes.TextMatrix(iLinha, iGrid_EscolhidoItem_Col)) = 0 Then
                objControl.Enabled = False
                Exit Sub
            Else
                objControl.Enabled = True
            End If
        Case QuantidadeProd.Name
            
            If iLinha = 0 Then
                objControl.Enabled = False
                Exit Sub
            End If
            'Verifica se a linha está selecionada está preenchido
            If giPodeAumentarQuant = MARCADO Then
                If StrParaInt(GridProdutos.TextMatrix(iLinha, iGrid_EscolhidoProd_Col)) = 0 Then
                    objControl.Enabled = False
                    Exit Sub
                Else
                    objControl.Enabled = True
                End If
            Else
                objControl.Enabled = False
            End If
                    
    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161411)

    End Select

    Exit Sub

End Sub

Private Sub BotaoDesmarcarTodos_Click(Index As Integer)
'Desmarca os itens de acordo com o índice determinado

    'De acordo com o índice do tab que está visível, chama a função específica
    Select Case Index

        Case 1

            'Desmarca todos os tipos de Produto da list TipoProduto
            Call DesmarcaTodos_TipoProduto

            giTabSelecao_Alterado = REGISTRO_ALTERADO

        Case 2

            'Desmarca todas as Requisicoes do GridRequisicoes
            Call DesmarcaTodas_Requisicoes

            giTabRequisicao_Alterado = REGISTRO_ALTERADO

        Case 3

            'Desmarca todos os itens do GridItensRequisicoes
            Call DesmarcaTodos_ItensRequisicoes

            giTabItens_Alterado = REGISTRO_ALTERADO

        Case 4

            'Desmarca todos os Produtos do GridProdutos
            Call DesmarcaTodos_Produtos

            giTabProdutos_Alterado = REGISTRO_ALTERADO

        Case 5

            'Desmarca todos os Fornecedores do GridFornecedores
            Call DesmarcaTodos_Fornecedores

            giTabFornecedor_Alterado = REGISTRO_ALTERADO

    End Select

End Sub


Private Sub MarcaTodos_TipoProduto()
'Marca todos os Tipos de Produto da tela

Dim iIndice As Integer

    'Percorre todas as checkbox de TipoProduto
    For iIndice = 0 To TipoProduto.ListCount - 1

        'Marca na tela o bloqueio em questão
        TipoProduto.Selected(iIndice) = True

    Next

End Sub

Private Sub MarcaTodas_Requisicoes()
'Marca todas as Requisicoes do GridRequisicoes

Dim objReqCompras As ClassRequisicaoCompras
Dim iLinha As Integer
Dim lErro As Long

On Error GoTo Erro_MarcaTodas_Requisicoes

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridRequisicoes.iLinhasExistentes
        
        If GridRequisicoes.TextMatrix(iLinha, iGrid_EscolhidoReq_Col) = GRID_CHECKBOX_INATIVO Then
        
            Set objReqCompras = gobjGeracaoCotacao.colReqCompra(iLinha)
    
            objReqCompras.iSelecionado = Selecionado
            
            'Atualiza coleções globais
            lErro = Atualiza_SelecaoReqCompra(objReqCompras, gcolCotacaoProduto, gcolFornecedorProdutoFF, gcolItemGridFornecedores)
            If lErro <> SUCESSO Then gError 77034
    
            'Marca na tela a Requisicao em questão
            GridRequisicoes.TextMatrix(iLinha, iGrid_EscolhidoReq_Col) = GRID_CHECKBOX_ATIVO
            
        End If

    Next

    'Atualiza na tela checkboxes marcadas
    Call Grid_Refresh_Checkbox(objGridRequisicoes)

    Exit Sub

Erro_MarcaTodas_Requisicoes:

    Select Case gErr

        Case 77034
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161412)

    End Select

    Exit Sub

End Sub

Private Sub MarcaTodos_ItensRequisicoes()
'Marca todos os Itens do GridItensRequisicoes

Dim iLinha As Integer
Dim lErro As Long
Dim objItemReqCompras As ClassItemReqCompras

On Error GoTo Erro_MarcaTodos_ItensRequisicoes

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridItensRequisicoes.iLinhasExistentes

        If GridItensRequisicoes.TextMatrix(iLinha, iGrid_EscolhidoItem_Col) = GRID_CHECKBOX_INATIVO Then
            
            'Faz objItemReqCompras apontar para elemento correspondente na coleção global
            lErro = ItemReqCompras_Escolhe(iLinha, objItemReqCompras)
            If lErro <> SUCESSO Then gError 77045
            
            objItemReqCompras.iSelecionado = Selecionado
            
            'Atualiza colecoes globais
            lErro = Atualiza_Selecao_ItemReqCompras(objItemReqCompras, gcolCotacaoProduto, gcolFornecedorProdutoFF, gcolItemGridFornecedores)
            If lErro <> SUCESSO Then gError 77043
            
            'Marca na tela o ItemRequisicao em questão
            GridItensRequisicoes.TextMatrix(iLinha, iGrid_EscolhidoItem_Col) = GRID_CHECKBOX_ATIVO

        End If

    Next

    'Atualiza na tela a checkbox marcada
    Call Grid_Refresh_Checkbox(objGridItensRequisicoes)

    Exit Sub

Erro_MarcaTodos_ItensRequisicoes:

    Select Case gErr

        Case 77043, 77045
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161413)

    End Select

    Exit Sub

End Sub

Private Sub MarcaTodos_Produtos()
'Marca todos os Produtos do GridProdutos

Dim iLinha As Integer
Dim lErro As Long
Dim objCotacaoProduto As ClassCotacaoProduto

On Error GoTo Erro_MarcaTodos_Produtos

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridProdutos.iLinhasExistentes

        'Faz objCotacaoProduto apontar para elemento correspondente na coleção global
        lErro = CotacaoProduto_Escolhe(iLinha, objCotacaoProduto)
        If lErro <> SUCESSO Then gError 77065

        objCotacaoProduto.iEscolhido = Selecionado
        
        'Marca na tela o Produto
        GridProdutos.TextMatrix(iLinha, iGrid_EscolhidoProd_Col) = GRID_CHECKBOX_ATIVO

    Next

    'Atualiza colecoes globais
    lErro = Atualiza_Selecao_CotacaoProduto(gcolCotacaoProduto, gcolFornecedorProdutoFF, gcolItemGridFornecedores)
    If lErro <> SUCESSO Then gError 77066

    'Atualiza na tela checkboxes marcadas
    Call Grid_Refresh_Checkbox(objGridProdutos)

    Exit Sub

Erro_MarcaTodos_Produtos:

    Select Case gErr

        Case 77065, 77066
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161414)

    End Select

    Exit Sub

End Sub

Private Sub MarcaTodos_Fornecedores()
'Marca todos os Fornecedores do GridFornecedores

Dim iLinha As Integer
Dim lErro As Long
Dim objFornecedorProdutoFF As ClassFornecedorProdutoFF
Dim objItemGridFornecedor As ClassItemGridFornecedores

On Error GoTo Erro_MarcaTodos_Fornecedores

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridFornecedores.iLinhasExistentes

        'Faz objItemGridFornecedor e objFornecedorProdutoFF apontarem para elementos correspondentes na colecao global
        lErro = ItemGridFornecedor_Escolhe(iLinha, objItemGridFornecedor, objFornecedorProdutoFF)
        If lErro <> SUCESSO Then gError 77076
            
        objItemGridFornecedor.sEscolhido = CStr(Selecionado)
        objFornecedorProdutoFF.iEscolhido = Selecionado
        
        'Marca na tela o pedido em questão
        GridFornecedores.TextMatrix(iLinha, iGrid_EscolhidoForn_Col) = GRID_CHECKBOX_ATIVO

    Next

    'Atualiza na tela a checkbox marcada
    Call Grid_Refresh_Checkbox(objGridFornecedores)

    Exit Sub

Erro_MarcaTodos_Fornecedores:
    
    Select Case gErr

        Case 77076
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161415)

    End Select

    Exit Sub

End Sub

Private Sub DesmarcaTodos_Fornecedores()
'Desmarca todos os Fornecedores do GridFornecedores

Dim iLinha As Integer
Dim lErro As Long
Dim objFornecedorProdutoFF As ClassFornecedorProdutoFF
Dim objItemGridFornecedor As ClassItemGridFornecedores

On Error GoTo Erro_DesmarcaTodos_Fornecedores

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridFornecedores.iLinhasExistentes
        
        'Faz objItemGridFornecedor e objFornecedorProdutoFF apontarem para elementos correspondentes na colecao global
        lErro = ItemGridFornecedor_Escolhe(iLinha, objItemGridFornecedor, objFornecedorProdutoFF)
        If lErro <> SUCESSO Then gError 77077
            
        objItemGridFornecedor.sEscolhido = CStr(NAO_SELECIONADO)
        objFornecedorProdutoFF.iEscolhido = NAO_SELECIONADO

        'Desmarca na tela o pedido em questão
        GridFornecedores.TextMatrix(iLinha, iGrid_EscolhidoForn_Col) = GRID_CHECKBOX_INATIVO

    Next

    'Atualiza na tela a checkbox desmarcada
    Call Grid_Refresh_Checkbox(objGridFornecedores)

    Exit Sub

Erro_DesmarcaTodos_Fornecedores:
    
    Select Case gErr

        Case 77077

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161416)

    End Select

    Exit Sub

End Sub

Private Sub DesmarcaTodos_Produtos()
'Desmarca todos os Produtos do GridProdutos

Dim iLinha As Integer
Dim lErro As Long
Dim objCotacaoProduto As ClassCotacaoProduto

On Error GoTo Erro_DesmarcaTodos_Produtos

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridProdutos.iLinhasExistentes

        'Faz objCotacaoProduto apontar para elemento correspondente na coleção global
        lErro = CotacaoProduto_Escolhe(iLinha, objCotacaoProduto)
        If lErro <> SUCESSO Then gError 77067

        objCotacaoProduto.iEscolhido = NAO_SELECIONADO
        
        'Desmarca na tela o Produto
        GridProdutos.TextMatrix(iLinha, iGrid_EscolhidoProd_Col) = GRID_CHECKBOX_INATIVO

    Next

    'Atualiza colecoes globais
    lErro = Atualiza_Selecao_CotacaoProduto(gcolCotacaoProduto, gcolFornecedorProdutoFF, gcolItemGridFornecedores)
    If lErro <> SUCESSO Then gError 77068

    'Atualiza na tela checkboxes marcadas
    Call Grid_Refresh_Checkbox(objGridProdutos)

    Exit Sub

Erro_DesmarcaTodos_Produtos:

    Select Case gErr

        Case 77067, 77068
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161417)

    End Select

    Exit Sub

End Sub

Private Sub DesmarcaTodos_ItensRequisicoes()
'Desmarca todos os Itens do GridItensRequisicoes

Dim iLinha As Integer
Dim lErro As Long
Dim objItemReqCompras As ClassItemReqCompras

On Error GoTo Erro_MarcaTodos_ItensRequisicoes

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridItensRequisicoes.iLinhasExistentes

        If GridItensRequisicoes.TextMatrix(iLinha, iGrid_EscolhidoItem_Col) = GRID_CHECKBOX_ATIVO Then
            
            'Faz objItemReqCompras apontar para elemento correspondente na coleção global
            lErro = ItemReqCompras_Escolhe(iLinha, objItemReqCompras)
            If lErro <> SUCESSO Then gError 77046
            
            objItemReqCompras.iSelecionado = NAO_SELECIONADO
            
            'Atualiza colecoes globais
            lErro = Atualiza_Selecao_ItemReqCompras(objItemReqCompras, gcolCotacaoProduto, gcolFornecedorProdutoFF, gcolItemGridFornecedores)
            If lErro <> SUCESSO Then gError 77044
            
            'Desmarca na tela o ItemRequisicao em questão
            GridItensRequisicoes.TextMatrix(iLinha, iGrid_EscolhidoItem_Col) = GRID_CHECKBOX_INATIVO

        End If

    Next

    'Atualiza na tela a checkbox marcada
    Call Grid_Refresh_Checkbox(objGridItensRequisicoes)

    Exit Sub

Erro_MarcaTodos_ItensRequisicoes:

    Select Case gErr

        Case 77044, 77046
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161418)

    End Select
    
    Exit Sub

End Sub

Private Sub DesmarcaTodas_Requisicoes()
'Desmarca todas as Requisicoes do GridRequisicoes

Dim objReqCompras As ClassRequisicaoCompras
Dim iLinha As Integer
Dim lErro As Long

On Error GoTo Erro_DesmarcaTodas_Requisicoes

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridRequisicoes.iLinhasExistentes
        
        If GridRequisicoes.TextMatrix(iLinha, iGrid_EscolhidoReq_Col) = GRID_CHECKBOX_ATIVO Then
        
            Set objReqCompras = gobjGeracaoCotacao.colReqCompra(iLinha)
    
            objReqCompras.iSelecionado = NAO_SELECIONADO
            
            'Atualiza coleções globais
            lErro = Atualiza_SelecaoReqCompra(objReqCompras, gcolCotacaoProduto, gcolFornecedorProdutoFF, gcolItemGridFornecedores)
            If lErro <> SUCESSO Then gError 77035
    
            'Desmarca na tela a Requisicao em questão
            GridRequisicoes.TextMatrix(iLinha, iGrid_EscolhidoReq_Col) = GRID_CHECKBOX_INATIVO
            
        End If

    Next

    'Atualiza na tela checkboxes marcadas
    Call Grid_Refresh_Checkbox(objGridRequisicoes)

    Exit Sub

Erro_DesmarcaTodas_Requisicoes:

    Select Case gErr

        Case 77035
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161419)

    End Select

    Exit Sub

End Sub

Private Sub DesmarcaTodos_TipoProduto()
'Desmarca todas as checkbox da ListBox TipoProduto

Dim iIndice As Integer

    'Percorre todas as checkbox de TipoProduto
    For iIndice = 0 To TipoProduto.ListCount - 1

        'Desmarca na tela o tipo de produto em questão
        TipoProduto.Selected(iIndice) = False

    Next

End Sub


Sub Busca_Produto(sProduto As String, objProduto As ClassProduto, colProdutos As Collection, bAchou As Boolean)

Dim objProdutoAux As ClassProduto

    bAchou = False
    
    'Verifica se o produto está na coleção de produtos lidos
    For Each objProdutoAux In colProdutos
        'Se encontrou transfere a descrição para o objProduto
        If objProdutoAux.sCodigo = sProduto Then
            bAchou = True 'simboliza que achou
            Set objProduto = objProdutoAux
            Exit For
        End If
    Next

    Exit Sub
    
End Sub

Sub Busca_Fornecedor(lFornecedor As Long, objFornecedor As ClassFornecedor, colFornecedor As Collection, bAchou As Boolean)

Dim objFornecAux As ClassFornecedor

    bAchou = False
    
    For Each objFornecAux In colFornecedor
    
        If objFornecAux.lCodigo = lFornecedor Then
            Set objFornecedor = objFornecAux
            bAchou = True
            Exit For
        End If
    Next
    
    Exit Sub
    
End Sub

Sub Busca_Almoxarifado(iAlmoxarifado As Integer, objAlmoxarifado As ClassAlmoxarifado, colAlmoxarifado As Collection, bAchou As Boolean)

Dim objAlmoxarifadoAux As ClassAlmoxarifado

    bAchou = False
    
    'Verifica se o Almoxarifado está na coleção de Almoxarifados lidos
    For Each objAlmoxarifadoAux In colAlmoxarifado
        'Se encontrou transfere a descrição para o objAlmoxarifado
        If objAlmoxarifadoAux.iCodigo = iAlmoxarifado Then
            bAchou = True 'simboliza que achou
            Set objAlmoxarifado = objAlmoxarifadoAux
            Exit For
        End If
    Next

    Exit Sub
    
End Sub

Private Sub BotaoEmail_Click()

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim sMailTo As String
Dim objPedidoCotacao As New ClassPedidoCotacao
Dim objFilialFornecedor As New ClassFilialFornecedor, objFornecedor As New ClassFornecedor
Dim objEndereco As New ClassEndereco
Dim objCotacao As New ClassCotacao
Dim colPedidoCotacao As New Collection
Dim bIgnorar As Boolean
Dim lResposta As Long, sInfoEmail As String

On Error GoTo Erro_BotaoEmail_Click

    Set gobjCotacao = Nothing

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 129319
    
    If gobjCotacao Is Nothing Then gError 129320

    GL_objMDIForm.MousePointer = vbHourglass
    
    bIgnorar = False
    
    For Each objPedidoCotacao In gcolPedidoCotacao
           
        sInfoEmail = ""
        
        If objPedidoCotacao.lFornecedor <> 0 And objPedidoCotacao.iFilial <> 0 Then
    
            objFilialFornecedor.lCodFornecedor = objPedidoCotacao.lFornecedor
            objFilialFornecedor.iCodFilial = objPedidoCotacao.iFilial
    
            lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 12929 Then gError 129323
             
            If lErro = SUCESSO Then
            
                objEndereco.lCodigo = objFilialFornecedor.lEndereco
                
                lErro = CF("Endereco_Le", objEndereco)
                If lErro <> SUCESSO Then gError 129324
            
                sMailTo = objEndereco.sEmail
                
            End If
            
            objFornecedor.lCodigo = objPedidoCotacao.lFornecedor
            lErro = CF("Fornecedor_Le", objFornecedor)
            If lErro <> SUCESSO And lErro <> 12729 Then gError 129323

            If lErro = SUCESSO Then sInfoEmail = "Fornecedor: " & CStr(objFornecedor.lCodigo) & " - " & objFornecedor.sNomeReduzido & " . Filial: " & CStr(objFilialFornecedor.iCodFilial) & " - " & objFilialFornecedor.sNome
            
        End If
        
        If Len(Trim(sMailTo)) = 0 And Not bIgnorar Then
            
            lResposta = Rotina_Aviso(vbYesNoCancel, "ERRO_PEDCOTACAO_IMPRESSAO_LOOP", objFilialFornecedor.iCodFilial, objFilialFornecedor.lCodFornecedor)
            
            Select Case lResposta
            
                Case vbYes
                
                Case vbNo
                    bIgnorar = True
                
                Case vbCancel
                    gError 129322
                
            End Select
            
        End If
        
        'Dispara a impressão do relatório
        Set objRelatorio = New AdmRelatorio
        lErro = objRelatorio.ExecutarDiretoEmail("Pedido de Cotação", "PEDCOTTO.NumIntDoc = @NPEDCOT", 0, "PEDCOT", "NPEDCOT", objPedidoCotacao.lNumIntDoc, "TTO_EMAIL", sMailTo, "TSUBJECT", "Pedido de Cotação " & CStr(objPedidoCotacao.lCodigo), "TALIASATTACH", "PedCot" & CStr(objPedidoCotacao.lCodigo), "TINFO_EMAIL", sInfoEmail)
        If lErro <> SUCESSO Then gError 129325
    
        'Atualiza data de emissao no BD para a data atual
        lErro = CF("PedidoCotacao_Atualiza_DataEmissao", objPedidoCotacao)
        If lErro <> SUCESSO And lErro <> 56348 Then gError 129321
       
    Next
    
    GL_objMDIForm.MousePointer = vbDefault

    Call Limpa_Tela_GeracaoCotacao
    
    giAlterado = 0

    Exit Sub
    
Erro_BotaoEmail_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 129319
        
        Case 129320
            Call Rotina_Erro(vbOKOnly, "ERRO_PED_COTACAO_NAO_GERADO", gErr)
                   
        Case 129321, 129323, 129324
        
        Case 129322
            
        Case 129325
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDCOTACAO_IMPRESSAO", gErr)
             
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161420)

    End Select
    
    Exit Sub
    
End Sub

'#########################################################
'Inserido por Wagner
Sub Refaz_Grid(ByVal objGridInt As AdmGrid, ByVal iNumLinhas As Integer)
    objGridInt.objGrid.Rows = iNumLinhas + 1

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
End Sub
'#########################################################

Private Function Preenche_CodigoPV(objRequisicaoCompra As ClassRequisicaoCompras, lCodigoPV As Long) As Long

Dim objOrdemProducao As New ClassOrdemDeProducao
Dim lErro As Long
Dim objItemOP As ClassItemOP
Dim iFilialPV As Integer

On Error GoTo Erro_Preenche_CodigoPV

    If Len(Trim(objRequisicaoCompra.sOPCodigo)) <> 0 Then
    
        objOrdemProducao.iFilialEmpresa = giFilialEmpresa
        objOrdemProducao.sCodigo = objRequisicaoCompra.sOPCodigo
    
        lErro = CF("ItensOrdemProducao_Le", objOrdemProducao)
        If lErro <> SUCESSO And lErro <> 30401 Then gError 178849

        If lErro <> SUCESSO Then
        
            lErro = CF("ItensOP_Baixada_Le", objOrdemProducao)
            If lErro <> SUCESSO And lErro <> 178689 Then gError 178850
        
        End If
        
        If lErro = SUCESSO Then
        
            For Each objItemOP In objOrdemProducao.colItens
                
                If objItemOP.lCodPedido <> 0 Then
                    lCodigoPV = objItemOP.lCodPedido
                    Exit For
                End If
                
                If objItemOP.lNumIntDocPai <> 0 Then
                
                    lErro = CF("ItensOP_Le_PV", objItemOP.lNumIntDocPai, lCodigoPV, iFilialPV)
                    If lErro <> SUCESSO And lErro <> 178696 And lErro <> 178697 Then gError 178851
            
                End If
            
                If lCodigoPV <> 0 Then
                    Exit For
                End If
            
            Next
    
        End If
    
    End If

    Preenche_CodigoPV = SUCESSO
    
    Exit Function
    
Erro_Preenche_CodigoPV:

    Preenche_CodigoPV = gErr
    
    Select Case gErr
    
        Case 178849 To 178851
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178852)

    End Select

    Exit Function

End Function


