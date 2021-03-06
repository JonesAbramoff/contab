VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl GeracaoPedCompraReqOcx 
   ClientHeight    =   9195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   ScaleHeight     =   9195
   ScaleWidth      =   16995
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8190
      Index           =   5
      Left            =   195
      TabIndex        =   108
      Top             =   795
      Visible         =   0   'False
      Width           =   16590
      Begin VB.CommandButton BotaoPedCotacao 
         Caption         =   "Pedido de Cota��o ..."
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
         Left            =   6345
         TabIndex        =   140
         Top             =   135
         Width           =   2205
      End
      Begin VB.Frame Frame4 
         Caption         =   "Op��o"
         Height          =   1536
         Index           =   1
         Left            =   12900
         TabIndex        =   137
         Top             =   6510
         Width           =   3450
         Begin VB.CommandButton BotaoGravaConcorrencia 
            Caption         =   "Grava Concorr�ncia"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   516
            Left            =   408
            TabIndex        =   139
            Top             =   288
            Width           =   2670
         End
         Begin VB.CommandButton BotaoGeraPedidos 
            Caption         =   "Gera Pedidos de Compra"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   516
            Left            =   408
            TabIndex        =   138
            Top             =   888
            Width           =   2670
         End
      End
      Begin VB.ComboBox OrdenacaoCot 
         Height          =   315
         ItemData        =   "GeracaoPedCompraReqOcx.ctx":0000
         Left            =   2310
         List            =   "GeracaoPedCompraReqOcx.ctx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   136
         Top             =   150
         Width           =   2325
      End
      Begin VB.Frame FrameCotacoes 
         Caption         =   "Cota��es"
         Height          =   5925
         Index           =   2
         Left            =   45
         TabIndex        =   110
         Top             =   480
         Width           =   16290
         Begin VB.ComboBox Moeda 
            Enabled         =   0   'False
            Height          =   315
            Left            =   825
            Style           =   2  'Dropdown List
            TabIndex        =   150
            Top             =   1275
            Width           =   1680
         End
         Begin MSMask.MaskEdBox PrecoUnitarioReal 
            Height          =   228
            Left            =   2496
            TabIndex        =   151
            Top             =   1668
            Width           =   1656
            _ExtentX        =   2937
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
            Format          =   "#,##0.00###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TaxaForn 
            Height          =   225
            Left            =   4020
            TabIndex        =   152
            Top             =   1440
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Cotacao 
            Height          =   225
            Left            =   5325
            TabIndex        =   153
            Top             =   1440
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.ComboBox MotivoEscolhaCot 
            Height          =   315
            Left            =   6360
            TabIndex        =   114
            Text            =   "MotivoEscolhaCot"
            Top             =   2145
            Width           =   1995
         End
         Begin VB.CheckBox EscolhidoCot 
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
            TabIndex        =   113
            Top             =   240
            Width           =   840
         End
         Begin VB.TextBox DescProdutoCot 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   2520
            MaxLength       =   50
            TabIndex        =   112
            Top             =   270
            Width           =   4000
         End
         Begin VB.ComboBox TipoTributacaoCot 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2220
            Style           =   2  'Dropdown List
            TabIndex        =   111
            Top             =   990
            Width           =   2565
         End
         Begin MSMask.MaskEdBox DataCotacao 
            Height          =   225
            Left            =   4230
            TabIndex        =   115
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
            Left            =   825
            TabIndex        =   116
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
            TabIndex        =   117
            Top             =   2310
            Width           =   1080
            _ExtentX        =   1905
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
            Format          =   "#,##0.00###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PedCotacao 
            Height          =   225
            Left            =   7260
            TabIndex        =   118
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
            TabIndex        =   119
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
            TabIndex        =   120
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
            Height          =   228
            Left            =   4092
            TabIndex        =   121
            Top             =   1824
            Width           =   1824
            _ExtentX        =   3201
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
            TabIndex        =   122
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
            TabIndex        =   123
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
            TabIndex        =   124
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
            TabIndex        =   125
            Top             =   2295
            Width           =   1005
            _ExtentX        =   1773
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
         Begin MSMask.MaskEdBox CondPagto 
            Height          =   225
            Left            =   1245
            TabIndex        =   126
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
            TabIndex        =   127
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
            TabIndex        =   128
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
            TabIndex        =   129
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
            TabIndex        =   130
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
            TabIndex        =   131
            Top             =   270
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridCotacoes 
            Height          =   5400
            Left            =   105
            TabIndex        =   132
            Top             =   255
            Width           =   16050
            _ExtentX        =   28310
            _ExtentY        =   9525
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
            TabIndex        =   133
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
            TabIndex        =   134
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
            TabIndex        =   135
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
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   3495
         Picture         =   "GeracaoPedCompraReqOcx.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   109
         ToolTipText     =   "Numera��o Autom�tica"
         Top             =   7665
         Width           =   300
      End
      Begin MSMask.MaskEdBox Descricao 
         Height          =   570
         Left            =   2340
         TabIndex        =   141
         Top             =   6990
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   1005
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   50
         PromptChar      =   " "
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
         Left            =   1170
         TabIndex        =   149
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor Total dos Itens:"
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
         Left            =   450
         TabIndex        =   148
         Top             =   6645
         Width           =   1845
      End
      Begin VB.Label TotalItens 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2340
         TabIndex        =   147
         Top             =   6600
         Width           =   1155
      End
      Begin VB.Label TaxaEmpresa 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5925
         TabIndex        =   146
         Top             =   6615
         Width           =   1155
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
         Height          =   195
         Left            =   4425
         TabIndex        =   145
         Top             =   6645
         Width           =   1455
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Concorr�ncia:"
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
         Left            =   1080
         TabIndex        =   144
         Top             =   7710
         Width           =   1215
      End
      Begin VB.Label Concorrencia 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2340
         TabIndex        =   143
         Top             =   7665
         Width           =   1155
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         Caption         =   "Descri��o:"
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
         Left            =   1365
         TabIndex        =   142
         Top             =   7050
         Width           =   930
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8115
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   16665
      Begin VB.Frame FrameProdutos 
         BorderStyle     =   0  'None
         Height          =   7215
         Index           =   2
         Left            =   150
         TabIndex        =   54
         Top             =   315
         Visible         =   0   'False
         Width           =   16380
         Begin VB.TextBox DescProduto2 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   2430
            MaxLength       =   50
            TabIndex        =   56
            Top             =   270
            Width           =   4000
         End
         Begin MSMask.MaskEdBox FilialDestino 
            Height          =   225
            Left            =   540
            TabIndex        =   62
            Top             =   3540
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
            Left            =   7050
            TabIndex        =   61
            Top             =   270
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
         Begin MSMask.MaskEdBox TipoDestinoProd 
            Height          =   225
            Left            =   6000
            TabIndex        =   60
            Top             =   300
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
            Left            =   1710
            TabIndex        =   63
            Top             =   3555
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
            TabIndex        =   58
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
            Left            =   3615
            TabIndex        =   64
            Top             =   3585
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
         Begin MSMask.MaskEdBox Quantidade2 
            Height          =   225
            Left            =   4995
            TabIndex        =   59
            Top             =   270
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
         Begin MSMask.MaskEdBox Produto2 
            Height          =   225
            Left            =   1155
            TabIndex        =   57
            Top             =   270
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridProdutos2 
            Height          =   6825
            Left            =   75
            TabIndex        =   55
            Top             =   330
            Width           =   16245
            _ExtentX        =   28654
            _ExtentY        =   12039
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
         Height          =   7140
         Index           =   1
         Left            =   150
         TabIndex        =   44
         Top             =   360
         Width           =   16380
         Begin VB.CommandButton BotaoDesmarcarTodosProd 
            Caption         =   "Desmarcar Todos"
            Height          =   570
            Left            =   1890
            Picture         =   "GeracaoPedCompraReqOcx.ctx":00EE
            Style           =   1  'Graphical
            TabIndex        =   88
            Top             =   6555
            Width           =   1425
         End
         Begin VB.CommandButton BotaoMarcarTodosProd 
            Caption         =   "Marcar Todos"
            Height          =   570
            Left            =   240
            Picture         =   "GeracaoPedCompraReqOcx.ctx":12D0
            Style           =   1  'Graphical
            TabIndex        =   87
            Top             =   6555
            Width           =   1425
         End
         Begin VB.TextBox DescProduto1 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   2535
            MaxLength       =   50
            TabIndex        =   48
            Top             =   270
            Width           =   4000
         End
         Begin VB.CheckBox EscolhidoProduto 
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
            Left            =   540
            TabIndex        =   46
            Top             =   240
            Width           =   990
         End
         Begin MSMask.MaskEdBox QuantUrgente 
            Height          =   225
            Left            =   6195
            TabIndex        =   51
            Top             =   270
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
         Begin MSMask.MaskEdBox UnidadeMed1 
            Height          =   225
            Left            =   4005
            TabIndex        =   49
            Top             =   300
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
            Height          =   225
            Left            =   3060
            TabIndex        =   53
            Top             =   2835
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
         Begin MSMask.MaskEdBox Fornecedor1 
            Height          =   225
            Left            =   1005
            TabIndex        =   52
            Top             =   2235
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
         Begin MSMask.MaskEdBox Quantidade1 
            Height          =   225
            Left            =   5130
            TabIndex        =   50
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
         Begin MSMask.MaskEdBox Produto1 
            Height          =   225
            Left            =   1260
            TabIndex        =   47
            Top             =   270
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridProdutos1 
            Height          =   6300
            Left            =   105
            TabIndex        =   45
            Top             =   120
            Width           =   16185
            _ExtentX        =   28549
            _ExtentY        =   11113
            _Version        =   393216
            Rows            =   12
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.CommandButton BotaoEditarProduto 
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
         Index           =   3
         Left            =   120
         TabIndex        =   65
         Top             =   7695
         Width           =   1395
      End
      Begin MSComctlLib.TabStrip TabProdutos 
         Height          =   7620
         Left            =   105
         TabIndex        =   82
         Top             =   0
         Width           =   16485
         _ExtentX        =   29078
         _ExtentY        =   13441
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Sele��o"
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
      Height          =   8115
      Index           =   3
      Left            =   210
      TabIndex        =   3
      Top             =   795
      Visible         =   0   'False
      Width           =   16530
      Begin VB.CommandButton BotaoMarcarTodosItensRC 
         Caption         =   "Marcar Todos"
         Height          =   570
         Left            =   60
         Picture         =   "GeracaoPedCompraReqOcx.ctx":22EA
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   7515
         Width           =   1425
      End
      Begin VB.CommandButton BotaoDesmarcarTodosItensRC 
         Caption         =   "Desmarcar Todos"
         Height          =   570
         Left            =   1770
         Picture         =   "GeracaoPedCompraReqOcx.ctx":3304
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   7515
         Width           =   1425
      End
      Begin VB.Frame Frame5 
         Caption         =   "Itens de Requisi��es"
         Height          =   7275
         Left            =   30
         TabIndex        =   70
         Top             =   180
         Width           =   16470
         Begin VB.TextBox ObservacaoItemRC 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   6975
            MaxLength       =   255
            TabIndex        =   43
            Top             =   3315
            Width           =   5130
         End
         Begin MSMask.MaskEdBox CclItemRC 
            Height          =   225
            Left            =   2295
            TabIndex        =   39
            Top             =   3210
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ExclusivoItemRC 
            Height          =   225
            Left            =   5745
            TabIndex        =   42
            Top             =   3165
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
            Left            =   990
            TabIndex        =   38
            Top             =   3300
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
         Begin MSMask.MaskEdBox QuantRecebida 
            Height          =   225
            Left            =   -90
            TabIndex        =   37
            Top             =   3255
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
            Left            =   1425
            TabIndex        =   36
            Top             =   2970
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
         Begin MSMask.MaskEdBox FilialFornItemRC 
            Height          =   225
            Left            =   4470
            TabIndex        =   41
            Top             =   3270
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
         Begin MSMask.MaskEdBox FornecedorItemRC 
            Height          =   225
            Left            =   3075
            TabIndex        =   40
            Top             =   3075
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
            Left            =   240
            TabIndex        =   35
            Top             =   2985
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
         Begin MSMask.MaskEdBox CodigoReqItem 
            Height          =   225
            Left            =   1830
            TabIndex        =   29
            Top             =   255
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FilialReqItem 
            Height          =   225
            Left            =   915
            TabIndex        =   28
            Top             =   285
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
            Left            =   300
            TabIndex        =   27
            Top             =   360
            Width           =   1035
         End
         Begin VB.TextBox DescProdutoItemRC 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   4980
            MaxLength       =   50
            TabIndex        =   32
            Top             =   330
            Width           =   4000
         End
         Begin MSMask.MaskEdBox Item 
            Height          =   225
            Left            =   2970
            TabIndex        =   30
            Top             =   225
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   3
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox UnidadeMedItemRC 
            Height          =   225
            Left            =   6480
            TabIndex        =   33
            Top             =   330
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   5
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantComprarItemRC 
            Height          =   225
            Left            =   7650
            TabIndex        =   34
            Top             =   345
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
         Begin MSMask.MaskEdBox ProdutoItemRC 
            Height          =   225
            Left            =   3795
            TabIndex        =   31
            Top             =   300
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
            Height          =   3135
            Left            =   90
            TabIndex        =   26
            Top             =   330
            Width           =   16260
            _ExtentX        =   28681
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8145
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   855
      Width           =   16680
      Begin VB.Frame Frame6 
         Caption         =   "Exibe Requisi��es"
         Height          =   6600
         Left            =   165
         TabIndex        =   73
         Top             =   30
         Width           =   8610
         Begin VB.ListBox ItensCategoria 
            Height          =   1860
            ItemData        =   "GeracaoPedCompraReqOcx.ctx":44E6
            Left            =   1185
            List            =   "GeracaoPedCompraReqOcx.ctx":44E8
            Style           =   1  'Checkbox
            TabIndex        =   157
            Top             =   915
            Width           =   2400
         End
         Begin VB.ComboBox Categoria 
            Height          =   315
            Left            =   1185
            Style           =   2  'Dropdown List
            TabIndex        =   156
            Top             =   330
            Width           =   2415
         End
         Begin VB.Frame Frame8 
            Caption         =   "Data Registro"
            Height          =   1425
            Left            =   3225
            TabIndex        =   101
            Top             =   3030
            Width           =   2385
            Begin MSComCtl2.UpDown UpDownDataDe 
               Height          =   300
               Left            =   1905
               TabIndex        =   102
               Top             =   345
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown UpDownDataAte 
               Height          =   300
               Left            =   1890
               TabIndex        =   103
               Top             =   870
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataDe 
               Height          =   300
               Left            =   735
               TabIndex        =   104
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
               TabIndex        =   105
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
               TabIndex        =   107
               Top             =   420
               Width           =   315
            End
            Begin VB.Label Label2 
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
               Height          =   195
               Left            =   285
               TabIndex        =   106
               Top             =   960
               Width           =   360
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "N�mero"
            Height          =   1425
            Left            =   6180
            TabIndex        =   96
            Top             =   3030
            Width           =   1980
            Begin MSMask.MaskEdBox CodigoDe 
               Height          =   315
               Left            =   780
               TabIndex        =   97
               Top             =   390
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
               TabIndex        =   98
               Top             =   960
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
               Height          =   195
               Left            =   375
               TabIndex        =   100
               Top             =   1020
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
               TabIndex        =   99
               Top             =   450
               Width           =   315
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Data Limite"
            Height          =   1425
            Left            =   270
            TabIndex        =   89
            Top             =   3045
            Width           =   2385
            Begin MSComCtl2.UpDown UpDownDataLimDe 
               Height          =   300
               Left            =   1905
               TabIndex        =   90
               Top             =   360
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown UpDownDataLimAte 
               Height          =   300
               Left            =   1890
               TabIndex        =   91
               Top             =   885
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataLimiteDe 
               Height          =   300
               Left            =   735
               TabIndex        =   92
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
               TabIndex        =   93
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
               Height          =   195
               Left            =   285
               TabIndex        =   95
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
               TabIndex        =   94
               Top             =   420
               Width           =   315
            End
         End
         Begin VB.CommandButton BotaoDesmarcarTodos 
            Caption         =   "Desmarcar Todos"
            Height          =   540
            Index           =   2
            Left            =   6765
            Picture         =   "GeracaoPedCompraReqOcx.ctx":44EA
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   840
            Width           =   1425
         End
         Begin VB.CommandButton BotaoMarcarTodos 
            Caption         =   "Marcar Todos"
            Height          =   540
            Index           =   2
            Left            =   6765
            Picture         =   "GeracaoPedCompraReqOcx.ctx":56CC
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   225
            Width           =   1425
         End
         Begin VB.ListBox TipoProduto 
            Height          =   2310
            Left            =   3825
            Style           =   1  'Checkbox
            TabIndex        =   5
            Top             =   465
            Width           =   2775
         End
         Begin VB.Frame Frame3 
            Caption         =   "Local de Entrega"
            Height          =   1485
            Left            =   270
            TabIndex        =   74
            Top             =   4515
            Width           =   7890
            Begin VB.Frame Frame2 
               Caption         =   "Tipo"
               Height          =   585
               Left            =   135
               TabIndex        =   80
               Top             =   720
               Width           =   3750
               Begin VB.OptionButton TipoDestino 
                  Caption         =   "Filial Empresa"
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
                  Left            =   420
                  TabIndex        =   9
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   1515
               End
               Begin VB.OptionButton TipoDestino 
                  Caption         =   "Fornecedor"
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
                  Left            =   2175
                  TabIndex        =   10
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
               TabIndex        =   8
               Top             =   330
               Width           =   2445
            End
            Begin VB.Frame FrameTipoDestino 
               BorderStyle     =   0  'None
               Height          =   675
               Index           =   1
               Left            =   4140
               TabIndex        =   75
               Top             =   645
               Visible         =   0   'False
               Width           =   3645
               Begin VB.ComboBox FilialFornec 
                  Height          =   315
                  Left            =   1260
                  TabIndex        =   13
                  Top             =   360
                  Width           =   2160
               End
               Begin MSMask.MaskEdBox Fornecedor 
                  Height          =   300
                  Left            =   1245
                  TabIndex        =   12
                  Top             =   30
                  Width           =   2145
                  _ExtentX        =   3784
                  _ExtentY        =   529
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   20
                  PromptChar      =   " "
               End
               Begin VB.Label FilFornDestLabel 
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
                  TabIndex        =   77
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
                  Left            =   165
                  MousePointer    =   14  'Arrow and Question
                  TabIndex        =   76
                  Top             =   60
                  Width           =   1035
               End
            End
            Begin VB.Frame FrameTipoDestino 
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   735
               Index           =   0
               Left            =   4260
               TabIndex        =   78
               Top             =   645
               Width           =   3555
               Begin VB.ComboBox FilialEmpresa 
                  Height          =   315
                  Left            =   1035
                  TabIndex        =   11
                  Top             =   150
                  Width           =   2160
               End
               Begin VB.Label FilEmprDestLabel 
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
                  Left            =   510
                  TabIndex        =   79
                  Top             =   180
                  Width           =   465
               End
            End
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Itens:"
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
            Left            =   645
            TabIndex        =   159
            Top             =   990
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Categoria:"
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
            TabIndex        =   158
            Top             =   375
            Width           =   870
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
            Left            =   3885
            TabIndex        =   81
            Top             =   240
            Width           =   1470
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8040
      Index           =   2
      Left            =   105
      TabIndex        =   2
      Top             =   915
      Visible         =   0   'False
      Width           =   16665
      Begin VB.CommandButton BotaoMarcarTodosReq 
         Caption         =   "Marcar Todos"
         Height          =   555
         Left            =   60
         Picture         =   "GeracaoPedCompraReqOcx.ctx":66E6
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   7470
         Width           =   1425
      End
      Begin VB.CommandButton BotaoDesmarcarTodosReq 
         Caption         =   "Desmarcar Todos"
         Height          =   555
         Left            =   1680
         Picture         =   "GeracaoPedCompraReqOcx.ctx":7700
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   7455
         Width           =   1425
      End
      Begin VB.CommandButton BotaoReqCompras 
         Caption         =   "Requisi��o de Compras..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   6705
         TabIndex        =   15
         Top             =   60
         Width           =   2040
      End
      Begin VB.Frame Frame7 
         Caption         =   "Requisi��es de Compra"
         Height          =   6855
         Left            =   150
         TabIndex        =   71
         Top             =   525
         Width           =   16410
         Begin MSMask.MaskEdBox CodigoPV 
            Height          =   240
            Left            =   2475
            TabIndex        =   160
            Top             =   885
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
         Begin MSMask.MaskEdBox Requisitante 
            Height          =   240
            Left            =   6345
            TabIndex        =   23
            Top             =   390
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
            Left            =   360
            TabIndex        =   17
            Top             =   315
            Width           =   975
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
            Left            =   5910
            TabIndex        =   22
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox ObservacaoReq 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   1260
            MaxLength       =   255
            TabIndex        =   25
            Top             =   2985
            Width           =   4725
         End
         Begin MSMask.MaskEdBox FilialReq 
            Height          =   225
            Left            =   1170
            TabIndex        =   18
            Top             =   360
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
         Begin MSMask.MaskEdBox CclReq 
            Height          =   225
            Left            =   270
            TabIndex        =   24
            Top             =   2985
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
            Left            =   3585
            TabIndex        =   20
            Top             =   360
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
         Begin MSMask.MaskEdBox CodigoReq 
            Height          =   225
            Left            =   2745
            TabIndex        =   19
            Top             =   360
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
            Left            =   4755
            TabIndex        =   21
            Top             =   375
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
            Height          =   6480
            Left            =   180
            TabIndex        =   16
            Top             =   285
            Width           =   16095
            _ExtentX        =   28390
            _ExtentY        =   11430
            _Version        =   393216
            Rows            =   16
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.ComboBox OrdenacaoReq 
         Height          =   315
         ItemData        =   "GeracaoPedCompraReqOcx.ctx":88E2
         Left            =   2610
         List            =   "GeracaoPedCompraReqOcx.ctx":88E4
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   135
         Width           =   2325
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
         Left            =   1470
         TabIndex        =   72
         Top             =   165
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   15165
      ScaleHeight     =   480
      ScaleWidth      =   1575
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   15
      Width           =   1635
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   60
         Picture         =   "GeracaoPedCompraReqOcx.ctx":88E6
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Imprimir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   570
         Picture         =   "GeracaoPedCompraReqOcx.ctx":89E8
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1080
         Picture         =   "GeracaoPedCompraReqOcx.ctx":8F1A
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   8580
      Left            =   90
      TabIndex        =   0
      Top             =   480
      Width           =   16800
      _ExtentX        =   29633
      _ExtentY        =   15134
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sele��o"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Requisi��es"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Itens de Requisi��es"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Produtos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Cota��es"
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
   Begin VB.Label Comprador 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1155
      TabIndex        =   155
      Top             =   75
      Width           =   2145
   End
   Begin VB.Label Label32 
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
      Left            =   165
      TabIndex        =   154
      Top             =   105
      Width           =   975
   End
End
Attribute VB_Name = "GeracaoPedCompraReqOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Const TAB_Selecao = 1
Private Const TAB_REQUISICOES = 2
Private Const TAB_ITENSREQ = 3
Private Const TAB_Produtos = 4
Private Const TAB_COTACOES = 5

'Vari�veis Globais
Dim iFrameAtual As Integer
Dim iFrameProdutoAtual As Integer
Dim iAlterado As Integer
Dim iFrameSelecaoAlterado As Integer
Dim iFrameTipoDestinoAtual As Integer
Dim gsOrdenacao As String
Dim giPodeAumentarQuant As Integer

Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoBotaoPedCotacao As AdmEvento
Attribute objEventoBotaoPedCotacao.VB_VarHelpID = -1
Private WithEvents objEventoBotaoReqCompras As AdmEvento
Attribute objEventoBotaoReqCompras.VB_VarHelpID = -1

'GridRequisicoes
Dim objGridRequisicoes As AdmGrid
Dim iGrid_EscolhidoReq_Col As Integer
Dim iGrid_FilialReq_Col As Integer
Dim iGrid_CodigoReq_Col As Integer
Dim iGrid_DataLimite_Col As Integer
Dim iGrid_DataReq_Col As Integer
Dim iGrid_Urgente_Col As Integer
Dim iGrid_Requisitante_Col As Integer
Dim iGrid_CclReq_Col As Integer
Dim iGrid_ObservacaoReq_Col As Integer
Dim iGrid_CodigoPV_Col As Integer

'GridItensRequisicoes
Dim objGridItensRequisicoes As AdmGrid
Dim iGrid_EscolhidoItem_Col As Integer
Dim iGrid_FilialReqItem_Col As Integer
Dim iGrid_CodigoReqItem_Col As Integer
Dim iGrid_Item_Col As Integer
Dim iGrid_ProdutoItemRC_Col As Integer
Dim iGrid_DescProdutoItem_Col As Integer
Dim iGrid_UnidadeMedItem_Col As Integer
Dim iGrid_QuantComprarItem_Col As Integer
Dim iGrid_QuantidadeItem_Col As Integer
Dim iGrid_QuantPedida_Col As Integer
Dim iGrid_QuantRecebida_Col As Integer
Dim iGrid_Almoxarifado_Col As Integer
Dim iGrid_CclItemRC_Col As Integer
Dim iGrid_FornecedorItemRC_Col As Integer
Dim iGrid_FilialFornItemRC_Col As Integer
Dim iGrid_ExclusivoItemRC_Col As Integer
Dim iGrid_ObservacaoItemRC_Col As Integer

'GridProdutos1
Dim objGridProdutos1 As AdmGrid
Dim iGrid_EscolhidoProduto_Col As Integer
Dim iGrid_Produto1_Col As Integer
Dim iGrid_DescProduto1_Col As Integer
Dim iGrid_UnidadeMed1_Col As Integer
Dim iGrid_Quantidade1_Col As Integer
Dim iGrid_QuantUrgente_Col As Integer
Dim iGrid_Fornecedor1_Col As Integer
Dim iGrid_FilialForn1_Col As Integer

'GridProdutos2
Dim objGridProdutos2 As AdmGrid
Dim iGrid_Produto2_Col As Integer
Dim iGrid_DescProduto2_Col As Integer
Dim iGrid_UnidadeMed2_Col As Integer
Dim iGrid_Quantidade2_Col As Integer
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
Dim iGrid_PrecoUnitarioCot_Col As Integer
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
Dim iGrid_Moeda_Col As Integer
Dim iGrid_PrecoUnitario_RS_Col As Integer
Dim iGrid_TaxaForn_Col As Integer
Dim iGrid_CotacaoMoeda_Col As Integer

'Vari�veis globais da tela
Dim iRequisicaoAlterada As Integer
Dim iFornecedorAlterado As Integer
Dim gsTipoTributacao As String
Dim gsOrdenacaoReq As String
Dim asOrdenacaoReq(2) As String
Dim asOrdenacaoReqString(2) As String
Dim gsOrdenacaoCot As String
Dim asOrdenacaoCot(2) As String
Dim asOrdenacaoCotString(2) As String

Dim gobjGeracaoPedCompraReq As New ClassGeracaoPedCompraReq
Dim gcolItemConcorrencia As Collection
Dim gColCotacoes As Collection

Public Sub Form_Load()

Dim objUsuario As New ClassUsuario
Dim objComprador As New ClassComprador
Dim lErro As Long
Dim lConcorrencia As Long
Dim iTipoTrib As Integer
Dim sDescricao As String
Dim iFilialEmpresa As Integer

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    iFrameProdutoAtual = 1
    iFrameSelecaoAlterado = REGISTRO_ALTERADO

    '#######################################
    'Inserido por Wagner
    Call Formata_Controles
    '#######################################


    Set gcolItemConcorrencia = New Collection
    Set objEventoBotaoReqCompras = New AdmEvento
    Set objEventoFornecedor = New AdmEvento
    Set objEventoBotaoPedCotacao = New AdmEvento

    Set objGridProdutos1 = New AdmGrid
    Set objGridProdutos2 = New AdmGrid
    Set objGridCotacoes = New AdmGrid
    Set objGridRequisicoes = New AdmGrid
    Set objGridItensRequisicoes = New AdmGrid

    Set gobjGeracaoPedCompraReq = New ClassGeracaoPedCompraReq
    Set gColCotacoes = New Collection

    objComprador.sCodUsuario = gsUsuario

    'Verifica se gsUsuario � comprador
    lErro = CF("Comprador_Le_Usuario", objComprador)
    If lErro <> SUCESSO And lErro <> 50059 Then gError 63857

    'Se gsUsuario nao � comprador==> Erro
    If lErro = 50059 Then gError 63858

    giPodeAumentarQuant = objComprador.iAumentaQuant

    objUsuario.sCodUsuario = objComprador.sCodUsuario
    
    'L� o usu�rio
    lErro = CF("Usuario_Le", objUsuario)
    If lErro <> SUCESSO And lErro <> 36347 Then gError 63872

    'Se n�o encontrou o usu�rio ==> Erro
    If lErro = 36347 Then gError 63871

    'Coloca o Nome Reduzido do Comprador na tela
    Comprador.Caption = objUsuario.sNomeReduzido

    'Carrega a combo FilialEmpresa
    lErro = Carrega_FilialEmpresa()
    If lErro <> SUCESSO Then gError 63860

    'Carrega a listbox TipoProduto
    lErro = Carrega_TipoProduto()
    If lErro <> SUCESSO Then gError 63861

    'Carrega Tipos de Tributa��o
    lErro = Carrega_TipoTributacao()
    If lErro <> SUCESSO Then gError 66122

    'Preenche a combo de Ordenacao de Requisicoes
    Call OrdenacaoReq_Carrega

    'Preenche a combo de Ordenacao de Cotacoes
    Call OrdenacaoCot_Carrega

    'Preenche a combo de MotivoEscolha
    lErro = Carrega_MotivoEscolha()
    If lErro <> SUCESSO Then gError 63862

    'Inicializa a m�scara de Produto1
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto1)
    If lErro <> SUCESSO Then gError 63863

    'Inicializa a m�scara de Produto2
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto2)
    If lErro <> SUCESSO Then gError 63864

    'Inicializa a m�scara de ProdutoCot
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoCot)
    If lErro <> SUCESSO Then gError 63904

    'Inicializa a m�scara de ProdutoItemRC
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoItemRC)
    If lErro <> SUCESSO Then gError 63905

    'Coloca as Quantidades da tela no formato de Estoque
    Quantidade1.Format = FORMATO_ESTOQUE
    Quantidade2.Format = FORMATO_ESTOQUE
    QuantComprarCot.Format = FORMATO_ESTOQUE
    QuantComprarItemRC.Format = FORMATO_ESTOQUE
    QuantidadeCot.Format = FORMATO_ESTOQUE
    QuantidadeEntrega.Format = FORMATO_ESTOQUE
    QuantidadeItemRC.Format = FORMATO_ESTOQUE
    QuantPedida.Format = FORMATO_ESTOQUE
    QuantRecebida.Format = FORMATO_ESTOQUE
    QuantUrgente.Format = FORMATO_ESTOQUE

    'Seleciona o TipoDestino FilialEmpresa
    SelecionaDestino.Value = vbChecked
    TipoDestino(TIPO_DESTINO_EMPRESA).Value = True
    Call CF("Filial_Seleciona", FilialEmpresa, giFilialEmpresa)

    'Coloca FiliaEmpresa Default na Tela
    iFilialEmpresa = giFilialEmpresa
    
    lErro = CF("FilialEmpresa_Customiza", iFilialEmpresa)
    If lErro <> SUCESSO Then gError 126949
    
    FilialEmpresa.Text = iFilialEmpresa
    
    Call FilialEmpresa_Validate(bSGECancelDummy)

    'Inicializa o GridRequisicoes
    lErro = Inicializa_Grid_Requisicoes(objGridRequisicoes)
    If lErro <> SUCESSO Then gError 63866

    'Inicializa o GridItensRequisicoes
    lErro = Inicializa_Grid_ItensRequisicoes(objGridItensRequisicoes)
    If lErro <> SUCESSO Then gError 63867

    'Inicializa o GridProdutos1
    lErro = Inicializa_Grid_Produtos1(objGridProdutos1)
    If lErro <> SUCESSO Then gError 63868

    'Inicializa o GridProdutos2
    lErro = Inicializa_Grid_Produtos2(objGridProdutos2)
    If lErro <> SUCESSO Then gError 63869

    'Inicializa o GridCotacoes
    lErro = Inicializa_Grid_Cotacoes(objGridCotacoes)
    If lErro <> SUCESSO Then gError 63870
    
    lErro = Carrega_Moeda()
    If lErro <> SUCESSO Then gError 108981

    'L� o tipo de tributa��o padr�o
    lErro = CF("TipoTributacaoPadrao_Le", iTipoTrib)
    If lErro <> SUCESSO And lErro <> 66597 Then gError 68085
    If lErro = SUCESSO Then

        'L� a descri��o do Tipo de Tributa��o
        lErro = CF("TiposTributacao_Le", iTipoTrib, sDescricao)
        If lErro <> SUCESSO Then gError 68086
    
        'Guarda o Tipo de Tributa��o
        gsTipoTributacao = CStr(iTipoTrib) & SEPARADOR & sDescricao
    End If
    
    'Coloca Taxa Financeira na tela
    TaxaEmpresa.Caption = Format(gobjCOM.dTaxaFinanceiraEmpresa, "Percent")
    
    lErro = Carrega_Categorias()
    If lErro <> SUCESSO Then gError 108980

    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case gErr

        Case 63857, 63859, 63860, 63904, 63905, 68085, 68086, 66122, 108980, 108981, 126949
            'Erros tratados nas rotinas chamadas

        Case 63858
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_COMPRADOR", gErr, objComprador.sCodUsuario)

        Case 63861 To 63870
            'Erros tratados nas rotinas chamadas

        Case 63871
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", gErr, objUsuario.sCodUsuario)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161190)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

     Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    'libera as variaveis globais
    Set objEventoBotaoReqCompras = Nothing
    Set objEventoFornecedor = Nothing
    Set objEventoBotaoPedCotacao = Nothing

    Set objGridRequisicoes = Nothing
    Set objGridItensRequisicoes = Nothing
    Set objGridProdutos1 = Nothing
    Set objGridProdutos2 = Nothing
    Set objGridCotacoes = Nothing

    Set gobjGeracaoPedCompraReq = Nothing
    Set gcolItemConcorrencia = Nothing
    Set gColCotacoes = Nothing

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161191)

    End Select

    Exit Sub

End Sub

Private Sub OrdenacaoReq_Carrega()
'preenche a combo de OrdenacaoReq e inicializa variaveis globais

Dim iIndice As Integer

    'Carregar os arrays de ordena��o das Requisicoes
    asOrdenacaoReq(0) = "RequisicaoCompra.FilialEmpresa,RequisicaoCompra.Codigo"
    asOrdenacaoReq(1) = "RequisicaoCompra.DataLimite,RequisicaoCompra.FilialEmpresa,RequisicaoCompra.Codigo"
    asOrdenacaoReq(2) = "RequisicaoCompra.Data,RequisicaoCompra.FilialEmpresa,RequisicaoCompra.Codigo"

    asOrdenacaoReqString(0) = "N�mero"
    asOrdenacaoReqString(1) = "Data Limite"
    asOrdenacaoReqString(2) = "Data da Requisi��o"

    'Carrega a Combobox OrdenacaoReq
    For iIndice = 0 To 2

        OrdenacaoReq.AddItem asOrdenacaoReqString(iIndice)
        OrdenacaoReq.ItemData(OrdenacaoReq.NewIndex) = iIndice

    Next

    'Seleciona a op��o FilialEmpresa + Codigo de sele��o
    OrdenacaoReq.ListIndex = 0
    gobjGeracaoPedCompraReq.sOrdenacaoReq = OrdenacaoReq.Text
    gsOrdenacaoReq = OrdenacaoReq.Text

    Exit Sub

End Sub

Private Sub OrdenacaoCot_Carrega()
'preenche a combo de OrdenacaoCot e inicializa variaveis globais

Dim iIndice As Integer

    'Carregar os arrays de ordena��o das Cotacoes
    asOrdenacaoCot(0) = "CotacaoProduto.Produto,CotacaoProduto.Fornecedor,CotacaoProduto.Filial"
    asOrdenacaoCot(1) = "CotacaoProduto.Fornecedor,CotacaoProduto.Filial,CotacaoProduto.Produto"

    asOrdenacaoCotString(0) = "Produto"
    asOrdenacaoCotString(1) = "Fornecedor"

    'Carrega a Combobox Ordenacao
    For iIndice = 0 To 1

        OrdenacaoCot.AddItem asOrdenacaoCotString(iIndice)
        OrdenacaoCot.ItemData(OrdenacaoCot.NewIndex) = iIndice

    Next

    'Seleciona a op��o Produto + Fornecedor + Filial de sele��o
    OrdenacaoCot.ListIndex = 0
    gsOrdenacaoCot = OrdenacaoCot.Text

    Exit Sub

End Sub

Private Function Carrega_FilialEmpresa() As Long
'Carrega a combobox FilialEmpresa

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_FilialEmpresa

    'L� o C�digo e o Nome de toda FilialEmpresa do BD
    lErro = CF("Cod_Nomes_Le_FilEmp", colCodigoNome)
    If lErro <> SUCESSO Then gError 63873

    'Carrega a combo de Filial Empresa com c�digo e nome
    For Each objCodigoNome In colCodigoNome
        FilialEmpresa.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        FilialEmpresa.ItemData(FilialEmpresa.NewIndex) = objCodigoNome.iCodigo
    Next

    Carrega_FilialEmpresa = SUCESSO

    Exit Function

Erro_Carrega_FilialEmpresa:

    Carrega_FilialEmpresa = gErr

    Select Case gErr

        Case 63873
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161192)

    End Select

    Exit Function

End Function
Private Function Carrega_TipoProduto() As Long
'Carrega a ListBox TipoProduto com tipos de produtos que possam ser comprados (Compras=1)

Dim lErro As Long
Dim colCod_DescReduzida As New AdmColCodigoNome
Dim objCod_DescReduzida As New AdmCodigoNome
Dim iIndice As Integer

On Error GoTo Erro_Carrega_TipoProduto

    'Le todos os Codigos e DescReduzida de tipos de produtos cadastrados
    lErro = CF("TiposProduto_Le_Todos", colCod_DescReduzida)
    If lErro <> SUCESSO Then gError 63874

    For Each objCod_DescReduzida In colCod_DescReduzida

        'Adiciona novo item na ListBox CondPagtos
        TipoProduto.AddItem CInt(objCod_DescReduzida.iCodigo) & SEPARADOR & objCod_DescReduzida.sNome
        TipoProduto.ItemData(TipoProduto.NewIndex) = objCod_DescReduzida.iCodigo

    Next

    'Marca todos os Tipos de Produto
    For iIndice = 0 To TipoProduto.ListCount - 1
        TipoProduto.Selected(iIndice) = True
    Next

    Carrega_TipoProduto = SUCESSO

    Exit Function

Erro_Carrega_TipoProduto:

    Carrega_TipoProduto = gErr

    Select Case gErr

        Case 63874
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161193)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_ItensRequisicoes(objGridInt As AdmGrid) As Long
'Executa a Inicializa��o do grid ItensRequisicoes

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_ItensRequisicoes

    'tela em quest�o
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Escolhido")
    objGridInt.colColuna.Add ("Filial Empresa")
    objGridInt.colColuna.Add ("Requisi��o")
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descri��o")
    objGridInt.colColuna.Add ("Unid. Med.")
    objGridInt.colColuna.Add ("A Comprar")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Em Pedido")
    objGridInt.colColuna.Add ("Recebido")
    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("Centro C/L")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial Fornecedor")
    objGridInt.colColuna.Add ("Exclusividade")
    objGridInt.colColuna.Add ("Observa��o")

    'campos de edi��o do grid
    objGridInt.colCampo.Add (EscolhidoItem.Name)
    objGridInt.colCampo.Add (FilialReqItem.Name)
    objGridInt.colCampo.Add (CodigoReqItem.Name)
    objGridInt.colCampo.Add (Item.Name)
    objGridInt.colCampo.Add (ProdutoItemRC.Name)
    objGridInt.colCampo.Add (DescProdutoItemRC.Name)
    objGridInt.colCampo.Add (UnidadeMedItemRC.Name)
    objGridInt.colCampo.Add (QuantComprarItemRC.Name)
    objGridInt.colCampo.Add (QuantidadeItemRC.Name)
    objGridInt.colCampo.Add (QuantPedida.Name)
    objGridInt.colCampo.Add (QuantRecebida.Name)
    objGridInt.colCampo.Add (Almoxarifado.Name)
    objGridInt.colCampo.Add (CclItemRC.Name)
    objGridInt.colCampo.Add (FornecedorItemRC.Name)
    objGridInt.colCampo.Add (FilialFornItemRC.Name)
    objGridInt.colCampo.Add (ExclusivoItemRC.Name)
    objGridInt.colCampo.Add (ObservacaoItemRC.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_EscolhidoItem_Col = 1
    iGrid_FilialReqItem_Col = 2
    iGrid_CodigoReqItem_Col = 3
    iGrid_Item_Col = 4
    iGrid_ProdutoItemRC_Col = 5
    iGrid_DescProdutoItem_Col = 6
    iGrid_UnidadeMedItem_Col = 7
    iGrid_QuantComprarItem_Col = 8
    iGrid_QuantidadeItem_Col = 9
    iGrid_QuantPedida_Col = 10
    iGrid_QuantRecebida_Col = 11
    iGrid_Almoxarifado_Col = 12
    iGrid_CclItemRC_Col = 13
    iGrid_FornecedorItemRC_Col = 14
    iGrid_FilialFornItemRC_Col = 15
    iGrid_ExclusivoItemRC_Col = 16
    iGrid_ObservacaoItemRC_Col = 17

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridItensRequisicoes

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_REQUISICAO + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 20

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
'    GridCotacoes.Width = 8295

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama rotina de Inicializa��o do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_ItensRequisicoes = SUCESSO

    Exit Function

Erro_Inicializa_Grid_ItensRequisicoes:

    Inicializa_Grid_ItensRequisicoes = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161194)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Produtos1(objGridInt As AdmGrid) As Long
'Executa a Inicializa��o do grid Produtos1

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Produtos1

    'tela em quest�o
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Escolhido")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descri��o")
    objGridInt.colColuna.Add ("Unid. Med.")
    objGridInt.colColuna.Add ("A Comprar")
    objGridInt.colColuna.Add ("Urgente")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial Fornecedor")

    'campos de edi��o do grid
    objGridInt.colCampo.Add (EscolhidoProduto.Name)
    objGridInt.colCampo.Add (Produto1.Name)
    objGridInt.colCampo.Add (DescProduto1.Name)
    objGridInt.colCampo.Add (UnidadeMed1.Name)
    objGridInt.colCampo.Add (Quantidade1.Name)
    objGridInt.colCampo.Add (QuantUrgente.Name)
    objGridInt.colCampo.Add (Fornecedor1.Name)
    objGridInt.colCampo.Add (FilialForn1.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_EscolhidoProduto_Col = 1
    iGrid_Produto1_Col = 2
    iGrid_DescProduto1_Col = 3
    iGrid_UnidadeMed1_Col = 4
    iGrid_Quantidade1_Col = 5
    iGrid_QuantUrgente_Col = 6
    iGrid_Fornecedor1_Col = 7
    iGrid_FilialForn1_Col = 8

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridProdutos1

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_GERACAO + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 20

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
'    GridCotacoes.Width = 8295

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Chama rotina de Inicializa��o do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Produtos1 = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Produtos1:

    Inicializa_Grid_Produtos1 = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161195)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Produtos2(objGridInt As AdmGrid) As Long
'Executa a Inicializa��o do grid Produtos2

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Produtos2

    'tela em quest�o
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descri��o")
    objGridInt.colColuna.Add ("Unid. Med.")
    objGridInt.colColuna.Add ("A Comprar")
    objGridInt.colColuna.Add ("Tipo Destino")
    objGridInt.colColuna.Add ("Destino")
    objGridInt.colColuna.Add ("Filial Destino")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial Fornecedor")

    'campos de edi��o do grid
    objGridInt.colCampo.Add (Produto2.Name)
    objGridInt.colCampo.Add (DescProduto2.Name)
    objGridInt.colCampo.Add (UnidadeMed2.Name)
    objGridInt.colCampo.Add (Quantidade2.Name)
    objGridInt.colCampo.Add (TipoDestinoProd.Name)
    objGridInt.colCampo.Add (Destino.Name)
    objGridInt.colCampo.Add (FilialDestino.Name)
    objGridInt.colCampo.Add (Fornecedor2.Name)
    objGridInt.colCampo.Add (FilialForn2.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_Produto2_Col = 1
    iGrid_DescProduto2_Col = 2
    iGrid_UnidadeMed2_Col = 3
    iGrid_Quantidade2_Col = 4
    iGrid_TipoDestino_Col = 5
    iGrid_Destino_Col = 6
    iGrid_FilialDestino_Col = 7
    iGrid_Fornecedor2_Col = 8
    iGrid_FilialForn2_Col = 9

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridProdutos2

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_GERACAO + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 25

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
'    GridCotacoes.Width = 8295

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama rotina de Inicializa��o do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Produtos2 = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Produtos2:

    Inicializa_Grid_Produtos2 = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161196)

    End Select

    Exit Function

End Function

Private Sub BotaoDesmarcarTodos_Click(Index As Integer)
'Desmarca todas as checkbox da ListBox TipoProduto

Dim iIndice As Integer

    'Percorre todas as checkbox de TipoProduto
    For iIndice = 0 To TipoProduto.ListCount - 1

        'Desmarca na tela o tipo de produto em quest�o
        TipoProduto.Selected(iIndice) = False

    Next

    iFrameSelecaoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

    Exit Function

End Function

Sub Limpa_Tela_GeracaoPedCompraReq()

Dim lErro As Long
Dim lConcorrencia As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Tela_GeracaoPedCompraReq

    Call Limpa_Tela(Me)

    'Desseleciona todos os tipos de produtos da listbox TipoProduto
    For iIndice = 0 To TipoProduto.ListCount - 1
        TipoProduto.Selected(iIndice) = True
    Next

    'Limpa os Grids da tela
    Call Grid_Limpa(objGridProdutos1)
    Call Grid_Limpa(objGridProdutos2)
    Call Grid_Limpa(objGridCotacoes)
    Call Grid_Limpa(objGridRequisicoes)
    Call Grid_Limpa(objGridItensRequisicoes)
    Call Calcula_TotalItens

    SelecionaDestino.Value = vbChecked
    
    'Limpa os outros campos da tela
    FilialFornec.Clear
    iAlterado = 0
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    Concorrencia.Caption = ""

    Set gobjGeracaoPedCompraReq = Nothing
    Set gColCotacoes = New Collection
    Set gobjGeracaoPedCompraReq.colRequisicao = Nothing
    Set gobjGeracaoPedCompraReq.colTipoProduto = Nothing
    Set gobjGeracaoPedCompraReq.colTipoCategoria = Nothing
    TipoDestino(TIPO_DESTINO_EMPRESA).Value = True

    Categoria.ListIndex = -1
    ItensCategoria.Clear
    
    Call Calcula_TotalItens

    Exit Sub

Erro_Limpa_Tela_GeracaoPedCompraReq:

    Select Case gErr

        Case 63914
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161197)

    End Select

    Exit Sub

End Sub

Private Sub BotaoImprimir_Click()

Dim lErro As Long
Dim objConcorrencia As New ClassConcorrencia
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_BotaoImprimir_Click

    'Verifica se o c�digo da concorrencia esta preenchido
    If Len(Trim(Concorrencia.Caption)) = 0 Then gError 76084

    objConcorrencia.lCodigo = StrParaLong(Concorrencia.Caption)
    objConcorrencia.iFilialEmpresa = giFilialEmpresa

    'L� a Concorrencia
    lErro = CF("Concorrencia_Le", objConcorrencia)
    If lErro <> SUCESSO And lErro <> 66788 Then gError 76079

    'Se n�o encontrou a concorrencia ==> erro
    If lErro = 66788 Then gError 76080

    'Executa o relat�rio
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161198)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
'Limpa a tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 63928

    'Limpa o restante da tela
    Call Limpa_Tela_GeracaoPedCompraReq

    iAlterado = 0
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 63928
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161199)

    End Select

    Exit Sub

End Sub

Private Sub BotaoMarcarTodos_Click(Index As Integer)
'Marca todas as checkbox da ListBox TipoProduto

Dim iIndice As Integer

    'Percorre todas as checkbox de TipoProduto
    For iIndice = 0 To TipoProduto.ListCount - 1

        'Marca na tela o bloqueio em quest�o
        TipoProduto.Selected(iIndice) = True

    Next

End Sub

Private Sub BotaoProxNum_Click()
'Gera o pr�ximo n�mero de Concorrencia

Dim lErro As Long
Dim lConcorrencia As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Gera o pr�ximo c�digo para Concorrencia
    lErro = CF("Concorrencia_Automatica", lConcorrencia)
    If lErro <> SUCESSO Then gError 76082

    'Coloca o c�digo gerado na tela
    Concorrencia.Caption = lConcorrencia
    
    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 76082
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161200)

    End Select

    Exit Sub

End Sub

Private Sub CodigoAte_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodigoAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoAte, iAlterado)

End Sub

Private Sub CodigoDe_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodigoDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoDe, iAlterado)

End Sub

Private Sub DataAte_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataAte, iAlterado)

End Sub

Private Sub DataDe_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataDe, iAlterado)

End Sub

Private Sub DataLimiteAte_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataLimiteAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataLimiteAte, iAlterado)

End Sub

Private Sub DataLimiteDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataLimiteDe, iAlterado)

End Sub

Private Sub EscolhidoCot_Click()

Dim objCotItemConc As ClassCotacaoItemConc

    iAlterado = REGISTRO_ALTERADO
    
    'Localiza a cota��o correspondente
    Call Localiza_ItemCotacao(objCotItemConc, GridCotacoes.Row)
    'Atuzaliza a escolha
    objCotItemConc.iEscolhido = GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_EscolhidoCot_Col)
    'Recalcula o total dos itens selecionados
    Call Calcula_TotalItens

End Sub

Private Sub EscolhidoCot_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCotacoes)

End Sub

Private Sub EscolhidoCot_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCotacoes)

End Sub

Private Sub EscolhidoCot_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCotacoes.objControle = EscolhidoCot
    lErro = Grid_Campo_Libera_Foco(objGridCotacoes)
    If lErro <> SUCESSO Then Cancel = True

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

Private Sub ObservacaoReq_GotFocus()
    gsOrdenacao = OrdenacaoReq.Text
End Sub

Private Sub OrdenacaoCot_GotFocus()
    gsOrdenacao = OrdenacaoCot.Text
End Sub

Private Sub PrecoUnitario_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PrecoUnitario_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCotacoes)

End Sub

Private Sub PrecoUnitario_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCotacoes)

End Sub

Private Sub PrecoUnitario_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCotacoes.objControle = PrecoUnitario
    lErro = Grid_Campo_Libera_Foco(objGridCotacoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long
Dim objRequisicaoCompras As New ClassRequisicaoCompras
Dim colPedidosCotacao As New Collection
Dim iLinha As Integer
Dim iFrameAnterior
Dim iIndiceItemCategoria As Integer
Dim bSelecionouitemCateg As Boolean

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado n�o for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index = iFrameAtual Then Exit Sub

    If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

    'Torna Frame correspondente ao Tab selecionado visivel
    Frame1(TabStrip1.SelectedItem.Index).Visible = True
    'Torna Frame atual invisivel
    Frame1(iFrameAtual).Visible = False
    'Armazena novo valor de iFrameAtual
    iFrameAtual = TabStrip1.SelectedItem.Index

    'Se foi clicado no TAB_Selecao
    'If TabStrip1.SelectedItem.Index = TAB_Selecao Then iFrameSelecaoAlterado = 0

    'Se o frame anterior foi o de Sele��o e ele foi alterado
    If iFrameAtual <> TAB_Selecao And iFrameSelecaoAlterado = REGISTRO_ALTERADO Then

        'Se a categoria estiver preenchida => Tem que haver um item selecionado ...
        If Len(Trim(Categoria.Text)) > 0 Then
            For iIndiceItemCategoria = 0 To ItensCategoria.ListCount - 1
                If ItensCategoria.Selected(iIndiceItemCategoria) = True Then
                    bSelecionouitemCateg = True
                    Exit For
                End If
            Next
            
            'Se nao encontrou => Erro
            If Not bSelecionouitemCateg Then gError 108985
            
        End If
        
        Set gobjGeracaoPedCompraReq = New ClassGeracaoPedCompraReq
        Set gcolItemConcorrencia = New Collection

        'Limpa a sele��o atual
        Call Grid_Limpa(objGridRequisicoes)
        Call Grid_Limpa(objGridItensRequisicoes)
        Call Grid_Limpa(objGridProdutos1)
        Call Grid_Limpa(objGridProdutos2)
        Call Grid_Limpa(objGridCotacoes)
        Call Calcula_TotalItens

        'Recolhe os dados do TAB_Selecao
        lErro = Move_TabSelecao_Memoria(gobjGeracaoPedCompraReq)
        If lErro <> SUCESSO Then gError 63929

        'Busca no BD todas as Requisicoes de Compra com as caracter�sticas definidas no Tab Selecao
        lErro = CF("Requisicoes_Le_GeracaoPC", gobjGeracaoPedCompraReq)
        If lErro <> SUCESSO Then gError 63930

        'Traz os dados das requisicoes e seus itens para a tela
        lErro = Traz_Requisicoes_Tela(gobjGeracaoPedCompraReq)
        If lErro <> SUCESSO Then gError 63931

        iFrameSelecaoAlterado = 0

    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case 63929 To 63932
        
        Case 108985
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_PRODUTO_ITEM_NAO_PREENCHIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161201)

    End Select

    Exit Sub

End Sub

Private Sub Busca_Na_Colecao(collCodigoNome As AdmCollCodigoNome, lCodigo As Long, iPosicao As Integer)
'Busca a chave lCodigo na cole��o

Dim objlCodigoNome As AdmlCodigoNome
Dim iIndice As Integer

    iPosicao = 0
    iIndice = 0
    
    'Para cada item da cole��o
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

Private Function Traz_Requisicoes_Tela(gobjGeracaoPedCompraReq As ClassGeracaoPedCompraReq) As Long
'Preenche as Requisi��es a partir da cole��o passada

Dim lErro As Long
Dim objReqCompra As ClassRequisicaoCompras
Dim objItemRC As ClassItemReqCompras
Dim iIndice As Integer

On Error GoTo Erro_Traz_Requisicoes_Tela

    'L� os Itens das requisi��o de compra selcionadas
    lErro = CF("ItensReqCompra_Le_GeracaoPC", gobjGeracaoPedCompraReq)
    If lErro <> SUCESSO Then gError 62738

    'Preenche o grid de requisi��e scom dados das requisi��es
    lErro = GridRequisicoes_Preenche()
    If lErro <> SUCESSO Then gError 62745

    'Preenche o grid de itens de Requisi��o
    lErro = GridItensReq_Preenche()
    If lErro <> SUCESSO Then gError 62746
    
    'Para cada Requisi��o de compra lida
    For Each objReqCompra In gobjGeracaoPedCompraReq.colRequisicao
        'Para cada item de requisi��o lido
        For Each objItemRC In objReqCompra.colItens
            'Inclui ou altera um item de concorr�ncia incluindo os
            'dados do itemRC sem ler as cota��es
            lErro = ItensConcorrencia_Cria_Altera(objItemRC)
            If lErro <> SUCESSO Then gError 62747
        Next
                
    Next
        
    'Para cada item de concorr�ncia gerado
    For iIndice = 1 To gcolItemConcorrencia.Count
        'Busca as cotacoes para o item de concorr�ncia
        lErro = ItemConcorrencia_Atualiza_Cotacoes(gcolItemConcorrencia, iIndice)
        If lErro <> SUCESSO Then gError 62749
    Next
    
    'Preenche os grids de produto correspondentes aos itens de concorr�ncia
    lErro = Grids_Produto_Preenche()
    If lErro <> SUCESSO Then gError 62748
    
    Traz_Requisicoes_Tela = SUCESSO

    Exit Function

Erro_Traz_Requisicoes_Tela:

    Traz_Requisicoes_Tela = gErr

    Select Case gErr

        Case 62738, 63981, 62745, 62746, 62747, 62748, 62749
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161202)

    End Select

    Exit Function

End Function

Function GridItensReq_Preenche() As Long
'Preenche o GridItensRequisicoes com os Itens da Requisicao passada como parametro

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
Dim iCount As Integer

On Error GoTo Erro_GridItensReq_Preenche

    'Limpa o grid de itens
    Call Grid_Limpa(objGridItensRequisicoes)
    
    '#######################################################################
    'Inserido por Wagner 25/05/2006
    For Each objRequisicao In gobjGeracaoPedCompraReq.colRequisicao
        iCount = iCount + objRequisicao.colItens.Count
    Next
    
    If iCount >= objGridItensRequisicoes.objGrid.Rows Then
        Call Refaz_Grid(objGridItensRequisicoes, iCount)
    End If
    '#######################################################################
    
    'Para cada requisicao
    For Each objRequisicao In gobjGeracaoPedCompraReq.colRequisicao
        'Se a req est� selecionada
        If objRequisicao.iSelecionado = MARCADO Then
            'Para cada item
            For Each objItemReqCompras In objRequisicao.colItens
        
                iLinha = iLinha + 1
                'BUsca a filial da req na colfiliais
                Call Busca_Na_Colecao(colFiliais, objRequisicao.iFilialEmpresa, iPosicao)
            
                If iPosicao = 0 Then
               
                    objFilialEmpresa.iCodFilial = objRequisicao.iFilialEmpresa
                    'L� a FilialEmpresa
                    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
                    If lErro <> SUCESSO And lErro <> 27378 Then gError 68059
        
                    'Se n�o encontrou a filial ==>erro
                    If lErro = 27378 Then gError 68060
        
                    Set objlCodigoNome = New AdmlCodigoNome
                    
                    objlCodigoNome.lCodigo = objFilialEmpresa.iCodFilial
                    objlCodigoNome.sNome = objFilialEmpresa.sNome
                    
                    colFiliais.Add objlCodigoNome.lCodigo, objlCodigoNome.sNome
        
                Else
                    Set objlCodigoNome = colFiliais(iPosicao)
                End If
        
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_EscolhidoItem_Col) = objItemReqCompras.iSelecionado
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_FilialReqItem_Col) = objlCodigoNome.lCodigo & SEPARADOR & objlCodigoNome.sNome
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_CodigoReqItem_Col) = objRequisicao.lCodigo
        
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_Item_Col) = objItemReqCompras.iItem
        
                'Mascara o Produto
                lErro = Mascara_RetornaProdutoEnxuto(objItemReqCompras.sProduto, sProdutoMascarado)
                If lErro <> SUCESSO Then gError 68064
        
                ProdutoItemRC.PromptInclude = False
                ProdutoItemRC.Text = sProdutoMascarado
                ProdutoItemRC.PromptInclude = True
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_ProdutoItemRC_Col) = ProdutoItemRC.Text
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_DescProdutoItem_Col) = objItemReqCompras.sDescProduto
                
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_UnidadeMedItem_Col) = objItemReqCompras.sUM
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_QuantidadeItem_Col) = Formata_Estoque(objItemReqCompras.dQuantidade)
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_QuantPedida_Col) = Formata_Estoque(objItemReqCompras.dQuantPedida)
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_QuantRecebida_Col) = Formata_Estoque(objItemReqCompras.dQuantRecebida)
        
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_QuantComprarItem_Col) = Formata_Estoque(objItemReqCompras.dQuantComprar)
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_QuantComprarItem_Col) = Formata_Estoque(objItemReqCompras.dQuantidade - objItemReqCompras.dQuantRecebida - objItemReqCompras.dQuantPedida - objItemReqCompras.dQuantCancelada)
        
                If objItemReqCompras.iAlmoxarifado <> 0 Then
                    
                    Call Busca_Na_Colecao(colAlmoxarifados, objItemReqCompras.iAlmoxarifado, iPosicao)
                
                    If iPosicao = 0 Then
                
                        objAlmoxarifado.iCodigo = objItemReqCompras.iAlmoxarifado
            
                        'L� o almoxarifado
                        lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                        If lErro <> SUCESSO And lErro <> 25056 Then gError 63984
            
                        'Se n�o encontrou ==> Erro
                        If lErro = 25056 Then gError 63985
        
                        Set objlCodigoNome = New AdmlCodigoNome
                        
                        objlCodigoNome.lCodigo = objAlmoxarifado.iCodigo
                        objlCodigoNome.sNome = objAlmoxarifado.sNomeReduzido
                        
                        colAlmoxarifados.Add objlCodigoNome.lCodigo, objlCodigoNome.sNome
        
                    Else
                        Set objlCodigoNome = colAlmoxarifados(iPosicao)
                    End If
                
                    GridItensRequisicoes.TextMatrix(iLinha, iGrid_Almoxarifado_Col) = objlCodigoNome.sNome
                
                End If
        
        
                If Len(Trim(objItemReqCompras.sCcl)) > 0 Then
        
                    'Mascara o Ccl
                    lErro = Mascara_MascararCcl(objItemReqCompras.sCcl, sCclMascarado)
                    If lErro <> SUCESSO Then gError 63986
                    
                    GridItensRequisicoes.TextMatrix(iLinha, iGrid_CclItemRC_Col) = sCclMascarado
                End If
        
        
                If objItemReqCompras.lFornecedor <> 0 And objItemReqCompras.iFilial <> 0 Then
                    
                    Call Busca_Na_Colecao(colFornecedor, objItemReqCompras.lFornecedor, iPosicao)
        
                    If iPosicao = 0 Then
        
                        objFornecedor.lCodigo = objItemReqCompras.lFornecedor
            
                        'L� o Fornecedor
                        lErro = CF("Fornecedor_Le", objFornecedor)
                        If lErro <> SUCESSO And lErro <> 12729 Then gError 63987
            
                        'Se n�o encontrou o Fornecedor==> Erro
                        If lErro = 12729 Then gError 63988
                        
                        Set objlCodigoNome = New AdmlCodigoNome
                    
                        objlCodigoNome.lCodigo = objFornecedor.lCodigo
                        objlCodigoNome.sNome = objFornecedor.sNomeReduzido
                        
                        colFornecedor.Add objlCodigoNome.lCodigo, objlCodigoNome.sNome
                    
                    Else
                        Set objlCodigoNome = colFornecedor(iPosicao)
                    End If
        
                    GridItensRequisicoes.TextMatrix(iLinha, iGrid_FornecedorItemRC_Col) = objlCodigoNome.sNome
        
                    Call Busca_FilialForn(colFilialForn, objItemReqCompras.lFornecedor, objItemReqCompras.iFilial, iPosicao)
                    
                    If iPosicao = 0 Then
                        Set objFilialFornecedor = New ClassFilialFornecedor
                        objFilialFornecedor.iCodFilial = objItemReqCompras.iFilial
                        objFilialFornecedor.lCodFornecedor = objItemReqCompras.lFornecedor
                        
                        'L� a FilialFornecedor
                        lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
                        If lErro <> SUCESSO And lErro <> 12929 Then gError 63989
            
                        'Se n�o encontrou==>Erro
                        If lErro = 12929 Then gError 63990
                    Else
                        Set objFilialFornecedor = colFilialForn(iPosicao)
                    End If
        
                    GridItensRequisicoes.TextMatrix(iLinha, iGrid_FilialFornItemRC_Col) = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome
                
                    If objItemReqCompras.iExclusivo = MARCADO Then
                        GridItensRequisicoes.TextMatrix(iLinha, iGrid_ExclusivoItemRC_Col) = "Exclusivo"
                    Else
                        GridItensRequisicoes.TextMatrix(iLinha, iGrid_ExclusivoItemRC_Col) = "Preferencial"
                    End If
                    
                End If
        
                GridItensRequisicoes.TextMatrix(iLinha, iGrid_ObservacaoItemRC_Col) = objItemReqCompras.sObservacao
        
            Next
        End If
    Next
    
    'Atualiza o n�mero de linhas existentes do GridItensRequisicoes
    objGridItensRequisicoes.iLinhasExistentes = iLinha
    
    Call Grid_Refresh_Checkbox(objGridItensRequisicoes)
    
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161203)

    End Select

    Exit Function
    
End Function

Private Sub BotaoMarcarTodosItensRC_Click()
'Marca todas CheckBox do GridItensRequisicoes

Dim lErro As Long
Dim iItem As Integer
Dim iIndice As Integer
Dim objItemRC As ClassItemReqCompras
Dim colIndices As New Collection
Dim objReqCompra As ClassRequisicaoCompras
Dim objItemConc As New ClassItemConcorrencia

On Error GoTo Erro_BotaoMarcarTodosItensRC_Click
    
    'Para cada Req selecionada
    For Each objReqCompra In gobjGeracaoPedCompraReq.colRequisicao
        'se a req est� selecionada
        If objReqCompra.iSelecionado = MARCADO Then
            'marca os itens de requisicao
            For Each objItemRC In objReqCompra.colItens
                If objItemRC.iSelecionado = DESMARCADO Then
                    objItemRC.iSelecionado = MARCADO
                    
                    'Cria ou Altera os itens de concorrencia existentes
                    lErro = ItensConcorrencia_Cria_Altera(objItemRC)
                    If lErro <> SUCESSO Then gError 62757
                         
                    Call Localiza_ItemConcorrencia(gcolItemConcorrencia, objItemConc, iItem, objItemRC)
                    
                    Call Adiciona_Codigo(colIndices, iItem)
                    
                End If
            Next
        End If
    Next
    
    'Atualiza as cota��es
    For iIndice = 1 To colIndices.Count
        lErro = ItemConcorrencia_Atualiza_Cotacoes(gcolItemConcorrencia, colIndices(iIndice))
        If lErro <> SUCESSO Then gError 62766
    Next
    
    'seleciona no grid
    For iIndice = 1 To objGridItensRequisicoes.iLinhasExistentes
        GridItensRequisicoes.TextMatrix(iIndice, iGrid_EscolhidoItem_Col) = MARCADO
    Next
    
    Call Grid_Refresh_Checkbox(objGridItensRequisicoes)
    
    'Prenche o grid de produtos
    lErro = Grids_Produto_Preenche()
    If lErro <> SUCESSO Then gError 62758
    
    Exit Sub

Erro_BotaoMarcarTodosItensRC_Click:

    Select Case gErr

        Case 62766, 62757, 62758

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161204)

    End Select

    Exit Sub

End Sub

Private Sub BotaoMarcarTodosProd_Click()
'Marca todas CheckBox do GridProdutos1

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_BotaoMarcarTodosProd_Click

    'Marca todos os Itens do GridProdutos1
    For iIndice = 1 To objGridProdutos1.iLinhasExistentes
        GridProdutos1.TextMatrix(iIndice, iGrid_EscolhidoProduto_Col) = GRID_CHECKBOX_ATIVO
        gcolItemConcorrencia(iIndice).iEscolhido = MARCADO
    Next

    Call Grid_Refresh_Checkbox(objGridProdutos1)
    
    'Preenche o grid de produtos
    lErro = GridProdutos2_Preenche()
    If lErro <> SUCESSO Then gError 62759

    'Preenche o grid de cota��es
    lErro = GridCotacoes_Preenche()
    If lErro <> SUCESSO Then gError 62760

    Exit Sub

Erro_BotaoMarcarTodosProd_Click:

    Select Case gErr

        Case 62759, 62760

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161205)

    End Select

    Exit Sub

End Sub

Private Sub BotaoMarcarTodosReq_Click()
'Marca todas CheckBox do GridRequisicoes

Dim lErro As Long
Dim iItem As Integer
Dim iLinha As Integer
Dim iIndice As Integer
Dim colItens As New Collection
Dim objItemRC As ClassItemReqCompras
Dim objItemConc As New ClassItemConcorrencia
Dim objReqCompras As ClassRequisicaoCompras

On Error GoTo Erro_BotaoMarcarTodosReq_Click
    
    Set gcolItemConcorrencia = New Collection
    
    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridRequisicoes.iLinhasExistentes
        
        'Marca na tela a linha em quest�o
        GridRequisicoes.TextMatrix(iLinha, iGrid_EscolhidoReq_Col) = GRID_CHECKBOX_ATIVO
        Set objReqCompras = gobjGeracaoPedCompraReq.colRequisicao(iLinha)
        objReqCompras.iSelecionado = MARCADO
        
        'Para cada Item
        For Each objItemRC In objReqCompras.colItens
            'Seleciona o item
            objItemRC.iSelecionado = True

            lErro = ItensConcorrencia_Cria_Altera(objItemRC)
            If lErro <> SUCESSO Then gError 62752
        
            Call Localiza_ItemConcorrencia(gcolItemConcorrencia, objItemConc, iItem, objItemRC)
            
            Call Adiciona_Codigo(colItens, iItem)
        
        Next

    Next
    
    'ATualiza as cota��es
    For iIndice = 1 To colItens.Count
       lErro = ItemConcorrencia_Atualiza_Cotacoes(gcolItemConcorrencia, colItens(iIndice))
       If lErro <> SUCESSO Then gError 62767
    Next
    
    'Atualiza na tela a checkbox marcada
    Call Grid_Refresh_Checkbox(objGridRequisicoes)
    
    'Preenche o grid de itens
    lErro = GridItensReq_Preenche()
    If lErro <> SUCESSO Then gError 62753
    
    'Preenche o grid de Produtos
    lErro = Grids_Produto_Preenche()
    If lErro <> SUCESSO Then gError 62754
    
    Exit Sub
    
Erro_BotaoMarcarTodosReq_Click:

    Select Case gErr
    
        Case 62752, 62753, 62754, 62767
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161206)
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoPedCotacao_Click()
'Chama a tela PedidoCotacaoLista

Dim objPedidoCotacao As New ClassPedidoCotacao
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoPedCotacao_Click

    'Verifica se existe alguma linha selecionada no GridCotacoes
    If GridCotacoes.Row = 0 Then gError 89433

    objPedidoCotacao.lCodigo = StrParaLong(GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_PedidoCot_Col))
    objPedidoCotacao.iFilialEmpresa = giFilialEmpresa

    'Chama a tela PedidoCotacao
    Call Chama_Tela("PedidoCotacao", objPedidoCotacao)

    Exit Sub
    
Erro_BotaoPedCotacao_Click:

    Select Case gErr

        Case 89433
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161207)

    End Select

    Exit Sub

End Sub

Private Sub EscolhidoReq_Click()

    iAlterado = REGISTRO_ALTERADO
    Call Requisicoes_Atualiza
End Sub

Private Sub FilialEmpresa_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialFornec_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
    iFrameSelecaoAlterado = REGISTRO_ALTERADO

End Sub


Private Sub FilialReq_Change()

    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Fornecedor_Change()

    iFornecedorAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
    iFrameSelecaoAlterado = REGISTRO_ALTERADO

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

    'Verifica se � uma filial selecionada
    If FilialFornec.ListIndex >= 0 Then Exit Sub

    'Tenta selecionar na combo de FilialFornec
    lErro = Combo_Seleciona(FilialFornec, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 63894

    'Se nao encontra o �tem com o c�digo informado
    If lErro = 6730 Then

        'Verifica de o fornecedor foi digitado
        If Len(Trim(Fornecedor.Text)) = 0 Then gError 63895

        sFornecedor = Fornecedor.Text

        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError 63896

        'Se nao existir
        If lErro = 18272 Then

            objFornecedor.sNomeReduzido = sFornecedor

            'Le o C�digo do Fornecedor --> Para Passar para a Tela de Filiais
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then gError 63897

            'Passa o C�digo do Fornecedor
            objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo

            'Sugere cadastrar nova Filial
            gError 63898

        End If

        'Coloca na tela o c�digo e o nome da FilialForn
        FilialFornec.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

    End If

    'N�o encontrou valor informado que era STRING
    If lErro = 6731 Then gError 63899

    Exit Sub

Erro_FilialFornec_Validate:

    Cancel = True

    Select Case gErr

        Case 63894, 63896, 63897 'Tratados nas Rotinas chamadas

        Case 63898
            'Pergunta se deseja criar nova filial para o fornecedor em questao
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

            If vbMsgRes = vbYes Then
                'Chama a tela FiliaisFornecedores
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 63895
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 63899
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", gErr, FilialFornec.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161208)

    End Select

    Exit Sub


End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Validate

    'Verifica se Fornecedor foi alterado
    If iFornecedorAlterado = 0 Then Exit Sub

    'Verifica se o Fornecedor esta preenchido
    If Len(Trim(Fornecedor.Text)) > 0 Then

        'Le o Fornecedor
        lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
        If lErro <> SUCESSO Then gError 63892

        'Le as Filiais do Fornecedor
        lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
        If lErro <> SUCESSO And lErro <> 6698 Then gError 63893

        'Preenche a combo FilialForn
        Call CF("Filial_Preenche", FilialFornec, colCodigoNome)

        'Seleciona a filial na combo de FilialForn
        Call CF("Filial_Seleciona", FilialFornec, iCodFilial)

    End If

    'Se o Fornecedor nao estiver preenchido
    If Len(Trim(Fornecedor.Text)) = 0 Then

        'Limpa a combo FilialFornec
        FilialFornec.Clear

    End If

    iFornecedorAlterado = 0

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True

    Select Case gErr

        Case 63892, 63893
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161209)

    End Select

    Exit Sub

End Sub

Private Sub GridProdutos1_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridProdutos1, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridProdutos1, iAlterado)
    End If

End Sub

Private Sub GridProdutos1_GotFocus()
    Call Grid_Recebe_Foco(objGridProdutos1)
End Sub

Private Sub GridProdutos1_EnterCell()
    Call Grid_Entrada_Celula(objGridProdutos1, iAlterado)
End Sub

Private Sub GridProdutos1_LeaveCell()
    Call Saida_Celula(objGridProdutos1)
End Sub

Private Sub GridProdutos1_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridProdutos1)
    
End Sub

Private Sub GridProdutos1_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Trata_Tecla(KeyAscii, objGridProdutos1, iExecutaEntradaCelula)

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

Private Sub GridProdutos2_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridProdutos2, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridProdutos2, iAlterado)
    End If

End Sub

Private Sub GridProdutos2_GotFocus()
    Call Grid_Recebe_Foco(objGridProdutos2)
End Sub

Private Sub GridProdutos2_EnterCell()
    Call Grid_Entrada_Celula(objGridProdutos2, iAlterado)
End Sub

Private Sub GridProdutos2_LeaveCell()
    Call Saida_Celula(objGridProdutos2)
End Sub

Private Sub GridProdutos2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridProdutos2)
End Sub

Private Sub GridProdutos2_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridProdutos2, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridProdutos2, iAlterado)
    End If

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

Private Sub GridCotacoes_GotFocus()
    Call Grid_Recebe_Foco(objGridCotacoes)
End Sub

Private Sub GridCotacoes_EnterCell()
    Call Grid_Entrada_Celula(objGridCotacoes, iAlterado)
End Sub

Private Sub GridCotacoes_LeaveCell()
    Call Saida_Celula(objGridCotacoes)
End Sub

Private Sub GridCotacoes_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call Grid_Trata_Tecla1(KeyCode, objGridCotacoes)

End Sub

Private Sub GridCotacoes_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Trata_Tecla(KeyAscii, objGridCotacoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCotacoes, iAlterado)
    End If

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

Private Sub GridRequisicoes_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridRequisicoes, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRequisicoes, iAlterado)
    End If
    
    Exit Sub

End Sub

Private Sub GridRequisicoes_GotFocus()
    Call Grid_Recebe_Foco(objGridRequisicoes)
End Sub

Private Sub GridRequisicoes_EnterCell()
    Call Grid_Entrada_Celula(objGridRequisicoes, iAlterado)
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
        Call Grid_Entrada_Celula(objGridRequisicoes, iAlterado)
    End If
    
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

Private Sub GridItensRequisicoes_Click()

Dim iExecutaEntradaCelula As Integer
Dim lErro As Long

On Error GoTo Erro_GridItensRequisicoes_Click

    Call Grid_Click(objGridItensRequisicoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItensRequisicoes, iAlterado)
    End If
    
    Exit Sub

Erro_GridItensRequisicoes_Click:

    Select Case gErr

        Case 62798

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161210)

    End Select

    Exit Sub

End Sub

Private Sub GridItensRequisicoes_GotFocus()
    Call Grid_Recebe_Foco(objGridItensRequisicoes)
End Sub

Private Sub GridItensRequisicoes_EnterCell()
    Call Grid_Entrada_Celula(objGridItensRequisicoes, iAlterado)
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
        Call Grid_Entrada_Celula(objGridItensRequisicoes, iAlterado)
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

Private Sub objEventoBotaoPedCotacao_evSelecao(obj1 As Object)
    Me.Show
End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor
Dim bCancel As Boolean

    Set objFornecedor = obj1

    'Coloca o nome reduzido do Fornecedor na tela
    Fornecedor.Text = objFornecedor.sNomeReduzido

    Fornecedor_Validate (bCancel)

    Me.Show

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da c�lula do grid que est� deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'Verifica qual o Grid em quest�o
        Select Case objGridInt.objGrid.Name

            'Se for o GridRequisicoes
            Case GridRequisicoes.Name

                lErro = Saida_Celula_GridRequisicoes(objGridInt)
                If lErro <> SUCESSO Then gError 63934

            'Se for o GridProdutos1
            Case GridProdutos1.Name

                lErro = Saida_Celula_GridProdutos1(objGridInt)
                If lErro <> SUCESSO Then gError 63935

            'Se for o GridProdutos2
            Case GridProdutos2.Name

                lErro = Saida_Celula_GridProdutos2(objGridInt)
                If lErro <> SUCESSO Then gError 63936

            'Se for o GridCotacoes
            Case GridCotacoes.Name

                lErro = Saida_Celula_GridCotacoes(objGridInt)
                If lErro <> SUCESSO Then gError 63937

           'Se for o GridItensRequisicoes
            Case GridItensRequisicoes.Name

                lErro = Saida_Celula_GridItensRequisicoes(objGridInt)
                If lErro <> SUCESSO Then gError 63938


        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 63939

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 63934 To 63939
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161211)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridProdutos1(objGridInt As AdmGrid) As Long
'Faz a critica da celula do GridProdutos1 que esta deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridProdutos1

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'EscolhidoProduto
        Case iGrid_EscolhidoProduto_Col
            lErro = Saida_Celula_EscolhidoProduto(objGridInt)
            If lErro <> SUCESSO Then gError 63940

    End Select

    Saida_Celula_GridProdutos1 = SUCESSO

    Exit Function

Erro_Saida_Celula_GridProdutos1:

    Saida_Celula_GridProdutos1 = gErr

    Select Case gErr

        Case 63940

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161212)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridProdutos2(objGridInt As AdmGrid) As Long
'Faz a critica da celula do GridProdutos2 que esta deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridProdutos2

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'EscolhidoProduto
        Case iGrid_Quantidade2_Col
            lErro = Saida_Celula_Quantidade2(objGridInt)
            If lErro <> SUCESSO Then gError 63967

    End Select

    Saida_Celula_GridProdutos2 = SUCESSO

    Exit Function

Erro_Saida_Celula_GridProdutos2:

    Saida_Celula_GridProdutos2 = gErr

    Select Case gErr

        Case 63967

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161213)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridCotacoes(objGridInt As AdmGrid) As Long
'Faz a critica da celula do GridCotacoes que esta deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridCotacoes

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'EscolhidoCot
        Case iGrid_EscolhidoCot_Col
            lErro = Saida_Celula_EscolhidoCot(objGridInt)
            If lErro <> SUCESSO Then gError 63945

        'QuantComprarCot
        Case iGrid_QuantComprarCot_Col
            lErro = Saida_Celula_QuantComprarCot(objGridInt)
            If lErro <> SUCESSO Then gError 63946

        'Pre�o Unit�rio
        Case iGrid_PrecoUnitarioCot_Col
            lErro = Saida_Celula_PrecoUnitarioCot(objGridInt)
            If lErro <> SUCESSO Then gError 70459

        'MotivoEscolhaCot
        Case iGrid_MotivoEscolhaCot_Col
            lErro = Saida_Celula_MotivoEscolhaCot(objGridInt)
            If lErro <> SUCESSO Then gError 63947

    End Select

    Saida_Celula_GridCotacoes = SUCESSO

    Exit Function

Erro_Saida_Celula_GridCotacoes:

    Saida_Celula_GridCotacoes = gErr

    Select Case gErr

        Case 63945, 63946, 63947, 70459

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161214)

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
            If lErro <> SUCESSO Then gError 63942

        'QuantComprarItem
        Case iGrid_QuantComprarItem_Col
            lErro = Saida_Celula_QuantComprarItemReq(objGridInt)
            If lErro <> SUCESSO Then gError 63943

    End Select

    Saida_Celula_GridItensRequisicoes = SUCESSO

    Exit Function

Erro_Saida_Celula_GridItensRequisicoes:

    Saida_Celula_GridItensRequisicoes = gErr

    Select Case gErr

        Case 63942, 63943

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161215)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridRequisicoes(objGridInt As AdmGrid) As Long
'Faz a critica da celula do GridRequisicoes que esta deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridRequisicoes

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'EscolhidoReq
        Case iGrid_EscolhidoReq_Col
            lErro = Saida_Celula_EscolhidoReq(objGridInt)
            If lErro <> SUCESSO Then gError 63961

    End Select

    Saida_Celula_GridRequisicoes = SUCESSO

    Exit Function

Erro_Saida_Celula_GridRequisicoes:

    Saida_Celula_GridRequisicoes = gErr

    Select Case gErr

        Case 63961

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161216)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_EscolhidoProduto(objGridInt As AdmGrid) As Long
'Faz a saida de c�lula de EscolhidoProduto

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_EscolhidoProduto

    Set objGridInt.objControle = EscolhidoProduto

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 63941

    Exit Function

Erro_Saida_Celula_EscolhidoProduto:

    Select Case gErr

        Case 63941
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161217)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_EscolhidoCot(objGridInt As AdmGrid) As Long
'Faz a saida de c�lula de EscolhidoCot

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_EscolhidoCot

    Set objGridInt.objControle = EscolhidoCot

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 63948

    Exit Function

Erro_Saida_Celula_EscolhidoCot:

    Select Case gErr

        Case 63948
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161218)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MotivoEscolhaCot(objGridInt As AdmGrid) As Long
'Faz a saida de celula de MotivoEscolha

Dim lErro As Long
Dim iCodigo As Integer
Dim objCotItemConc As ClassCotacaoItemConc

On Error GoTo Erro_Saida_Celula_MotivoEscolhaCot

    Set objGridInt.objControle = MotivoEscolhaCot

    'Verifica se o MotivoEscolhaCot est� preenchido
    If Len(Trim(MotivoEscolhaCot.Text)) > 0 Then

        'Verifica se MotivoEscolhaCot n�o est� selecionado
        If MotivoEscolhaCot.ListIndex = -1 Then
                        
            If UCase(MotivoEscolhaCot.Text) = UCase(MOTIVO_EXCLUSIVO_DESCRICAO) Then gError 62715
            
            'Seleciona o MotivoEscolhaCot na combobox
            lErro = Combo_Item_Seleciona(MotivoEscolhaCot)
            If lErro <> SUCESSO And lErro <> 12250 Then gError 63741

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 63743

    Call Localiza_ItemCotacao(objCotItemConc, GridCotacoes.Row)

    objCotItemConc.sMotivoEscolha = GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_MotivoEscolhaCot_Col)

    Saida_Celula_MotivoEscolhaCot = SUCESSO

    Exit Function

Erro_Saida_Celula_MotivoEscolhaCot:

    Saida_Celula_MotivoEscolhaCot = gErr

    Select Case gErr

        Case 62715
            Call Rotina_Erro(vbOKOnly, "ERRO_MOTIVO_EXCLUSIVO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 63741, 63743
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161219)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantComprarCot(objGridInt As AdmGrid) As Long
'Faz a saida de celula de QuantComprarCot

Dim lErro As Long
Dim dQuantidade As Double
Dim objCotItemConc As ClassCotacaoItemConc

On Error GoTo Erro_Saida_Celula_QuantComprarCot

     Set objGridInt.objControle = QuantComprarCot
    
    'Verifica se a QuantComprarCot esta preenchida
    If Len(Trim(QuantComprarCot.ClipText)) > 0 Then

        'Critica a quantidade
        lErro = Valor_Positivo_Critica(QuantComprarCot.Text)
        If lErro <> SUCESSO Then gError 63739

        dQuantidade = StrParaDbl(QuantComprarCot.Text)

        'Coloca a quantidade com o formato de estoque da tela
         QuantComprarCot.Text = Formata_Estoque(dQuantidade)

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 63740
    
    'Localiza o ItemCotacao selecionado
    Call Localiza_ItemCotacao(objCotItemConc, GridCotacoes.Row)
    
    'Atualiza a quantidade a comprar
    objCotItemConc.dQuantidadeComprar = dQuantidade
    'Atualiza o valor do item
    GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_ValorItem_Col) = Format(objCotItemConc.dPrecoAjustado * objCotItemConc.dQuantidadeComprar, "STANDARD")
    
    'recalcula o total
    Call Calcula_TotalItens
    
    Saida_Celula_QuantComprarCot = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantComprarCot:

    Saida_Celula_QuantComprarCot = gErr

    Select Case gErr

        Case 63739, 63740
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161220)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PrecoUnitarioCot(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objCotItemConc As New ClassCotacaoItemConc
Dim dValorPresente As Double
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_Saida_Celula_PrecoUnitarioCot

    Set objGridInt.objControle = PrecoUnitario

    'Se o Pre�o unit�rio estiver preenchido
    If Len(Trim(PrecoUnitario.ClipText)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(PrecoUnitario.Text)
        If lErro <> SUCESSO Then gError 70482

    End If
        
    Call Localiza_ItemCotacao(objCotItemConc, GridCotacoes.Row)
    
    objCotItemConc.dPrecoAjustado = StrParaDbl(PrecoUnitario.Text)
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 70483

    'Se a condi��o de pagamento n�o for a vista
    If Codigo_Extrai(GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_CondPagtoCot_Col)) <> COD_A_VISTA And PercentParaDbl(TaxaEmpresa.Caption) > 0 Then
        
        objCondicaoPagto.iCodigo = Codigo_Extrai(objCotItemConc.sCondPagto)
        
        'Recalcula o Valor Presente
        lErro = CF("Calcula_ValorPresente", objCondicaoPagto, objCotItemConc.dPrecoAjustado, PercentParaDbl(TaxaEmpresa.Caption), dValorPresente, gdtDataAtual)
        If lErro <> SUCESSO Then gError 62736
        
        If objCotItemConc.iMoeda <> MOEDA_REAL Then
        
            GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_ValorPresenteCot_Col) = Format(dValorPresente * objCotItemConc.dTaxa, ValorPresente.Format) 'Alterado por Wagner
            objCotItemConc.dValorPresente = dValorPresente * objCotItemConc.dTaxa
            
        Else
        
            GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_ValorPresenteCot_Col) = Format(dValorPresente, ValorPresente.Format) 'Alterado por Wagner
            objCotItemConc.dValorPresente = dValorPresente
            
        End If
        
    ElseIf Codigo_Extrai(GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_CondPagtoCot_Col)) = COD_A_VISTA Then
        
        If objCotItemConc.iMoeda <> MOEDA_REAL Then
            GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_ValorPresenteCot_Col) = Format((StrParaDbl(PrecoUnitario.Text)) * objCotItemConc.dTaxa, ValorPresente.Format) 'Alterado por Wagner
            objCotItemConc.dValorPresente = dValorPresente * objCotItemConc.dTaxa
        Else
            GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_ValorPresenteCot_Col) = Format((StrParaDbl(PrecoUnitario.Text)), ValorPresente.Format) 'Alterado por Wagner
            objCotItemConc.dValorPresente = dValorPresente
        End If
        
    End If
    
    If objCotItemConc.iMoeda <> MOEDA_REAL Then
        'Atualiza o valor desse item alterado
        GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_ValorItem_Col) = Format(objCotItemConc.dPrecoAjustado * objCotItemConc.dQuantidadeComprar * objCotItemConc.dTaxa, "STANDARD")
    Else
        'Atualiza o valor desse item alterado
        GridCotacoes.TextMatrix(GridCotacoes.Row, iGrid_ValorItem_Col) = Format(objCotItemConc.dPrecoAjustado * objCotItemConc.dQuantidadeComprar, "STANDARD")
    End If
    
    'Atuliza o valor dos itens selecionados
    Call Calcula_TotalItens
    
    Saida_Celula_PrecoUnitarioCot = SUCESSO

    Exit Function

Erro_Saida_Celula_PrecoUnitarioCot:

    Saida_Celula_PrecoUnitarioCot = gErr

    Select Case gErr

        Case 62736, 70482, 70483
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161221)

    End Select

    Exit Function

End Function


Private Function Saida_Celula_EscolhidoItem(objGridInt As AdmGrid) As Long
'Faz a saida de c�lula de EscolhidoItem

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_EscolhidoItem

    Set objGridInt.objControle = EscolhidoItem

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 63944

    Exit Function

Erro_Saida_Celula_EscolhidoItem:

    Select Case gErr

        Case 63944
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161222)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantComprarItemReq(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objReqCompra As ClassRequisicaoCompras
Dim objItemRC As ClassItemReqCompras
Dim dQuantPosterior As Double
Dim dQuantAnterior As Double
Dim iIndice1 As Integer, iItem As Integer
Dim iIndice2 As Integer
Dim bAchou As Boolean, objProduto As New ClassProduto
Dim dQuantDiferenca As Double, dFator As Double
Dim objItemConcorrencia As ClassItemConcorrencia

On Error GoTo Erro_Saida_Celula_QuantComprarItemReq

    Set objGridInt.objControle = QuantComprarItemRC
    
    'Guarda a quantidade anterior do grid
    dQuantAnterior = StrParaDbl(GridItensRequisicoes.TextMatrix(GridItensRequisicoes.Row, iGrid_QuantComprarItem_Col))

    'Se quantidade estiver preenchida
    If Len(Trim(QuantComprarItemRC.ClipText)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(QuantComprarItemRC.Text)
        If lErro <> SUCESSO Then gError 63964
        
        'Guarda a qt alterada
        dQuantPosterior = StrParaDbl(QuantComprarItemRC.Text)

    Else
        gError 62799
    End If
    
    'Calula a diferen�a entre a quant anterior e a atual
    dQuantDiferenca = Round(dQuantPosterior - dQuantAnterior, 4)
        
    'Se houve altera��o na quantidade
    If dQuantDiferenca <> 0 Then
        
        'Localiza o item e a requisi��o da linha selecionada
        For iIndice1 = 1 To gobjGeracaoPedCompraReq.colRequisicao.Count
            Set objReqCompra = gobjGeracaoPedCompraReq.colRequisicao(iIndice1)
            
            For iIndice2 = 1 To objReqCompra.colItens.Count
                
                Set objItemRC = objReqCompra.colItens(iIndice2)
                
                If objItemRC.iItem = StrParaInt(GridItensRequisicoes.TextMatrix(GridItensRequisicoes.Row, iGrid_Item_Col)) And _
                   objReqCompra.lCodigo = StrParaLong(GridItensRequisicoes.TextMatrix(GridItensRequisicoes.Row, iGrid_CodigoReqItem_Col)) And _
                   objReqCompra.iFilialEmpresa = Codigo_Extrai(GridItensRequisicoes.TextMatrix(GridItensRequisicoes.Row, iGrid_FilialReqItem_Col)) Then
                    'Achou
                    bAchou = True
                    Exit For
                End If
            Next
            'Se j� achou --> sai
            If bAchou Then Exit For
        Next
        
        
        'Verifica se a quantidade digitada � maior que a quant que falta comprar do itemrc
        If dQuantPosterior > objItemRC.dQuantidade - objItemRC.dQuantCancelada - objItemRC.dQuantPedida - objItemRC.dQuantRecebida Then gError 63965
        
        'Localiza o ItemConcorr�ncia vinculado ao Item RC
        Call Localiza_ItemConcorrencia(gcolItemConcorrencia, objItemConcorrencia, iItem, objItemRC)
        
        objProduto.sCodigo = objItemConcorrencia.sProduto
        
        'L� o produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 23080 Then gError 62756
        If lErro <> SUCESSO Then gError 62757
        
        lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemRC.sUM, objProduto.sSiglaUMCompra, dFator)
        If lErro <> SUCESSO Then gError 62758
        
        'Converte a quantidade p\ UM de compra
        dQuantDiferenca = dQuantDiferenca * dFator
        
        objItemRC.dQuantComprar = dQuantPosterior
        objItemRC.dQuantNaConcorrencia = objItemRC.dQuantComprar * dFator
                
        'Se a quantidade foi aumentada
        If dQuantDiferenca > 0 Then
            'Aumenta a quantidade do item de concorr�ncia
            lErro = ItemConcorrencia_Inclui_QuantComprar(objItemConcorrencia, iItem, objReqCompra, objItemRC, dQuantDiferenca)
            If lErro <> SUCESSO Then gError 62759
            
        'Se a quantidade foi diminuida
        ElseIf iItem > 0 Then
        
            'Diminui a quantidade no item de concorr�ncia
            lErro = ItemConcorrencia_Exclui_QuantComprar(objItemConcorrencia, iItem, objReqCompra, objItemRC, Abs(dQuantDiferenca))
            If lErro <> SUCESSO Then gError 62760
            
        End If
        
        'Atualiza as cota��es
        lErro = ItemConcorrencia_Atualiza_Cotacoes(gcolItemConcorrencia, iItem)
        If lErro <> SUCESSO Then gError 62771
            
        'Preenche o grid de produtos
        lErro = Grids_Produto_Preenche()
        If lErro <> SUCESSO Then gError 62761
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 63966
    
    Saida_Celula_QuantComprarItemReq = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantComprarItemReq:

    Saida_Celula_QuantComprarItemReq = gErr

    Select Case gErr

        Case 62756, 63964, 63966, 62758, 62759, 62760, 62761, 62771
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 62757
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 62799
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTCOMPRAR_NAO_PREENCHIDA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 63965
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTCOMPRAR_SUPERIOR_MAXIMA", gErr, dQuantPosterior, objItemRC.dQuantidade - objItemRC.dQuantCancelada - objItemRC.dQuantPedida - objItemRC.dQuantRecebida)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161223)

    End Select

    Exit Function

End Function

Sub Atualiza_GridCotacoes(iLinhaGrid As Integer, dQuantDiferenca As Double)
'Atualiza o Grid de Cota��es (quantComprar e itens selecionados)

Dim sProduto As String
Dim sFornecedor As String
Dim iFilial As Integer
Dim iIndice As Integer

    sProduto = GridItensRequisicoes.TextMatrix(iLinhaGrid, iGrid_ProdutoItemRC_Col)
    sFornecedor = GridItensRequisicoes.TextMatrix(iLinhaGrid, iGrid_FornecedorItemRC_Col)
    iFilial = Codigo_Extrai(GridItensRequisicoes.TextMatrix(iLinhaGrid, iGrid_FilialFornItemRC_Col))

    'Procura no GridCota��es o mesmo Produto
    For iIndice = 1 To objGridCotacoes.iLinhasExistentes

        'Se achou
        If sProduto = GridCotacoes.TextMatrix(iIndice, iGrid_ProdutoCot_Col) And _
           sFornecedor = GridCotacoes.TextMatrix(iIndice, iGrid_FornecedorCot_Col) And _
           iFilial = Codigo_Extrai(GridCotacoes.TextMatrix(iIndice, iGrid_FilialFornCot_Col)) Then

            'Se a quantidade aumentou
            If dQuantDiferenca > 0 Then

                'Se a quantidade atual mais a quantidade que aumentou for menor ou igual a m�xima a comprar
                If StrParaDbl(GridCotacoes.TextMatrix(iIndice, iGrid_QuantComprarCot_Col)) + dQuantDiferenca <= StrParaDbl(GridCotacoes.TextMatrix(iIndice, iGrid_QuantidadeCot_Col)) Then

                    'Atualiza a nova quantidade a comprar somando o que aumentou
                    GridCotacoes.TextMatrix(iIndice, iGrid_QuantComprarCot_Col) = Formata_Estoque(StrParaDbl(GridCotacoes.TextMatrix(iIndice, iGrid_QuantComprarCot_Col)) + dQuantDiferenca)

                    'Marca linha do GridCota��es
                    GridCotacoes.TextMatrix(iIndice, iGrid_EscolhidoCot_Col) = GRID_CHECKBOX_ATIVO

                    Exit For

                'Se n�o
                Else

                    'Guarda nova diferen�a de quantidade
                    dQuantDiferenca = dQuantDiferenca - (StrParaDbl(GridCotacoes.TextMatrix(iIndice, iGrid_QuantidadeCot_Col)) - StrParaDbl(GridCotacoes.TextMatrix(iIndice, iGrid_QuantComprarCot_Col)))
                    GridCotacoes.TextMatrix(iIndice, iGrid_QuantComprarCot_Col) = GridCotacoes.TextMatrix(iIndice, iGrid_QuantidadeCot_Col)

                    'Marca linha do GridCota��es
                    GridCotacoes.TextMatrix(iIndice, iGrid_EscolhidoCot_Col) = GRID_CHECKBOX_ATIVO

                End If

            'Se a quantidade diminuiu
            Else

                'Se a quantidade atual menos a quantidade que diminuiu for maior ou igual a zero
                If StrParaDbl(GridCotacoes.TextMatrix(iIndice, iGrid_QuantComprarCot_Col)) + dQuantDiferenca >= 0 Then

                    'Atualiza a nova quantidade a comprar somando o que aumentou
                    GridCotacoes.TextMatrix(iIndice, iGrid_QuantComprarCot_Col) = Formata_Estoque(StrParaDbl(GridCotacoes.TextMatrix(iIndice, iGrid_QuantComprarCot_Col)) + dQuantDiferenca)

                    'Se zerou a quantidade a comprar
                    If StrParaDbl(GridCotacoes.TextMatrix(iIndice, iGrid_QuantComprarCot_Col)) = 0 Then

                        'desmarca linha do GridCota��es
                        GridCotacoes.TextMatrix(iIndice, iGrid_EscolhidoCot_Col) = "0"
                    End If

                    Exit For

                'Se n�o
                Else

                    dQuantDiferenca = dQuantDiferenca + StrParaDbl(GridCotacoes.TextMatrix(iIndice, iGrid_QuantComprarCot_Col))
                    GridCotacoes.TextMatrix(iIndice, iGrid_QuantComprarCot_Col) = Formata_Estoque(0)

                    'desmarca linha do GridCota��es
                    GridCotacoes.TextMatrix(iIndice, iGrid_EscolhidoCot_Col) = GRID_CHECKBOX_INATIVO

                End If

            End If

        End If

    Next

    Call Grid_Refresh_Checkbox(objGridCotacoes)

End Sub

Private Function Saida_Celula_Quantidade2(objGridInt As AdmGrid) As Long

Dim lErro As Long, dQuantidade As Double
Dim iIndice As Integer, dQuantTotalRC As Double
Dim sFornecedor As String, iFilial As Integer
Dim sProduto As String, dQuantAnterior As Double
Dim dQuantDiferenca As Double, iItem As Integer

On Error GoTo Erro_Saida_Celula_Quantidade2

    Set objGridInt.objControle = Quantidade2
    
    dQuantAnterior = StrParaDbl(GridProdutos2.TextMatrix(GridProdutos2.Row, iGrid_Quantidade2_Col))

    'Se quantidade estiver preenchida
    If Len(Trim(Quantidade2.ClipText)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(Quantidade2.Text)
        If lErro <> SUCESSO Then gError 63963

        dQuantidade = CDbl(Quantidade2.Text)

        'Coloca o valor Formatado na tela
        Quantidade2.Text = Formata_Estoque(dQuantidade)
    Else
        gError 62744
    End If

    'Calcula a diferen�a entre a quant anterior e essa
    dQuantDiferenca = StrParaDbl(Formata_Estoque(dQuantidade - dQuantAnterior))
    
    'Guarda campos da linha em quest�o de GridProdutos2
    sProduto = GridProdutos2.TextMatrix(GridProdutos2.Row, iGrid_Produto2_Col)
    sFornecedor = GridProdutos2.TextMatrix(GridProdutos2.Row, iGrid_Fornecedor2_Col)
    iFilial = Codigo_Extrai(GridProdutos2.TextMatrix(GridProdutos2.Row, iGrid_FilialForn2_Col))

    'Atualiza o valor da cole��o de qt suplementares
    ' e verifica se a qt digitada � < que a qt dos itens req
    For iIndice = 1 To objGridProdutos1.iLinhasExistentes
        If sProduto = GridProdutos1.TextMatrix(iIndice, iGrid_Produto1_Col) And sFornecedor = GridProdutos1.TextMatrix(iIndice, iGrid_Fornecedor1_Col) And iFilial = Codigo_Extrai(GridProdutos1.TextMatrix(iIndice, iGrid_FilialForn1_Col)) Then
            lErro = Atualiza_QuantSupl(gcolItemConcorrencia(iIndice), dQuantDiferenca, GridProdutos2.Row)
            If lErro <> SUCESSO Then gError 63965
            Exit For
        End If
    Next

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 63964

    'Se a quant foi alterada
    If dQuantDiferenca <> 0 Then
    
        'Atualiza a quantidade a comprar no GridProdutos1
        For iIndice = 1 To objGridProdutos1.iLinhasExistentes
            If sProduto = GridProdutos1.TextMatrix(iIndice, iGrid_Produto1_Col) And sFornecedor = GridProdutos1.TextMatrix(iIndice, iGrid_Fornecedor1_Col) And iFilial = Codigo_Extrai(GridProdutos1.TextMatrix(iIndice, iGrid_FilialForn1_Col)) Then
                GridProdutos1.TextMatrix(iIndice, iGrid_Quantidade1_Col) = Formata_Estoque(StrParaDbl(GridProdutos1.TextMatrix(iIndice, iGrid_Quantidade1_Col)) + dQuantDiferenca)
                
                'Se a qt foi diminuida
                If dQuantDiferenca < 0 Then
                    'Exclui a quant no item de conc
                    lErro = ItemConcorrencia_Exclui_QuantComprar(gcolItemConcorrencia(iIndice), iIndice, , , Abs(dQuantDiferenca))
                    If lErro <> SUCESSO Then gError 62761
                'Sen�o
                Else
                    'Inclui a quant no item de conc
                    lErro = ItemConcorrencia_Inclui_QuantComprar(gcolItemConcorrencia(iIndice), iIndice, , , dQuantDiferenca)
                    If lErro <> SUCESSO Then gError 62762
                End If
                
                'Atualiza as cota��e spara a nova quantidade
                lErro = ItemConcorrencia_Atualiza_Cotacoes(gcolItemConcorrencia, iIndice)
                If lErro <> SUCESSO Then gError 62763
                
                Exit For
            End If
        Next

        'Preenche o grid de Cota��es
        lErro = GridCotacoes_Preenche()
        If lErro <> SUCESSO Then gError 62764
    End If
    
    Saida_Celula_Quantidade2 = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade2:

    Saida_Celula_Quantidade2 = gErr

    Select Case gErr

        Case 63963, 63964, 62761, 62762, 62763, 62764, 63965
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 62744
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTCOMPRAR_NAO_PREENCHIDA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161224)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_EscolhidoReq(objGridInt As AdmGrid) As Long
'Faz a saida de c�lula de EscolhidoReq

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_EscolhidoReq

    Set objGridInt.objControle = EscolhidoReq

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 63962

    Exit Function

Erro_Saida_Celula_EscolhidoReq:

    Select Case gErr

        Case 63962
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161225)

    End Select

    Exit Function

End Function

Private Sub OrdenacaoReq_Click()

Dim lErro As Long
Dim colReqCompraSaida As New Collection
Dim colCampos As New Collection

On Error GoTo Erro_OrdenacaoReq_Click

    If gsOrdenacaoReq = "" Then Exit Sub

    'Verifica se OrdenacaoReq da tela � diferente de gsOrdenacao
    If OrdenacaoReq.Text <> gsOrdenacaoReq Then

        Call Monta_Colecao_Campos_Requisicao(colCampos, OrdenacaoReq.ListIndex)
        'Ordena
        lErro = Ordena_Colecao(gobjGeracaoPedCompraReq.colRequisicao, colReqCompraSaida, colCampos)
        If lErro <> SUCESSO Then gError 63908

        Set gobjGeracaoPedCompraReq.colRequisicao = colReqCompraSaida

    End If

    'COloca as Requsiicoes na tela ordenadamente
    lErro = GridRequisicoes_Preenche()
    If lErro <> SUCESSO Then gError 62750
    
    'Coloca os itens na tela de acordo com a ordem das requisi��es.
    lErro = GridItensReq_Preenche()
    If lErro <> SUCESSO Then gError 62751

    Exit Sub

Erro_OrdenacaoReq_Click:

    Select Case gErr

        Case 62750, 62751, 63907 To 63909

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161226)

    End Select

    Exit Sub

End Sub

Sub Monta_Colecao_Campos_Requisicao(colCampos As Collection, iOrdenacao As Integer)

    Select Case iOrdenacao

        Case 0

            colCampos.Add "iFilialEmpresa"
            colCampos.Add "lCodigo"

        Case 1

            colCampos.Add "dtDataLimite"
            colCampos.Add "iFilialEmpresa"
            colCampos.Add "lCodigo"

        Case 2

            colCampos.Add "dtData"
            colCampos.Add "iFilialEmpresa"
            colCampos.Add "lCodigo"

    End Select

End Sub

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

        objReqCompras.lCodigo = GridRequisicoes.TextMatrix(iIndice, iGrid_CodigoReq_Col)
        objReqCompras.dtDataLimite = StrParaDate(GridRequisicoes.TextMatrix(iIndice, iGrid_DataLimite_Col))
        objReqCompras.dtData = StrParaDate(GridRequisicoes.TextMatrix(iIndice, iGrid_DataReq_Col))
        objReqCompras.lUrgente = GridRequisicoes.TextMatrix(iIndice, iGrid_Urgente_Col)
        objReqCompras.lRequisitante = LCodigo_Extrai(GridRequisicoes.TextMatrix(iIndice, iGrid_Requisitante_Col))
        objReqCompras.sCcl = GridRequisicoes.TextMatrix(iIndice, iGrid_CclReq_Col)
        objReqCompras.sObservacao = GridRequisicoes.TextMatrix(iIndice, iGrid_ObservacaoReq_Col)
        objReqCompras.iFilialEmpresa = Codigo_Extrai(GridRequisicoes.TextMatrix(iIndice, iGrid_FilialReq_Col))

        'Adiciona em colRequisicao
        colRequisicao.Add objReqCompras

    Next

    GridRequisicoes_Recolhe = SUCESSO

    Exit Function

Erro_GridRequisicoes_Recolhe:

    GridRequisicoes_Recolhe = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161227)

    End Select

    Exit Function

End Function

Private Sub OrdenacaoCot_Click()

Dim lErro As Long

On Error GoTo Erro_Ordenacao_Click

    If gsOrdenacao = "" Then Exit Sub

    If gsOrdenacao <> OrdenacaoCot.Text Then
    
        gsOrdenacao = OrdenacaoCot.Text
        
        'Devolve os elementos ordenados para o  GridCotacoes
        lErro = GridCotacoes_Preenche()
        If lErro <> SUCESSO Then gError 63809

    End If

    Exit Sub

Erro_Ordenacao_Click:

    Select Case gErr

        Case 63807 To 63809
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161228)

    End Select

    Exit Sub

End Sub

Private Sub Quantidade2_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Quantidade2_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos2)

End Sub

Private Sub Quantidade2_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos2)

End Sub

Private Sub Quantidade2_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos2.objControle = Quantidade2
    lErro = Grid_Campo_Libera_Foco(objGridProdutos2)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub EscolhidoReq_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRequisicoes)

End Sub

Private Sub EscolhidoReq_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRequisicoes)

End Sub

Private Sub EscolhidoReq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCotacoes.objControle = EscolhidoReq
    lErro = Grid_Campo_Libera_Foco(objGridRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub EscolhidoItem_Click()

    iAlterado = REGISTRO_ALTERADO
    Call Atualiza_ItensReq
         
End Sub

Private Sub EscolhidoItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensRequisicoes)

End Sub

Private Sub EscolhidoItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensRequisicoes)

End Sub

Private Sub EscolhidoItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCotacoes.objControle = EscolhidoItem
    lErro = Grid_Campo_Libera_Foco(objGridItensRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantComprarItemRC_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantComprarItemRC_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensRequisicoes)

End Sub

Private Sub QuantComprarItemRC_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensRequisicoes)

End Sub

Private Sub QuantComprarItemRC_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCotacoes.objControle = QuantComprarItemRC
    lErro = Grid_Campo_Libera_Foco(objGridItensRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub EscolhidoProduto_Click()

Dim lErro As Long
Dim objItemConcorrencia As ClassItemConcorrencia

On Error GoTo Erro_EscolhidoProduto_Click

    iAlterado = REGISTRO_ALTERADO
        
    'Pega o item de concorr�ncia clicado
    
    Set objItemConcorrencia = gcolItemConcorrencia(GridProdutos1.Row)
    'Atualiza a escolha
    objItemConcorrencia.iEscolhido = GridProdutos1.TextMatrix(GridProdutos1.Row, iGrid_EscolhidoProduto_Col)

    'Repreenche o grid de produtos
    lErro = GridProdutos2_Preenche
    If lErro <> SUCESSO Then gError 62758
        
    Call Indica_Melhores
    Call GridCotacoes_Preenche
    
    Exit Sub
    
Erro_EscolhidoProduto_Click:

    Select Case gErr
    
        Case 62758
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161229)
            
    End Select
    
    Exit Sub
        
End Sub

Private Sub EscolhidoProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProdutos1)

End Sub

Private Sub EscolhidoProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos1)

End Sub

Private Sub EscolhidoProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCotacoes.objControle = EscolhidoProduto
    lErro = Grid_Campo_Libera_Foco(objGridProdutos1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantComprarCot_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantComprarCot_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCotacoes)

End Sub

Private Sub QuantComprarCot_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCotacoes)

End Sub

Private Sub QuantComprarCot_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCotacoes.objControle = QuantComprarCot
    lErro = Grid_Campo_Libera_Foco(objGridCotacoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub SelecionaDestino_Click()

Dim iIndice As Integer
Dim bCancel As Boolean

    'Verifica se SelecionaDestino estiver desmarcado
    If SelecionaDestino.Value = vbUnchecked Then

        'Desabilita todos os TipoDestino
        TipoDestino(TIPO_DESTINO_EMPRESA).Enabled = False
        TipoDestino(TIPO_DESTINO_FORNECEDOR).Enabled = False
        FornecedorLabel.Enabled = False
        FilEmprDestLabel.Enabled = False
        FilFornDestLabel.Enabled = False

        'Limpa os campos do Frame Destino()
        FilialEmpresa.Text = ""
        Fornecedor.Text = ""
        FilialFornec.ListIndex = -1

    'Verifica se SelecionaDestino est� marcado
    ElseIf SelecionaDestino.Value = vbChecked Then

        'Haabilita todos os TipoDestino
        TipoDestino(TIPO_DESTINO_EMPRESA).Enabled = True
        TipoDestino(TIPO_DESTINO_FORNECEDOR).Enabled = True
        FornecedorLabel.Enabled = True
        FilEmprDestLabel.Enabled = True
        FilFornDestLabel.Enabled = True

        Fornecedor.Enabled = True
        FilialFornec.Enabled = True
        FilialEmpresa.Enabled = True

        'Se nenhuma FilialEmpresa estiver selecionada
        If FilialEmpresa.ListIndex = -1 Then FilialEmpresa.Text = giFilialEmpresa
        Call FilialEmpresa_Validate(bCancel)

    End If

    iFrameSelecaoAlterado = REGISTRO_ALTERADO

    Exit Sub

End Sub

Private Sub TabProdutos_Click()

    'Se frame selecionado n�o for o atual esconde o frame atual, mostra o novo.
    If TabProdutos.SelectedItem.Index <> iFrameProdutoAtual Then

        If TabStrip_PodeTrocarTab(iFrameProdutoAtual, TabProdutos, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        FrameProdutos(TabProdutos.SelectedItem.Index).Visible = True
        'Torna Frame atual invisivel
        FrameProdutos(iFrameProdutoAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameProdutoAtual = TabProdutos.SelectedItem.Index

    End If

End Sub


Private Sub TipoDestino_Click(Index As Integer)

    'Se o TipoDestino for o mesmo j� selecionado, sai da rotina
    If Index = iFrameTipoDestinoAtual Then Exit Sub

    'Torna invisivel o FrameDestino com �ndice igual a iFrameDestinoAtual
    FrameTipoDestino(iFrameTipoDestinoAtual).Visible = False

    'Torna vis�vel o FrameDestino com �ndice igual a Index
    FrameTipoDestino(Index).Visible = True

    'Armazena novo valor de giFrameDestinoAtual
    iFrameTipoDestinoAtual = Index

    iFrameSelecaoAlterado = REGISTRO_ALTERADO

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
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 63888

        'Se nao encontra o �tem com o c�digo informado
        If lErro = 6730 Then

            'preeenche objFilialEmpresa
            objFilialEmpresa.iCodFilial = iCodigo

            'Le a FilialEmpresa
            lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
            If lErro <> SUCESSO And lErro <> 27378 Then gError 63889

            'Se nao encontrou => Erro
            If lErro = 27378 Then gError 63890

            If lErro = SUCESSO Then

                'Coloca na tela o codigo e o nome da FilialEmpresa
                FilialEmpresa.Text = objFilialEmpresa.lCodEmpresa & SEPARADOR & objFilialEmpresa.sNome

            End If

        End If

        'Se nao encontrou e nao era codigo
        If lErro = 6731 Then gError 63891

    Exit Sub

Erro_FilialEmpresa_Validate:

    Cancel = True

    Select Case gErr

        Case 63888, 63889
            'Erros tratados nas rotinas chamadas

        Case 63890
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, iCodigo)

        Case 63891
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA1", gErr, FilialEmpresa.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161230)

    End Select

    Exit Sub

End Sub

Private Sub TipoProduto_Click()

    iAlterado = REGISTRO_ALTERADO
    iFrameSelecaoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoProduto_ItemCheck(Item As Integer)
    iAlterado = REGISTRO_ALTERADO
    iFrameSelecaoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownDataAte_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownDataDe_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownDataLimAte_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownDataLimDe_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownDataLimDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataLimDe_DownClick

    'Diminui a Data
    lErro = Data_Up_Down_Click(DataLimiteDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 63880

    Exit Sub


Erro_UpDownDataLimDe_DownClick:

    Select Case gErr

        Case 63880
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161231)

    End Select

    Exit Sub

End Sub
Private Sub UpDownDataLimAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataLimAte_DownClick

    'Diminui a Data
    lErro = Data_Up_Down_Click(DataLimiteAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 63881

    Exit Sub


Erro_UpDownDataLimAte_DownClick:

    Select Case gErr

        Case 63881
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161232)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Diminui a Data
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 63882

    Exit Sub


Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 63882
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161233)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Diminui a Data
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 63883

    Exit Sub


Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 63883
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161234)

    End Select

    Exit Sub

End Sub


Private Sub UpDownDataLimDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataLimDe_UpClick

    'Aumenta a Data
    lErro = Data_Up_Down_Click(DataLimiteDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 63884

    Exit Sub

Erro_UpDownDataLimDe_UpClick:

    Select Case gErr

        Case 63884
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161235)

    End Select

    Exit Sub

End Sub


Private Sub UpDownDataLimAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataLimAte_UpClick

    'Aumenta a Data
    lErro = Data_Up_Down_Click(DataLimiteAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 63885

    Exit Sub

Erro_UpDownDataLimAte_UpClick:

    Select Case gErr

        Case 63885
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161236)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Aumenta a Data
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 63886

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 63886
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161237)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Aumenta a Data
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 63887

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 63887
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161238)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Grid_Requisicoes(objGridInt As AdmGrid) As Long
'Executa a Inicializa��o do grid Requisicoes

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Requisicoes

    'tela em quest�o
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Escolhido")
    objGridInt.colColuna.Add ("FilialEmpresa")
    objGridInt.colColuna.Add ("N�mero")
    objGridInt.colColuna.Add ("P.V.")
    objGridInt.colColuna.Add ("Data Limite")
    objGridInt.colColuna.Add ("Data RC")
    objGridInt.colColuna.Add ("Urgente")
    objGridInt.colColuna.Add ("Requisitante")
    objGridInt.colColuna.Add ("Centro C/L")
    objGridInt.colColuna.Add ("Observa��o")

    'campos de edi��o do grid
    objGridInt.colCampo.Add (EscolhidoReq.Name)
    objGridInt.colCampo.Add (FilialReq.Name)
    objGridInt.colCampo.Add (CodigoReq.Name)
    objGridInt.colCampo.Add (CodigoPV.Name)
    objGridInt.colCampo.Add (DataLimite.Name)
    objGridInt.colCampo.Add (DataReq.Name)
    objGridInt.colCampo.Add (Urgente.Name)
    objGridInt.colCampo.Add (Requisitante.Name)
    objGridInt.colCampo.Add (CclReq.Name)
    objGridInt.colCampo.Add (ObservacaoReq.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_EscolhidoReq_Col = 1
    iGrid_FilialReq_Col = 2
    iGrid_CodigoReq_Col = 3
    iGrid_CodigoPV_Col = 4
    iGrid_DataLimite_Col = 5
    iGrid_DataReq_Col = 6
    iGrid_Urgente_Col = 7
    iGrid_Requisitante_Col = 8
    iGrid_CclReq_Col = 9
    iGrid_ObservacaoReq_Col = 10

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridRequisicoes

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_REQUISICOES + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 20

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
'    GridCotacoes.Width = 8295

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Chama rotina de Inicializa��o do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Requisicoes = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Requisicoes:

    Inicializa_Grid_Requisicoes = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161239)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Cotacoes(objGridInt As AdmGrid) As Long
'Executa a Inicializa��o do grid Cotacoes

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Cotacoes

    'tela em quest�o
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Escolhido")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descri��o")
    objGridInt.colColuna.Add ("Prefer�ncia")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial Fornecedor")
    objGridInt.colColuna.Add ("Moeda")
    objGridInt.colColuna.Add ("Pre�o Unit�rio")
    objGridInt.colColuna.Add ("Taxa Forn.")
    objGridInt.colColuna.Add ("Cota��o")
    objGridInt.colColuna.Add ("Pre�o Unit�rio (R$)")
    objGridInt.colColuna.Add ("Cond. Pagto")
    objGridInt.colColuna.Add ("Quant. Cotada")
    objGridInt.colColuna.Add ("A Comprar")
    objGridInt.colColuna.Add ("UM")
    objGridInt.colColuna.Add ("Valor Presente (R$)")
    objGridInt.colColuna.Add ("Valor Item (R$)")
    objGridInt.colColuna.Add ("Tipo Tributacao")
    objGridInt.colColuna.Add ("Al�quota IPI")
    objGridInt.colColuna.Add ("Al�quota ICMS")
    objGridInt.colColuna.Add ("Ped. Cota��o")
    objGridInt.colColuna.Add ("Data Cota��o")
    objGridInt.colColuna.Add ("Data Validade")
    objGridInt.colColuna.Add ("Prazo Entrega")
    objGridInt.colColuna.Add ("Data Entrega")
    objGridInt.colColuna.Add ("Data Necessidade")
    objGridInt.colColuna.Add ("Para Entrega")
    objGridInt.colColuna.Add ("Motivo da Escolha")

    'campos de edi��o do grid
    objGridInt.colCampo.Add (EscolhidoCot.Name)
    objGridInt.colCampo.Add (ProdutoCot.Name)
    objGridInt.colCampo.Add (DescProdutoCot.Name)
    objGridInt.colCampo.Add (Preferencia.Name)
    objGridInt.colCampo.Add (FornecedorCot.Name)
    objGridInt.colCampo.Add (FilialFornCot.Name)
    objGridInt.colCampo.Add (Moeda.Name)
    objGridInt.colCampo.Add (PrecoUnitario.Name)
    objGridInt.colCampo.Add (TaxaForn.Name)
    objGridInt.colCampo.Add (Cotacao.Name)
    objGridInt.colCampo.Add (PrecoUnitarioReal.Name)
    objGridInt.colCampo.Add (CondPagto.Name)
    objGridInt.colCampo.Add (QuantidadeCot.Name)
    objGridInt.colCampo.Add (QuantComprarCot.Name)
    objGridInt.colCampo.Add (UnidadeMedCot.Name)
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
    iGrid_Moeda_Col = 7
    iGrid_PrecoUnitarioCot_Col = 8
    iGrid_TaxaForn_Col = 9
    iGrid_CotacaoMoeda_Col = 10
    iGrid_PrecoUnitario_RS_Col = 11
    iGrid_CondPagtoCot_Col = 12
    iGrid_QuantidadeCot_Col = 13
    iGrid_QuantComprarCot_Col = 14
    iGrid_UMCot_Col = 15
    iGrid_ValorPresenteCot_Col = 16
    iGrid_ValorItem_Col = 17
    iGrid_TipoTributacaoCot_Col = 18
    iGrid_AliquotaIPI_Col = 19
    iGrid_AliquotaICMS_Col = 20
    iGrid_PedidoCot_Col = 21
    iGrid_DataCotacaoCot_Col = 22
    iGrid_DataValidadeCot_Col = 23
    iGrid_PrazoEntrega_Col = 24
    iGrid_DataEntrega_Col = 25
    iGrid_DataNecessidade_Col = 26
    iGrid_QuantidadeEntrega_Col = 27
    iGrid_MotivoEscolhaCot_Col = 28

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridCotacoes

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_COTACOES + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 15

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
'    GridCotacoes.Width = 8295
    GridCotacoes.ColWidth(0) = 400
    
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama rotina de Inicializa��o do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Cotacoes = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Cotacoes:

    Inicializa_Grid_Cotacoes = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161240)

    End Select

    Exit Function

End Function

Private Function Carrega_MotivoEscolha() As Long
'Carrega a combobox FilialEmpresa

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_MotivoEscolha

    'L� o C�digo e o Nome de todo MotivoEscolha do BD
    lErro = CF("Cod_Nomes_Le", "Motivo", "Codigo", "Motivo", STRING_NOME_TABELA, colCodigoNome)
    If lErro <> SUCESSO Then gError 63875

    'Carrega a combo de Motivo Escolha com c�digo e nome
    For Each objCodigoNome In colCodigoNome

        'Verifica se o MotivoEscolha � diferente de Exclusividade
        If objCodigoNome.iCodigo <> MOTIVO_EXCLUSIVO Then

            MotivoEscolhaCot.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
            MotivoEscolhaCot.ItemData(MotivoEscolhaCot.NewIndex) = objCodigoNome.iCodigo

        End If

    Next

    Carrega_MotivoEscolha = SUCESSO

    Exit Function

Erro_Carrega_MotivoEscolha:

    Carrega_MotivoEscolha = gErr

    Select Case gErr

        Case 63875
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161241)

    End Select

    Exit Function

End Function

Private Sub BotaoReqCompras_Click()
'Chama a tela ReqComprasEnv

Dim objRequisicaoCompras As New ClassRequisicaoCompras
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoReqCompras_Click

    'Verifica se existe alguma linha selecionada no GridRequisicoes
    If GridRequisicoes.Row = 0 Then gError 89445

    objRequisicaoCompras.lCodigo = StrParaLong(GridRequisicoes.TextMatrix(GridRequisicoes.Row, iGrid_CodigoReq_Col))
    objRequisicaoCompras.iFilialEmpresa = Codigo_Extrai(GridRequisicoes.TextMatrix(GridRequisicoes.Row, iGrid_FilialReq_Col))

    'Chama a tela ReqComprasEnv
    Call Chama_Tela("ReqComprasCons", objRequisicaoCompras)

    Exit Sub

Erro_BotaoReqCompras_Click:

    Select Case gErr
    
        Case 89445
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161242)
            
    End Select
    
    Exit Sub

End Sub

Private Sub DataLimiteDe_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataLimiteDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataLimiteDe_Validate

    'Verifica se  DataLimiteDe foi preenchida
    If Len(Trim(DataLimiteDe.Text)) = 0 Then Exit Sub

    'Critica DataLimiteDe
    lErro = Data_Critica(DataLimiteDe.Text)
    If lErro <> SUCESSO Then gError 63876

    Exit Sub

Erro_DataLimiteDe_Validate:

    Cancel = True

    Select Case gErr

        Case 63876
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161243)

    End Select

    Exit Sub

End Sub

Private Sub DataLimiteAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataLimiteAte_Validate

    'Verifica se  DataLimiteAte foi preenchida
    If Len(Trim(DataLimiteAte.Text)) = 0 Then Exit Sub

    'Critica DataLimiteAte
    lErro = Data_Critica(DataLimiteAte.Text)
    If lErro <> SUCESSO Then gError 63877

    Exit Sub

Erro_DataLimiteAte_Validate:

    Cancel = True

    Select Case gErr

        Case 63877
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161244)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'Verifica se  DataDe foi preenchida
    If Len(Trim(DataDe.Text)) = 0 Then Exit Sub

    'Critica DataDe
    lErro = Data_Critica(DataDe.Text)
    If lErro <> SUCESSO Then gError 63878

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 63878
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161245)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'Verifica se  DataAte foi preenchida
    If Len(Trim(DataAte.Text)) = 0 Then Exit Sub

    'Critica DataAte
    lErro = Data_Critica(DataAte.Text)
    If lErro <> SUCESSO Then gError 63879

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 63879
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161246)

    End Select

    Exit Sub

End Sub

Private Sub FornecedorLabel_Click()
'Chama a tela FornecedorLista

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As New Collection

    'Coloca o Fornecedor que est� na tela no objFornecedor
    objFornecedor.sNomeReduzido = Fornecedor.Text

    'Chama a tela FornecedorLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)


End Sub

Sub Calcula_Preferencia(objCotItemConc As ClassCotacaoItemConc, sProduto As String, dQuantComprar As Double)
'Calcula a Prefer�ncia

Dim iIndice As Integer
Dim dQuantPreferencial As Double
Dim dQuantComprarItem As Double
    
    dQuantPreferencial = 0
    
    If dQuantComprar = 0 Then Exit Sub
    
    For iIndice = 1 To objGridItensRequisicoes.iLinhasExistentes
    
        If StrParaInt(GridItensRequisicoes.TextMatrix(iIndice, iGrid_EscolhidoItem_Col)) = MARCADO Then
        
            If GridItensRequisicoes.TextMatrix(iIndice, iGrid_ProdutoItemRC_Col) = sProduto And _
              GridItensRequisicoes.TextMatrix(iIndice, iGrid_FilialFornItemRC_Col) = objCotItemConc.sFilial And _
              GridItensRequisicoes.TextMatrix(iIndice, iGrid_FornecedorItemRC_Col) = objCotItemConc.sFornecedor And _
              GridItensRequisicoes.TextMatrix(iIndice, iGrid_ExclusivoItemRC_Col) = "Preferencial" Then
                
                Call Busca_QuantComprar_ItemReq(StrParaLong(GridItensRequisicoes.TextMatrix(iIndice, iGrid_CodigoReqItem_Col)), Codigo_Extrai(GridItensRequisicoes.TextMatrix(iIndice, iGrid_FilialReqItem_Col)), StrParaInt(GridItensRequisicoes.TextMatrix(iIndice, iGrid_Item_Col)), dQuantComprarItem)
              
                dQuantPreferencial = dQuantPreferencial + dQuantComprarItem
            End If
        End If
    Next
            
    objCotItemConc.dPreferencia = dQuantPreferencial / dQuantComprar
    Exit Sub

End Sub

Function Move_TabSelecao_Memoria(gobjGeracaoPedCompraReq As ClassGeracaoPedCompraReq) As Long
'Recolhe dados do TAB de Sele��o

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iIndice As Integer

On Error GoTo Erro_Move_TabSelecao_Memoria

    'DataDe
    If Len(Trim(DataDe.ClipText)) > 0 Then
        gobjGeracaoPedCompraReq.dtDataDe = DataDe.Text
    Else
        gobjGeracaoPedCompraReq.dtDataDe = DATA_NULA
    End If

    'DataAte
    If Len(Trim(DataAte.ClipText)) > 0 Then
        gobjGeracaoPedCompraReq.dtDataAte = DataAte.Text
    Else
        gobjGeracaoPedCompraReq.dtDataAte = DATA_NULA
    End If

    'DataLimiteDe
    If Len(Trim(DataLimiteDe.ClipText)) > 0 Then
        gobjGeracaoPedCompraReq.dtDataLimiteDe = DataLimiteDe.Text
    Else
        gobjGeracaoPedCompraReq.dtDataLimiteDe = DATA_NULA
    End If

    'DataLimiteAte
    If Len(Trim(DataLimiteAte.ClipText)) > 0 Then
        gobjGeracaoPedCompraReq.dtDataLimiteAte = DataLimiteAte.Text
    Else
        gobjGeracaoPedCompraReq.dtDataLimiteAte = DATA_NULA
    End If

    'Local de Entrega
    gobjGeracaoPedCompraReq.iSelecionaDestino = SelecionaDestino.Value

    If SelecionaDestino.Value = vbChecked Then

        'Tipo de Destino
        If TipoDestino.Item(TIPO_DESTINO_EMPRESA).Value = True Then
            gobjGeracaoPedCompraReq.iTipoDestino = TIPO_DESTINO_EMPRESA

            'Filial Empresa Destino
            gobjGeracaoPedCompraReq.iFilialDestino = Codigo_Extrai(FilialEmpresa.Text)

        Else

            gobjGeracaoPedCompraReq.iTipoDestino = TIPO_DESTINO_FORNECEDOR

            'Se o Fornecedor foi preenchida
            If Len(Trim(Fornecedor.Text)) > 0 Then

                'Fornecedor e Filial Destino
                objFornecedor.sNomeReduzido = Fornecedor.Text

                'L� o c�digo do Fornecedor
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO And lErro <> 6681 Then gError 63954
                If lErro = 6681 Then gError 63955

                gobjGeracaoPedCompraReq.lFornCliDestino = objFornecedor.lCodigo
                gobjGeracaoPedCompraReq.iFilialDestino = Codigo_Extrai(FilialFornec.Text)

            End If

        End If

    End If

    'C�digo De
    gobjGeracaoPedCompraReq.lCodigoDe = StrParaLong(CodigoDe.Text)

    'C�digo At�
    gobjGeracaoPedCompraReq.lCodigoAte = StrParaLong(CodigoAte.Text)

    Set gobjGeracaoPedCompraReq.colTipoProduto = New Collection
    Set gobjGeracaoPedCompraReq.colTipoCategoria = New Collection

    'Armazena em colTipoProduto os Tipos de Produtos selecionados no TabSelecao
    For iIndice = 0 To TipoProduto.ListCount - 1
        If TipoProduto.Selected(iIndice) = True Then
            gobjGeracaoPedCompraReq.colTipoProduto.Add (Codigo_Extrai(TipoProduto.List(iIndice)))
        End If
    Next
    
    'Verifica se algum tipo de produto foi selecionado
    If gobjGeracaoPedCompraReq.colTipoProduto.Count = 0 Then gError 74868
    
    'Armazena em colTipoProduto os Tipos de Produtos selecionados no TabSelecao
    For iIndice = 0 To ItensCategoria.ListCount - 1
        If ItensCategoria.Selected(iIndice) = True Then
            gobjGeracaoPedCompraReq.colTipoCategoria.Add (ItensCategoria.List(iIndice))
        End If
    Next

    gobjGeracaoPedCompraReq.dtDataEnvio = DATA_NULA
    
    gobjGeracaoPedCompraReq.sCategoria = Categoria.Text

    Move_TabSelecao_Memoria = SUCESSO

    Exit Function

Erro_Move_TabSelecao_Memoria:

    Move_TabSelecao_Memoria = gErr

    Select Case gErr

        Case 63954

        Case 63955
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)

        Case 74868
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_TIPOPRODUTO_SELECIONADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161247)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Gera��o de Pedidos de Compra por Requisi��es"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "GeracaoPedCompraReq"

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

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub
Private Sub FornecedorLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecedorLabel, Source, X, Y)
End Sub

Private Sub FornecedorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel, Button, Shift, X, Y)
End Sub

Private Sub Label37_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label37, Source, X, Y)
End Sub

Private Sub Label37_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label37, Button, Shift, X, Y)
End Sub

Private Sub Label32_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label32, Source, X, Y)
End Sub

Private Sub Label32_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label32, Button, Shift, X, Y)
End Sub

Private Sub Comprador_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Comprador, Source, X, Y)
End Sub

Private Sub Comprador_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Comprador, Button, Shift, X, Y)
End Sub


Private Sub Concorrencia_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Concorrencia, Source, X, Y)
End Sub

Private Sub Concorrencia_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Concorrencia, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub TaxaEmpresa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TaxaEmpresa, Source, X, Y)
End Sub

Private Sub TaxaEmpresa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TaxaEmpresa, Button, Shift, X, Y)
End Sub



Private Sub Label17_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label17, Source, X, Y)
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label17, Button, Shift, X, Y)
End Sub

Private Sub Label45_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label45, Source, X, Y)
End Sub

Private Sub Label45_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label45, Button, Shift, X, Y)
End Sub

Private Sub Label31_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label31, Source, X, Y)
End Sub

Private Sub Label31_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label31, Button, Shift, X, Y)
End Sub

Private Sub Label57_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label57, Source, X, Y)
End Sub

Private Sub Label57_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label57, Button, Shift, X, Y)
End Sub

Private Sub BotaoDesmarcarTodosItensRC_Click()
'Desmarca todas CheckBox do GridItensRequisicoes

Dim iIndice As Integer
Dim objReqCompras As ClassRequisicaoCompras
Dim objItemRC As ClassItemReqCompras

    'Desmarca na cole��o todos os itens
    For Each objReqCompras In gobjGeracaoPedCompraReq.colRequisicao
        For Each objItemRC In objReqCompras.colItens
            objItemRC.iSelecionado = DESMARCADO
        Next
    Next
    
    'Desmarca no grid todos os itens
    For iIndice = 1 To objGridItensRequisicoes.iLinhasExistentes
        GridItensRequisicoes.TextMatrix(iIndice, iGrid_EscolhidoItem_Col) = DESMARCADO
    Next

    'Limpa a cole��o de itens de concorr�ncia
    Set gcolItemConcorrencia = New Collection
    
    Call Grid_Refresh_Checkbox(objGridItensRequisicoes)
    
    Call Grid_Limpa(objGridProdutos1)
    Call Grid_Limpa(objGridProdutos2)
    Call Grid_Limpa(objGridCotacoes)
    Call Calcula_TotalItens
    
    Exit Sub

End Sub

Private Sub BotaoDesmarcarTodosProd_Click()
'Desmarca todas CheckBox do GridProdutos1
Dim iIndice As Integer

    'Marca todos os Itens do GridProdutos1
    For iIndice = 1 To objGridProdutos1.iLinhasExistentes
        GridProdutos1.TextMatrix(iIndice, iGrid_EscolhidoProduto_Col) = DESMARCADO
        gcolItemConcorrencia(iIndice).iEscolhido = DESMARCADO
    Next

    Call Grid_Refresh_Checkbox(objGridProdutos1)

    Call Grid_Limpa(objGridProdutos2)
    Call Grid_Limpa(objGridCotacoes)
    Call Calcula_TotalItens
    
    Exit Sub

End Sub

Private Sub BotaoDesmarcarTodosReq_Click()
'Desmarca todas CheckBox do GridRequisicoes

Dim iLinha As Integer

    Set gcolItemConcorrencia = New Collection
    
    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridRequisicoes.iLinhasExistentes
    
        'Desmarca na tela a linha em quest�o
        GridRequisicoes.TextMatrix(iLinha, iGrid_EscolhidoReq_Col) = GRID_CHECKBOX_INATIVO
        gobjGeracaoPedCompraReq.colRequisicao(iLinha).iSelecionado = DESMARCADO

    Next

    'Atualiza na tela a checkbox desmarcada
    Call Grid_Refresh_Checkbox(objGridRequisicoes)
    
    Call Grid_Limpa(objGridItensRequisicoes)
    Call Grid_Limpa(objGridProdutos1)
    Call Grid_Limpa(objGridProdutos2)
    Call Grid_Limpa(objGridCotacoes)
    Call Calcula_TotalItens

    Exit Sub

End Sub

Private Sub BotaoEditarProduto_Click(Index As Integer)
'Chama a tela de Produtos

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer

On Error GoTo Erro_BotaoEditarProduto_Click

    'Se est� editando um produto do GridProdutos1
    If FrameProdutos(1).Visible = True Then

        'Verifica se tem alguma linha selecionada no GridProdutos1
        If GridProdutos1.Row = 0 Then gError 63900

        'Verifica se o Produto est� preenchido
        If Len(Trim(GridProdutos1.TextMatrix(GridProdutos1.Row, iGrid_Produto1_Col))) > 0 Then
            lErro = CF("Produto_Formata", GridProdutos1.TextMatrix(GridProdutos1.Row, iGrid_Produto1_Col), sProduto, iPreenchido)
            If lErro <> SUCESSO Then gError 63901
            If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""
        End If

    'Se est� editando um produto do GridProdutos2
    Else

        'Verifica se tem alguma linha selecionada no GridProdutos1
        If GridProdutos2.Row = 0 Then gError 63902

        'Verifica se o Produto est� preenchido
        If Len(Trim(GridProdutos2.TextMatrix(GridProdutos2.Row, iGrid_Produto2_Col))) > 0 Then
            lErro = CF("Produto_Formata", GridProdutos2.TextMatrix(GridProdutos2.Row, iGrid_Produto2_Col), sProduto, iPreenchido)
            If lErro <> SUCESSO Then gError 63903
            If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""
        End If

    End If

    objProduto.sCodigo = sProduto

    'Chama a Tela Produto
    Call Chama_Tela("Produto", objProduto)

    Exit Sub

Erro_BotaoEditarProduto_Click:

    Select Case gErr

        Case 63900, 63902
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 63901, 63903
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161248)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGeraPedidos_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGeraPedidos_Click

    'Gera Pedido de Compra
    lErro = Gravar_Pedidos()
    If lErro <> SUCESSO Then gError 63913

    iAlterado = 0

    Exit Sub

Erro_BotaoGeraPedidos_Click:

    Select Case gErr

        Case 63913
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161249)

    End Select

    Exit Sub

End Sub

Function Gravar_Pedidos() As Long

Dim lErro As Long
Dim dQuantCotacao As Double
Dim objConcorrencia As New ClassConcorrencia
Dim colPedidoCompra As New Collection
Dim iIndice As Integer
Dim sProduto As String
Dim sFornecedor As String
Dim iFilial As Integer
Dim dQuantRC As Double
Dim iLinha As Integer
Dim iProdutoNaColecao As Integer
Dim objItemConc As New ClassItemConcorrencia
Dim colQuantidades As New Collection

On Error GoTo Erro_Gravar_Pedidos

    GL_objMDIForm.MousePointer = vbHourglass

    'Recolhe os dados da tela
    lErro = Move_Concorrencia_Memoria(objConcorrencia)
    If lErro <> SUCESSO Then gError 63920

    'Atualiza a Concorrencia no Banco de Dados
    lErro = CF("Concorrencia_Grava", objConcorrencia)
    If lErro <> SUCESSO Then gError 63921

    'Carrega em colPedidoCompras os Pedidos de Compra gerados a partir de diferentes Fornecedores e FiliaisFornecedores
    lErro = Carrega_Dados_Pedidos(objConcorrencia, colPedidoCompra)
    If lErro <> SUCESSO Then gError 63922

    'Grava o Pedido de Compras
    lErro = CF("PedCompra_Concorrencia_Grava", objConcorrencia, colPedidoCompra)
    If lErro <> SUCESSO Then gError 63923

    '#####################################
    'Inserido por Wagner
    If colPedidoCompra.Count > 0 Then
        Call Rotina_Aviso(vbOKOnly, "AVISO_INFORMA_CODIGO_PEDCOMPRA_GRAVADO", colPedidoCompra.Item(1).lCodigo, colPedidoCompra.Item(colPedidoCompra.Count).lCodigo)
    End If
    '#####################################

    'Limpa a tela
    Call Limpa_Tela_GeracaoPedCompraReq

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Pedidos = SUCESSO

    Exit Function

Erro_Gravar_Pedidos:

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Pedidos = gErr

    Select Case gErr

        Case 63919 To 63923

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161250)

    End Select

    Exit Function

End Function

Private Function Move_Concorrencia_Memoria(objConcorrencia As ClassConcorrencia) As Long
'Recolhe os dados da tela e armazena em objConcorrencia

Dim lErro As Long
Dim objUsuario As New ClassUsuario
Dim objComprador As New ClassComprador
Dim objFornecedor As New ClassFornecedor
Dim iLinha As Integer

On Error GoTo Erro_Move_Concorrencia_Memoria
    
    'Verifica se o GridRequisicoes est� vazio
    If objGridRequisicoes.iLinhasExistentes = 0 Then gError 63924
    
    'Verifica se existe algum Item de Requisicao selecionado
    For iLinha = 1 To objGridItensRequisicoes.iLinhasExistentes
        If GridItensRequisicoes.TextMatrix(iLinha, iGrid_EscolhidoItem_Col) = GRID_CHECKBOX_ATIVO Then
            Exit For
        End If
    Next

    If iLinha > objGridItensRequisicoes.iLinhasExistentes Then gError 63925
    
    'Verifica se existe algum Item de Requisicao selecionado
    For iLinha = 1 To objGridProdutos1.iLinhasExistentes
        If GridProdutos1.TextMatrix(iLinha, iGrid_EscolhidoProduto_Col) = GRID_CHECKBOX_ATIVO Then
            Exit For
        End If
    Next

    If iLinha > objGridProdutos1.iLinhasExistentes Then gError 63749
    
    If SelecionaDestino.Value = vbChecked Then
        'Verifica o Tipo de Destino selecionado � FilialEmpresa
        If TipoDestino(TIPO_DESTINO_EMPRESA).Value = True Then
    
            'Verifica se a FilialEmpresa est� preenchida
            If Len(Trim(FilialEmpresa.Text)) = 0 Then gError 63746
            
            objConcorrencia.iTipoDestino = TIPO_DESTINO_EMPRESA
            objConcorrencia.iFilialDestino = Codigo_Extrai(FilialEmpresa.Text)
    
        'Verifica se o TipoDestino � Fornecedor
        ElseIf TipoDestino(TIPO_DESTINO_FORNECEDOR).Value = True Then
    
            'Verifica se o Fornecedor est� preenchido
            If Len(Trim(Fornecedor.Text)) = 0 Then gError 63747
    
            'Verifica se a Filial do Fornecedor est� preenchida
            If Len(Trim(FilialFornec.Text)) = 0 Then gError 63748
    
            objConcorrencia.iTipoDestino = TIPO_DESTINO_FORNECEDOR
            objConcorrencia.iFilialDestino = Codigo_Extrai(FilialFornec.Text)
            
            'L� o Fornecedor
            objFornecedor.sNomeReduzido = Fornecedor.Text
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then gError 63775
                
            'Se o Fornecedor n�o estiver cadastrado, Erro
            If lErro = 6681 Then gError 70491
            objConcorrencia.lFornCliDestino = objFornecedor.lCodigo
        End If
    Else
        objConcorrencia.iTipoDestino = TIPO_DESTINO_AUSENTE
    End If
    
    'Verifica se o GridProdutos est� vazio
    If objGridProdutos1.iLinhasExistentes = 0 Then gError 63749
    
    objConcorrencia.dTaxaFinanceira = PercentParaDbl(TaxaEmpresa.Caption)
    
    'verifica se o c�digo da concorrencia est� preenchido
    If Len(Trim(Concorrencia.Caption)) = 0 Then gError 76083
    
    objConcorrencia.lCodigo = StrParaLong(Concorrencia.Caption)

    objUsuario.sNomeReduzido = Comprador.Caption

    'L� o usuario a partir do nome reduzido
    lErro = CF("Usuario_Le_NomeRed", objUsuario)
    If lErro <> SUCESSO And lErro <> 57269 Then gError 63774
    If lErro = 57269 Then gError 63777

    objComprador.sCodUsuario = objUsuario.sCodUsuario

    'L� o comprador a partir do codUsuario
    lErro = CF("Comprador_Le_Usuario", objComprador)
    If lErro <> SUCESSO And lErro <> 50059 Then gError 63820

    'Se n�o encontrou o comprador==>erro
    If lErro = 50059 Then gError 70490

    objConcorrencia.iComprador = objComprador.iCodigo
    objConcorrencia.iFilialEmpresa = giFilialEmpresa
    objConcorrencia.dtData = gdtDataAtual
    objConcorrencia.sDescricao = Descricao.Text

    'Move os itens da concorr�ncia para a mem�ria
    lErro = Move_ItensConcorrencia_Memoria(objConcorrencia)
    If lErro <> SUCESSO Then gError 63776

    Move_Concorrencia_Memoria = SUCESSO

    Exit Function

Erro_Move_Concorrencia_Memoria:

    Move_Concorrencia_Memoria = gErr

    Select Case gErr

        Case 63924
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAO_NAO_SELECIONADA", gErr)

        Case 63925
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_REQUISICAO_NAO_SELECIONADO", gErr)

        Case 63746
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_DESTINO_NAO_PREENCHIDA", gErr)

        Case 63747
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_DESTINO_NAO_PREENCHIDO", gErr)
        
        Case 63748
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORN_DESTINO_NAO_PREENCHIDA", gErr)
        
        Case 63749
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_ITEMCONC_SELECIONADO", gErr)

        Case 63777
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_INEXISTENTE", gErr, objUsuario.sNomeReduzido)
        
        Case 63820, 63774, 63775, 63776
            'Erros tratados nas rotinas chamadas

        Case 70490
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_COMPRADOR", gErr, objComprador.sCodUsuario)

        Case 70491
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
        
        Case 76083
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CONCORRENCIA_NAO_PREENCHIDO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161251)

    End Select

    Exit Function

End Function

Function Move_ItensConcorrencia_Memoria(objConcorrencia As ClassConcorrencia) As Long
'Move os dados dos Itens da Concorr�ncia (GridProdutos1) para a mem�ria

Dim lErro As Long
Dim iItem As Integer
Dim objItemConcorrencia As ClassItemConcorrencia

On Error GoTo Erro_Move_ItensConcorrencia_Memoria
            
    iItem = 0
    'Para cada item de concorrencia
    For Each objItemConcorrencia In gcolItemConcorrencia
        
        iItem = iItem + 1
        
        If objItemConcorrencia.iEscolhido = MARCADO Then
            'verifica se a quantidade foi preenchida
            If objItemConcorrencia.dQuantidade = 0 Then gError 63750
            
            'valida a quantidade do item de concorr�ncia
            lErro = Valida_Quantidade(objItemConcorrencia, iItem)
            If lErro <> SUCESSO Then gError 70492
            
            objConcorrencia.colItens.Add objItemConcorrencia
        End If
    Next
    
    Move_ItensConcorrencia_Memoria = SUCESSO

    Exit Function

Erro_Move_ItensConcorrencia_Memoria:

    Move_ItensConcorrencia_Memoria = gErr

    Select Case gErr

        Case 63750
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTCOMPRAR_NAO_PREENCHIDA", gErr)

        Case 70492

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161252)

    End Select

    Exit Function

End Function


Private Sub BotaoGravaConcorrencia_Click()
'Grava a Concorrencia

Dim lErro As Long

On Error GoTo Erro_BotaoGravaConcorrencia_Click
    
    'Insere ou Altera uma concorrencia no BD
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 63672
    
    Exit Sub

Erro_BotaoGravaConcorrencia_Click:
   
    Select Case gErr
        
        Case 63672
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161253)

    End Select

    Exit Sub
    
End Sub

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
Dim sItem As String
Dim lCodigoPV As Long

On Error GoTo Erro_GridRequisicoes_Preenche

    'Limpa o Grid de Requisi��es
    Call Grid_Limpa(objGridRequisicoes)

    '#######################################################################
    'Inserido por Wagner 25/05/2006
    If gobjGeracaoPedCompraReq.colRequisicao.Count >= objGridRequisicoes.objGrid.Rows Then
        Call Refaz_Grid(objGridRequisicoes, gobjGeracaoPedCompraReq.colRequisicao.Count)
    End If
    '#######################################################################
    
    If gobjGeracaoPedCompraReq.colRequisicao.Count > 0 Then

        'Preenche o GridRequisicoes
        For Each objRequisicao In gobjGeracaoPedCompraReq.colRequisicao

            iLinha = objGridRequisicoes.iLinhasExistentes + 1
            
            Call Busca_Na_Colecao(colFiliais, objRequisicao.iFilialEmpresa, iPosicao)
    
            If iPosicao = 0 Then
    
                objFilialEmpresa.iCodFilial = objRequisicao.iFilialEmpresa
    
                'L� a FilialEmpresa
                lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
                If lErro <> SUCESSO And lErro <> 27378 Then gError 63976
    
                'Se n�o encontrou ==>Erro
                If lErro = 27378 Then gError 63977
    
                Set objlCodigoNome = New AdmlCodigoNome
    
                objlCodigoNome.lCodigo = objFilialEmpresa.iCodFilial
                objlCodigoNome.sNome = objFilialEmpresa.sNome
    
                colFiliais.Add objlCodigoNome.lCodigo, objlCodigoNome.sNome
    
            Else
    
                Set objlCodigoNome = colFiliais(iPosicao)
    
            End If
    
            'Preenche a Filial de Requisicao com c�digo e nome reduzido
            GridRequisicoes.TextMatrix(iLinha, iGrid_FilialReq_Col) = objlCodigoNome.lCodigo & SEPARADOR & objlCodigoNome.sNome
            GridRequisicoes.TextMatrix(iLinha, iGrid_CodigoReq_Col) = objRequisicao.lCodigo
    
            'Verifica se DataLimite � diferente de Data Nula
            If objRequisicao.dtDataLimite <> DATA_NULA Then GridRequisicoes.TextMatrix(iLinha, iGrid_DataLimite_Col) = Format(objRequisicao.dtDataLimite, "dd/mm/yyyy")
    
            'Verifica se Data � diferente de Data Nula
            If objRequisicao.dtData <> DATA_NULA Then GridRequisicoes.TextMatrix(iLinha, iGrid_DataReq_Col) = Format(objRequisicao.dtData, "dd/mm/yyyy")
    
            GridRequisicoes.TextMatrix(iLinha, iGrid_Urgente_Col) = objRequisicao.lUrgente
    
            Call Busca_Na_Colecao(colRequisitantes, objRequisicao.lRequisitante, iPosicao)
            
            If iPosicao = 0 Then
                objRequisitante.lCodigo = objRequisicao.lRequisitante
        
                'L� o requisitante
                lErro = CF("Requisitante_Le", objRequisitante)
                If lErro <> SUCESSO And lErro <> 49084 Then gError 63978
        
                'Se n�o encontrou o Requisitante ==> Erro
                If lErro = 49084 Then gError 63979
                
                Set objlCodigoNome = New AdmlCodigoNome
                
                objlCodigoNome.lCodigo = objRequisitante.lCodigo
                objlCodigoNome.sNome = objRequisitante.sNomeReduzido
                
                colRequisitantes.Add objlCodigoNome.lCodigo, objlCodigoNome.sNome
                
            Else
                Set objlCodigoNome = colRequisitantes(iPosicao)
            End If
            
            'Preenche o Requisitante com o c�digo e o nome reduzido
            GridRequisicoes.TextMatrix(iLinha, iGrid_Requisitante_Col) = objlCodigoNome.lCodigo & SEPARADOR & objlCodigoNome.sNome
    
            'Se o Ccl est� preenchida
            If Len(Trim(objRequisicao.sCcl)) > 0 Then
    
                'Mascara o Produto
                lErro = Mascara_MascararCcl(objRequisicao.sCcl, sCclMascarado)
                If lErro <> SUCESSO Then gError 63980
    
                'Preenche o Ccl
                GridRequisicoes.TextMatrix(iLinha, iGrid_CclReq_Col) = sCclMascarado
    
            End If
    
            'Preenche a Observacao
            GridRequisicoes.TextMatrix(iLinha, iGrid_ObservacaoReq_Col) = objRequisicao.sObservacao
    
            'Selecionado
            GridRequisicoes.TextMatrix(iLinha, iGrid_EscolhidoReq_Col) = objRequisicao.iSelecionado
                       
            If Len(Trim(objRequisicao.sOPCodigo)) > 0 Then
    
                lErro = Preenche_CodigoPV(objRequisicao, lCodigoPV)
                If lErro <> SUCESSO Then gError 178875
        
                If lCodigoPV <> 0 Then
                    GridRequisicoes.TextMatrix(iLinha, iGrid_CodigoPV_Col) = lCodigoPV
                End If
    
            End If
                       
                       
                       
            objGridRequisicoes.iLinhasExistentes = iLinha
            
        Next
    
        Call Grid_Refresh_Checkbox(objGridRequisicoes)

    End If
    
    GridRequisicoes_Preenche = SUCESSO
    
    Exit Function
    
Erro_GridRequisicoes_Preenche:

    GridRequisicoes_Preenche = Err
    
    Select Case gErr
    
        Case 63976, 63978, 63980, 178875
            'Erros tratados nas rotinas chamadas

        Case 63977
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case 63979
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_NAO_CADASTRADO", gErr, objRequisitante.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161254)

    End Select
        
End Function

Function ItensConcorrencia_Cria_Altera(objItemRC As ClassItemReqCompras) As Long

Dim lErro As Long
Dim lForn As Long
Dim dFator As Double
Dim bAchou As Boolean
Dim iFilForn As Integer
Dim iPosicao As Integer
Dim objProduto As New ClassProduto
Dim objReqCompra As New ClassRequisicaoCompras
Dim objQuantSupl As ClassQuantSuplementar
Dim dQuantComprar As Double
Dim objCotItemConc As ClassCotacaoItemConc
Dim objItemRCItemConc As ClassItemRCItemConcorrencia
Dim objItemConcorrencia As ClassItemConcorrencia
Dim dQuantReq As Double

On Error GoTo Erro_ItensConcorrencia_Cria_Altera
    
    objProduto.sCodigo = objItemRC.sProduto
    
    'L� os dados do produto envolvido
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 23080 Then gError 62775
    If lErro <> SUCESSO Then gError 62776
    
    'Se o item Rc for exclusivo
    If objItemRC.iExclusivo = MARCADO Then
        'guarda o fornc e filial do item de conc
        lForn = objItemRC.lFornecedor: iFilForn = objItemRC.iFilial
    'Sen�o
    Else
        'O item n�o estar� vinculado a filial fornecedor
        lForn = 0: iFilForn = 0
    End If
        
    'Verica se j� existe um item de concorr�ncia copm os dados
    'determinados pelo item de requisi��o
    bAchou = False
    iPosicao = 0
    For Each objItemConcorrencia In gcolItemConcorrencia
        iPosicao = iPosicao + 1
        
        If objItemConcorrencia.sProduto = objItemRC.sProduto And _
           objItemConcorrencia.lFornecedor = lForn And _
           objItemConcorrencia.iFilial = iFilForn Then
           'Encontrou o item de concorr�ncia
           bAchou = True
           Exit For
        End If
    Next

    'Busca os dados da requisi��o de compra ligada ao ItemRC passado
    Call Obtem_ReqCompra(gobjGeracaoPedCompraReq.colRequisicao, objItemRC.lReqCompra, objReqCompra)
    
    'Faz a convers�o da quantidade a comprar do item para UM compra
    lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemRC.sUM, objProduto.sSiglaUMCompra, dFator)
    If lErro <> SUCESSO Then gError 62777
    
    dQuantComprar = objItemRC.dQuantComprar * dFator
    objItemRC.dQuantNaConcorrencia = dQuantComprar
    
    'Se o item concorr�ncia j� existe
    If bAchou Then
        'recolhe o item de concorr�ncia
        Set objItemConcorrencia = gcolItemConcorrencia(iPosicao)
        
        bAchou = False
        iPosicao = 0
        'Verifica se j� um registro de quant suplementar para o tipo de destino do ItemRC
        For Each objQuantSupl In objItemConcorrencia.colQuantSuplementar
            iPosicao = iPosicao + 1
            If objQuantSupl.iFilialDestino = objReqCompra.iFilialDestino And _
               objQuantSupl.iTipoDestino = objReqCompra.iTipoDestino And _
               objQuantSupl.lFornCliDestino = objReqCompra.lFornCliDestino Then
                'encontrou
                bAchou = True
                Exit For
            End If
        Next
        
        'Se encontrou registro de quant supl.
        If bAchou Then
            'Atualiza a quantidade suplementar
            Set objQuantSupl = objItemConcorrencia.colQuantSuplementar(iPosicao)
            objQuantSupl.dQuantidade = objQuantSupl.dQuantidade + dQuantComprar
            objQuantSupl.dQuantRequisitada = objQuantSupl.dQuantRequisitada + dQuantComprar
        'Sen�o
        Else
            'cria um novo registro de quant suplementar
            Set objQuantSupl = New ClassQuantSuplementar

            objQuantSupl.dQuantidade = dQuantComprar
            objQuantSupl.dQuantRequisitada = dQuantComprar
            objQuantSupl.iFilialDestino = objReqCompra.iFilialDestino
            objQuantSupl.iTipoDestino = objReqCompra.iTipoDestino
            objQuantSupl.lFornCliDestino = objReqCompra.lFornCliDestino
                    
            objItemConcorrencia.colQuantSuplementar.Add objQuantSupl
        End If
                
    ' Se n�o
    Else
        'Cria um novo item de concorr�ncia
        Set objItemConcorrencia = New ClassItemConcorrencia
        
        objItemConcorrencia.iEscolhido = MARCADO
        objItemConcorrencia.iFilial = iFilForn
        objItemConcorrencia.lFornecedor = lForn
        objItemConcorrencia.sProduto = objProduto.sCodigo
        objItemConcorrencia.sDescricao = objProduto.sDescricao
        objItemConcorrencia.sUM = objProduto.sSiglaUMCompra
        objItemConcorrencia.dtDataNecessidade = DATA_NULA
        
        'Cria um registro de quant suplementar p\ o destino da Req do ItemRC
        Set objQuantSupl = New ClassQuantSuplementar
        
        objQuantSupl.dQuantidade = dQuantComprar
        objQuantSupl.dQuantRequisitada = dQuantComprar
        objQuantSupl.iFilialDestino = objReqCompra.iFilialDestino
        objQuantSupl.iTipoDestino = objReqCompra.iTipoDestino
        objQuantSupl.lFornCliDestino = objReqCompra.lFornCliDestino
                
        objItemConcorrencia.colQuantSuplementar.Add objQuantSupl
        
        'Adiciona o novo item de concorr�ncia na cole��o global
        gcolItemConcorrencia.Add objItemConcorrencia
        
    End If
        
    If objReqCompra.dtDataLimite <> DATA_NULA Then
        If (objItemConcorrencia.dtDataNecessidade = DATA_NULA) Or (objReqCompra.dtDataLimite < objItemConcorrencia.dtDataNecessidade) Then objItemConcorrencia.dtDataNecessidade = objReqCompra.dtDataLimite
    End If
    
    If objReqCompra.lUrgente = MARCADO Then objItemConcorrencia.dQuantUrgente = objItemConcorrencia.dQuantUrgente + dQuantComprar
    
    'Cria o link entre o item de req e o item de concorr�ncia
    Set objItemRCItemConc = New ClassItemRCItemConcorrencia
    
    objItemRCItemConc.dQuantidade = dQuantComprar
    objItemRCItemConc.lItemReqCompra = objItemRC.lNumIntDoc
    
    objItemConcorrencia.colItemRCItemConcorrencia.Add objItemRCItemConc
    
    'Atualiza a quantidade do item de concorr�ncia
    objItemConcorrencia.dQuantidade = objItemConcorrencia.dQuantidade + dQuantComprar

    ItensConcorrencia_Cria_Altera = SUCESSO
    
    Exit Function
    
Erro_ItensConcorrencia_Cria_Altera:

    ItensConcorrencia_Cria_Altera = Err
    
    Select Case gErr
    
        Case 62775, 62777
        
        Case 62776
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161255)
            
    End Select
    
    Exit Function

End Function

Private Sub Obtem_ReqCompra(colRequisicao As Collection, lNumIntReq As Long, objReqCompra As ClassRequisicaoCompras)
'Devolve os dados da Requisi��o de compras do Item de Requisi��o de compras passado

Dim objRequisicao As ClassRequisicaoCompras

    'Busca a Requisicao de compras
    For Each objRequisicao In colRequisicao
        'Se � a Requisi��o procurada
        If objRequisicao.lNumIntDoc = lNumIntReq Then
            'Guarda a requisi��o
            Set objReqCompra = objRequisicao
            'Sai da fun��o
            Exit For
        End If
    Next

    Exit Sub

End Sub

Private Sub Escolher_Cotacoes(objItemConcorrencia As ClassItemConcorrencia)
'recebe a cole��o de Itens de cota��o lida do BD e Escolhe para
'o usu�rio aquelas que possuem melhor pre�o ,ou melhor preco + prazo entrega
'como defaut
Dim dMelhorPreco As Double
Dim objCotItemConcMelhor As ClassCotacaoItemConc
Dim objCotItemConc As ClassCotacaoItemConc
Dim dValorPresente As Double
Dim lErro As Long
Dim dTaxa As Double
Dim dValorPresenteReal As Double
Dim objCotacaoMoeda As New ClassCotacaoMoeda
Dim iIndice As Integer
Dim objCondicaoPagto As ClassCondicaoPagto

On Error GoTo Erro_Escolher_Cotacoes
    
    dMelhorPreco = 0
      
    'Se est� amarrado com for e filial --> sai
    If objItemConcorrencia.lFornecedor > 0 And objItemConcorrencia.iFilial > 0 Then Exit Sub
        
    If objItemConcorrencia.colCotacaoItemConc.Count = 0 Then Exit Sub
    
    Set objCotItemConcMelhor = objItemConcorrencia.colCotacaoItemConc(1)
    
    For iIndice = 1 To objItemConcorrencia.colCotacaoItemConc.Count
        
        Set objCotItemConcMelhor = objItemConcorrencia.colCotacaoItemConc(iIndice)
    
        If objCotItemConcMelhor.iMoeda <> MOEDA_REAL Then
            If objCotItemConcMelhor.dTaxa > 0 Then
                dTaxa = objCotItemConcMelhor.dTaxa
                Exit For
            Else
                objCotacaoMoeda.iMoeda = objCotItemConcMelhor.iMoeda
                objCotacaoMoeda.dtData = gdtDataHoje
                
                lErro = CF("CotacaoMoeda_Le", objCotacaoMoeda)
                If lErro <> SUCESSO And lErro <> 80267 Then gError 108983
                If lErro = SUCESSO Then
                    dTaxa = objCotItemConcMelhor.dTaxa
                    Exit For
                End If
            End If
        Else
            dTaxa = 1
            Exit For
        End If
    Next
    
    
    dMelhorPreco = objCotItemConcMelhor.dPrecoUnitario * dTaxa
    
    Set objCondicaoPagto = New ClassCondicaoPagto
    objCondicaoPagto.iCodigo = Codigo_Extrai(objCotItemConcMelhor.sCondPagto)
    
    'Recalcula o Valor Presente
    lErro = CF("Calcula_ValorPresente", objCondicaoPagto, objCotItemConcMelhor.dPrecoAjustado * dTaxa, PercentParaDbl(TaxaEmpresa.Caption), dValorPresenteReal, gdtDataAtual)
    If lErro <> SUCESSO Then gError 62733
    
    objCotItemConcMelhor.iSelecionada = MARCADO
    objCotItemConcMelhor.iEscolhido = MARCADO
    objCotItemConcMelhor.sMotivoEscolha = MOTIVO_MELHORPRECO_DESCRICAO
    
    'Para cada cota��o do item
    For Each objCotItemConc In objItemConcorrencia.colCotacaoItemConc
        
        Set objCondicaoPagto = New ClassCondicaoPagto
        objCondicaoPagto.iCodigo = Codigo_Extrai(objCotItemConc.sCondPagto)
        
        'Recalcula o Valor Presente
        lErro = CF("Calcula_ValorPresente", objCondicaoPagto, objCotItemConc.dPrecoAjustado, PercentParaDbl(TaxaEmpresa.Caption), dValorPresente, gdtDataAtual)
        If lErro <> SUCESSO Then gError 62733

        'Calcula o valor presente
        objCotItemConc.dValorPresente = dValorPresente

        If objCotItemConc.iMoeda <> MOEDA_REAL Then
            If objCotItemConc.dTaxa > 0 Then
                dTaxa = objCotItemConc.dTaxa
            Else
                objCotacaoMoeda.iMoeda = objCotItemConc.iMoeda
                objCotacaoMoeda.dtData = gdtDataHoje
                
                lErro = CF("CotacaoMoeda_Le", objCotacaoMoeda)
                If lErro <> SUCESSO And lErro <> 80267 Then gError 108983

                dTaxa = objCotItemConc.dTaxa
            End If
        Else
            dTaxa = 1
        End If
        
        dValorPresenteReal = dValorPresente * dTaxa
        
        'Se a Cota��o for em Real ou se for em outra moeda para a qual _
         a Cota��o esteja informada ent�o pode-se analisar qual � a _
         melhor op��o de pre�o convertendo todos para Real
        If ((objCotItemConc.iMoeda = MOEDA_REAL) Or (objCotItemConc.iMoeda <> MOEDA_REAL And dTaxa > 0)) Then

            'Se o valor presente � melhor que o menor pre�o at� agora
            If (dValorPresenteReal < dMelhorPreco) Then
    
                objCotItemConcMelhor.sMotivoEscolha = ""
                objCotItemConcMelhor.iEscolhido = DESMARCADO
                objCotItemConcMelhor.iSelecionada = DESMARCADO
                
                'Guarda essa cota��o como a de melhor pre�o
                dMelhorPreco = dValorPresenteReal
                
                Set objCotItemConcMelhor = objCotItemConc
                
                objCotItemConcMelhor.sMotivoEscolha = MOTIVO_MELHORPRECO_DESCRICAO
                objCotItemConcMelhor.iEscolhido = MARCADO
                objCotItemConcMelhor.iSelecionada = MARCADO
    
            'Se o valor for igual ao da cota��o de melhor pre�o
            ElseIf dValorPresenteReal = dMelhorPreco Then
    
                If objCotItemConc.iPrazoEntrega <> 0 And objCotItemConcMelhor.iPrazoEntrega <> 0 Then
                    'Escolhe a cota��o com o melhor prazo de entrega
                    If objCotItemConc.iPrazoEntrega < objCotItemConcMelhor.iPrazoEntrega Then
                                                
                        objCotItemConcMelhor.sMotivoEscolha = ""
                        objCotItemConcMelhor.iEscolhido = DESMARCADO
                        objCotItemConcMelhor.iSelecionada = DESMARCADO
                        
                        dMelhorPreco = objCotItemConc.dValorPresente
                        Set objCotItemConcMelhor = objCotItemConc
                        objCotItemConcMelhor.sMotivoEscolha = MOTIVO_PRECO_PRAZO_DESCRICAO
                        objCotItemConcMelhor.iEscolhido = MARCADO
                        objCotItemConcMelhor.iSelecionada = MARCADO
                    End If
                End If
            End If
        End If
    Next
    
    Exit Sub
    
Erro_Escolher_Cotacoes:

    Select Case gErr
    
        Case 62733
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161256)
            
    End Select
        
    Exit Sub
        
End Sub

Function Grids_Produto_Preenche() As Long

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
    
    'Ordena os itens de concorr�ncia por produto
    lErro = Ordena_Colecao(gcolItemConcorrencia, colItensSaida, colCampos)
    If lErro <> SUCESSO Then gError 63808

    Set gcolItemConcorrencia = colItensSaida
    
    iLinha1 = 0
    iLinha2 = 0
    
    '#######################################################################
    'Inserido por Wagner 25/05/2006
    If gcolItemConcorrencia.Count >= objGridProdutos1.objGrid.Rows Then
        Call Refaz_Grid(objGridProdutos1, gcolItemConcorrencia.Count)
    End If
    '#######################################################################
    
    'Para cada item de concorr�ncia
    For Each objItemConc In gcolItemConcorrencia
        
        iLinha1 = iLinha1 + 1
        'Preenche a sele��o
        GridProdutos1.TextMatrix(iLinha1, iGrid_EscolhidoProduto_Col) = objItemConc.iEscolhido
        
        lErro = Mascara_RetornaProdutoEnxuto(objItemConc.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 62778
        
        Produto1.PromptInclude = False
        Produto1.Text = sProdutoEnxuto
        Produto1.PromptInclude = True
        
        'Preenche o produto
        GridProdutos1.TextMatrix(iLinha1, iGrid_Produto1_Col) = Produto1.Text
        GridProdutos1.TextMatrix(iLinha1, iGrid_DescProduto1_Col) = objItemConc.sDescricao
        GridProdutos1.TextMatrix(iLinha1, iGrid_UnidadeMed1_Col) = objItemConc.sUM
        GridProdutos1.TextMatrix(iLinha1, iGrid_Quantidade1_Col) = Formata_Estoque(objItemConc.dQuantidade)
        GridProdutos1.TextMatrix(iLinha1, iGrid_Urgente_Col) = Formata_Estoque(objItemConc.dQuantUrgente)
        
        'Se o Fornecedor est� preenchido
        If objItemConc.lFornecedor > 0 And objItemConc.iFilial > 0 Then
            
            'verifica se esse forn j� foi lido
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
            
            'Verifica se essa filial j� foi lida
            Call Busca_FilialForn(colFilForn, objItemConc.lFornecedor, objItemConc.iFilial, iPosicao)
            
            If iPosicao = 0 Then
                Set objFilialFornecedor = New ClassFilialFornecedor
                objFilialFornecedor.lCodFornecedor = objItemConc.lFornecedor
                objFilialFornecedor.iCodFilial = objItemConc.iFilial
                
                lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
                If lErro <> SUCESSO And lErro <> 12929 Then gError 63989
                
                'Se n�o encontrou==>Erro
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
    lErro = GridProdutos2_Preenche
    If lErro <> SUCESSO Then gError 62781
    
    Call GridCotacoes_Preenche
    
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161257)
            
    End Select
    
    Exit Function

End Function

Private Function GridCotacoes_Preenche() As Long
'Preenche Grid de Cota��es

Dim lErro As Long
Dim iIndiceMoeda As Integer
Dim objCotacaoMoeda As New ClassCotacaoMoeda
Dim iIndice As Integer, iIndice2 As Integer
Dim colCampos As New Collection
Dim iCondPagto As Integer
Dim colGeracao As New Collection
Dim dValorPresente As Double
Dim colCotacaoSaida As New Collection
Dim sProdutoMascarado As String
Dim objCotItemConcAux As ClassCotacaoItemConcAux
Dim objItemCotItemConc As ClassCotacaoItemConc
Dim objItemConcorrencia As New ClassItemConcorrencia
Dim objCondicaoPagto As ClassCondicaoPagto

On Error GoTo Erro_GridCotacoes_Preenche
    
    Call Grid_Limpa(objGridCotacoes)
           
    For Each objItemConcorrencia In gcolItemConcorrencia
        If objItemConcorrencia.iEscolhido = MARCADO Then
            'Coloca na cole��o as cota��es que aparecem na tela
             For Each objItemCotItemConc In objItemConcorrencia.colCotacaoItemConc
                    
                Set objCotItemConcAux = New ClassCotacaoItemConcAux
                
                Set objCotItemConcAux.objCotacaoItemConc = objItemCotItemConc
                objCotItemConcAux.sCondPagto = objItemCotItemConc.sCondPagto
                objCotItemConcAux.sDescricao = objItemConcorrencia.sDescricao
                objCotItemConcAux.sFilial = objItemCotItemConc.sFilial
                objCotItemConcAux.sFornecedor = objItemCotItemConc.sFornecedor
                objCotItemConcAux.sProduto = objItemConcorrencia.sProduto
                objCotItemConcAux.dtDataNecessidade = objItemConcorrencia.dtDataNecessidade
                
                colGeracao.Add objCotItemConcAux
             Next
        End If
    Next
    
    'Carrega os campos base para a ordena��o utilizados na rotina de ordena��o
    Call Monta_Colecao_Campos_Cotacao(colCampos, OrdenacaoCot.ListIndex)

    If colGeracao.Count > 0 Then
        lErro = Ordena_Colecao(colGeracao, colCotacaoSaida, colCampos)
        If lErro <> SUCESSO Then gError 63808
    End If
    
    Set colGeracao = colCotacaoSaida
    
    iIndice = 0
    
    '#######################################################################
    'Inserido por Wagner 25/05/2006
    If colGeracao.Count >= objGridCotacoes.objGrid.Rows Then
        Call Refaz_Grid(objGridCotacoes, colGeracao.Count)
    End If
    '#######################################################################
    
    For Each objCotItemConcAux In colGeracao

        iIndice = iIndice + 1
        GridCotacoes.TextMatrix(iIndice, iGrid_EscolhidoCot_Col) = objCotItemConcAux.objCotacaoItemConc.iEscolhido

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
        GridCotacoes.TextMatrix(iIndice, iGrid_PrecoUnitarioCot_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dPrecoAjustado, gobjCOM.sFormatoPrecoUnitario) ' "STANDARD") 'Alterado por Wagner
        
        If objCotItemConcAux.objCotacaoItemConc.sMotivoEscolha <> MOTIVO_EXCLUSIVO_DESCRICAO Then
            Call Calcula_Preferencia(objCotItemConcAux.objCotacaoItemConc, Produto1.Text, objCotItemConcAux.objCotacaoItemConc.dQuantidadeComprar)
            GridCotacoes.TextMatrix(iIndice, iGrid_Preferencia_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dPreferencia, "Percent")
        Else
            GridCotacoes.TextMatrix(iIndice, iGrid_Preferencia_Col) = "Exclusivo"
        End If
        
        iCondPagto = Codigo_Extrai(objCotItemConcAux.objCotacaoItemConc.sCondPagto)
        
        'Se a condi��o de pagamento n�o for a vista
        If iCondPagto <> COD_A_VISTA And PercentParaDbl(TaxaEmpresa.Caption) > 0 Then
            
            Set objCondicaoPagto = New ClassCondicaoPagto
            objCondicaoPagto.iCodigo = Codigo_Extrai(objCotItemConcAux.objCotacaoItemConc.sCondPagto)
            
            'Recalcula o Valor Presente
            lErro = CF("Calcula_ValorPresente", objCondicaoPagto, objCotItemConcAux.objCotacaoItemConc.dPrecoAjustado, PercentParaDbl(TaxaEmpresa.Caption), dValorPresente, gdtDataAtual)
            If lErro <> SUCESSO Then gError 62733
            
            If objCotItemConcAux.objCotacaoItemConc.iMoeda <> MOEDA_REAL Then
                objCotItemConcAux.objCotacaoItemConc.dValorPresente = dValorPresente * objCotItemConcAux.objCotacaoItemConc.dTaxa
            Else
                objCotItemConcAux.objCotacaoItemConc.dValorPresente = dValorPresente
            End If
                
        Else
            
            If objCotItemConcAux.objCotacaoItemConc.iMoeda <> MOEDA_REAL Then
                objCotItemConcAux.objCotacaoItemConc.dValorPresente = objCotItemConcAux.objCotacaoItemConc.dPrecoUnitario * objCotItemConcAux.objCotacaoItemConc.dTaxa
            Else
                objCotItemConcAux.objCotacaoItemConc.dValorPresente = objCotItemConcAux.objCotacaoItemConc.dPrecoUnitario
            End If
                
        End If
                                          
        GridCotacoes.TextMatrix(iIndice, iGrid_ValorPresenteCot_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dValorPresente, ValorPresente.Format) 'Alterado por Wagner
        
        If objCotItemConcAux.objCotacaoItemConc.iMoeda <> MOEDA_REAL Then
            GridCotacoes.TextMatrix(iIndice, iGrid_ValorItem_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dPrecoAjustado * objCotItemConcAux.objCotacaoItemConc.dQuantidadeComprar * objCotItemConcAux.objCotacaoItemConc.dTaxa, "STANDARD")
        Else
            GridCotacoes.TextMatrix(iIndice, iGrid_ValorItem_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dPrecoAjustado * objCotItemConcAux.objCotacaoItemConc.dQuantidadeComprar, "STANDARD")
        End If
        
        GridCotacoes.TextMatrix(iIndice, iGrid_FornecedorCot_Col) = objCotItemConcAux.objCotacaoItemConc.sFornecedor
        GridCotacoes.TextMatrix(iIndice, iGrid_FilialFornCot_Col) = objCotItemConcAux.objCotacaoItemConc.sFilial
        GridCotacoes.TextMatrix(iIndice, iGrid_PedidoCot_Col) = objCotItemConcAux.objCotacaoItemConc.lPedCotacao
        If objCotItemConcAux.objCotacaoItemConc.dQuantEntrega > 0 Then GridCotacoes.TextMatrix(iIndice, iGrid_QuantidadeEntrega_Col) = Formata_Estoque(objCotItemConcAux.objCotacaoItemConc.dQuantEntrega)
        
        'Data da Cotacao
        If objCotItemConcAux.objCotacaoItemConc.dtDataPedidoCotacao <> DATA_NULA Then
            GridCotacoes.TextMatrix(iIndice, iGrid_DataCotacaoCot_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dtDataPedidoCotacao, "dd/mm/yyyy")
        End If
    
        For iIndice2 = 0 To TipoTributacaoCot.ListCount - 1
            If objCotItemConcAux.objCotacaoItemConc.iTipoTributacao = TipoTributacaoCot.ItemData(iIndice2) Then
                GridCotacoes.TextMatrix(iIndice, iGrid_TipoTributacaoCot_Col) = TipoTributacaoCot.List(iIndice2)
                Exit For
            End If
        Next
        
        GridCotacoes.TextMatrix(iIndice, iGrid_AliquotaIPI_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dAliquotaIPI, "Percent")
        GridCotacoes.TextMatrix(iIndice, iGrid_AliquotaICMS_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dAliquotaICMS, "Percent")
        
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
                
        'Quantidade a comprar M�xima
        GridCotacoes.TextMatrix(iIndice, iGrid_QuantidadeCot_Col) = Formata_Estoque(objCotItemConcAux.objCotacaoItemConc.dQuantCotada)

        'Motivo escolha
        GridCotacoes.TextMatrix(iIndice, iGrid_MotivoEscolhaCot_Col) = objCotItemConcAux.objCotacaoItemConc.sMotivoEscolha
        
        If objCotItemConcAux.dtDataNecessidade <> DATA_NULA Then
            GridCotacoes.TextMatrix(iIndice, iGrid_DataNecessidade_Col) = Format(objCotItemConcAux.dtDataNecessidade, "dd/mm/yyyy")
        End If
        
        'Moeda
        For iIndiceMoeda = 0 To Moeda.ListCount - 1
            If Moeda.ItemData(iIndiceMoeda) = objCotItemConcAux.objCotacaoItemConc.iMoeda Then
                GridCotacoes.TextMatrix(iIndice, iGrid_Moeda_Col) = Moeda.List(iIndiceMoeda)
                Exit For
            End If
        Next
        
        'TaxaForn
        GridCotacoes.TextMatrix(iIndice, iGrid_TaxaForn_Col) = IIf(objCotItemConcAux.objCotacaoItemConc.dTaxa = 0, "", Format(objCotItemConcAux.objCotacaoItemConc.dTaxa, "#.0000"))
        
        If Moeda.ItemData(iIndiceMoeda) <> MOEDA_REAL Then
            
            'Cotacao
            objCotacaoMoeda.iMoeda = Moeda.ItemData(iIndiceMoeda)
            objCotacaoMoeda.dtData = gdtDataHoje
            
            lErro = CF("CotacaoMoeda_Le", objCotacaoMoeda)
            If lErro <> SUCESSO And lErro <> 80267 Then gError 108983
            
            If objCotacaoMoeda.dValor > 0 Then GridCotacoes.TextMatrix(iIndice, iGrid_CotacaoMoeda_Col) = Format(objCotacaoMoeda.dValor, "#.0000")
            
            'Preco unitario R$
            GridCotacoes.TextMatrix(iIndice, iGrid_PrecoUnitario_RS_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dPrecoUnitario * objCotItemConcAux.objCotacaoItemConc.dTaxa, gobjCOM.sFormatoPrecoUnitario) ' "STANDARD") 'Alterado por Wagner
        Else
            'Preco unitario R$
            GridCotacoes.TextMatrix(iIndice, iGrid_PrecoUnitario_RS_Col) = Format(objCotItemConcAux.objCotacaoItemConc.dPrecoUnitario, gobjCOM.sFormatoPrecoUnitario) ' "STANDARD") 'Alterado por Wagner
            
        End If
        
        objGridCotacoes.iLinhasExistentes = objGridCotacoes.iLinhasExistentes + 1
        
    Next

    Call Grid_Refresh_Checkbox(objGridCotacoes)
    
    Call Calcula_TotalItens
    
    Exit Function

Erro_GridCotacoes_Preenche:

    Select Case gErr

        Case 62733, 63808, 68358, 108983
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161258)

    End Select

    Exit Function

End Function

Sub Monta_Colecao_Campos_Cotacao(colCampos As Collection, iOrdenacao As Integer)
'monta a cole��o de campos para a ordena��o

Dim objCotacaoItemConc As New ClassCotacaoItemConc
Dim objItemConcorrencia As New ClassItemConcorrencia

    Select Case iOrdenacao

        Case 0

            colCampos.Add "sProduto"
            colCampos.Add "sCondPagto"
            colCampos.Add "sFornecedor"
            colCampos.Add "sFilial"

        Case 1

            colCampos.Add "sFornecedor"
            colCampos.Add "sFilial"
            colCampos.Add "sProduto"
            colCampos.Add "sCondPagto"

    End Select

End Sub

Private Sub Calcula_TotalItens()
'Calcula o valor total dos itens selecionados

Dim dTotalItens As Double
Dim iIndice As Integer
    
    dTotalItens = 0
    
    For iIndice = 1 To objGridCotacoes.iLinhasExistentes
        If StrParaInt(GridCotacoes.TextMatrix(iIndice, iGrid_EscolhidoCot_Col)) = MARCADO Then
            dTotalItens = dTotalItens + StrParaDbl(GridCotacoes.TextMatrix(iIndice, iGrid_ValorItem_Col))
        End If
    Next

    TotalItens.Caption = Format(dTotalItens, "STANDARD")
    
    Exit Sub

End Sub

Function ItensConcorrencia_Atualiza(objReqCompra As ClassRequisicaoCompras, objItemRC As ClassItemReqCompras)

Dim lErro As Long
Dim objItemConcorrencia As ClassItemConcorrencia
Dim objItemRCOutros As ClassItemReqCompras
Dim objReqCompraOutras As ClassRequisicaoCompras
Dim iItem As Integer

On Error GoTo Erro_ItensConcorrencia_Atualiza
    
    'Localiza o item de concorr�ncia correspondente
    Call Localiza_ItemConcorrencia(gcolItemConcorrencia, objItemConcorrencia, iItem, objItemRC)
    
    'Se a requisi��o est� sendo desmarcada
    If objReqCompra.iSelecionado = DESMARCADO Then
        'Se o item da requisi��o est� marcado
        If objItemRC.iSelecionado = MARCADO And iItem > 0 Then
            lErro = ItemConcorrencia_Exclui_QuantComprar(objItemConcorrencia, iItem, objReqCompra, objItemRC)
            If lErro <> SUCESSO Then gError 62782
            
        End If
    'se a requisicao est� marcada
    Else
        
        If objItemRC.iSelecionado = MARCADO Then
            'Inclui os dados do item de requisicao
            lErro = ItensConcorrencia_Cria_Altera(objItemRC)
            If lErro <> SUCESSO Then gError 62782
                    
        ElseIf iItem > 0 Then
            
            Set objItemConcorrencia = gcolItemConcorrencia(iItem)
            
            lErro = ItemConcorrencia_Exclui_QuantComprar(objItemConcorrencia, iItem, objReqCompra, objItemRC)
            If lErro <> SUCESSO Then gError 62783
        
        End If
    
    End If
    
    Call Localiza_ItemConcorrencia(gcolItemConcorrencia, objItemConcorrencia, iItem, objItemRC)

    If iItem > 0 Then
        'Renova as cotacoes dos itens alterados
        lErro = ItemConcorrencia_Atualiza_Cotacoes(gcolItemConcorrencia, iItem)
        If lErro <> SUCESSO Then gError 62784
    End If
    
    ItensConcorrencia_Atualiza = SUCESSO
        
    Exit Function
    
Erro_ItensConcorrencia_Atualiza:

    ItensConcorrencia_Atualiza = Err
    
    Select Case gErr
    
        Case 62782, 62783, 62784
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161259)
            
    End Select
    
    Exit Function
    
End Function

Function ItemConcorrencia_Exclui_QuantComprar(objItemConcorrencia As ClassItemConcorrencia, iItem As Integer, Optional objReqCompra As ClassRequisicaoCompras, Optional objItemRC As ClassItemReqCompras, Optional dQuantidade As Double = 0)

Dim iIndice As Integer
Dim objItemRCItemConc As ClassItemRCItemConcorrencia
Dim objQtSupl As ClassQuantSuplementar
Dim lErro As Long
Dim bExclui As Boolean
Dim objProduto As New ClassProduto
Dim dFator As Double
    
On Error GoTo Erro_ItemConcorrencia_Exclui_QuantComprar
    
    'Se a quantidade n�o foi passada
    If dQuantidade = 0 Then
        
        objProduto.sCodigo = objItemRC.sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 23080 Then gError 62785
        If lErro <> SUCESSO Then gError 62786
        
        lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemRC.sUM, objProduto.sSiglaUMCompra, dFator)
        If lErro <> SUCESSO Then gError 62787
        
        'A quantidade a exclui � a do ItemRC passado
        dQuantidade = objItemRC.dQuantComprar * dFator
        'Exclui a liga��o do item RC com o item conc
        bExclui = True
    End If

    'diminui a quantidade a comprar do item de concorrencia vinculado
    objItemConcorrencia.dQuantidade = objItemConcorrencia.dQuantidade - dQuantidade
    
       
    iIndice = 0
    
    'Se algum item de req foi passado
    If Not (objItemRC Is Nothing) Then
        'Exclui o vinculo entre o item de requisicao e o item de concorrencia
        For Each objItemRCItemConc In objItemConcorrencia.colItemRCItemConcorrencia
            iIndice = iIndice + 1
            'BUsca o vinculo do ItemRc e ItemConc
            If objItemRCItemConc.lItemReqCompra = objItemRC.lNumIntDoc Then
                'Se a quant do item foi toda exclu�da
                If bExclui Then
                    'exclui o link entre o item RC e o item conc
                    objItemConcorrencia.colItemRCItemConcorrencia.Remove iIndice
                'sen�o
                Else
                    'Diminui a quantidade exclu�da
                    objItemRCItemConc.dQuantidade = objItemRCItemConc.dQuantidade - dQuantidade
                End If
                            
                Exit For
            End If
        Next

        iIndice = 0
        'Diminui a quantidade a comprar do correspondente em quant suplementares
        For Each objQtSupl In objItemConcorrencia.colQuantSuplementar
            iIndice = iIndice + 1
            If objQtSupl.iTipoDestino = objReqCompra.iTipoDestino And objQtSupl.iFilialDestino = objReqCompra.iFilialDestino And objQtSupl.lFornCliDestino = objReqCompra.lFornCliDestino Then
                objQtSupl.dQuantidade = objQtSupl.dQuantidade - dQuantidade
                objQtSupl.dQuantRequisitada = objQtSupl.dQuantRequisitada - dQuantidade
                If objQtSupl.dQuantidade <= 0 Then objItemConcorrencia.colQuantSuplementar.Remove iIndice
                Exit For
            End If
        Next
        If objReqCompra.lUrgente = MARCADO Then objItemConcorrencia.dQuantUrgente = objItemConcorrencia.dQuantUrgente - dQuantidade
    End If
        
    'Se o item de concorrencia n�o est� vinculado a nenum outro itemRC
    If objItemConcorrencia.colItemRCItemConcorrencia.Count = 0 Then
        'Exclui o item de concorr�ncia
        gcolItemConcorrencia.Remove iItem
    Else
        
        If iItem = 0 Then
            'Altera os dados de compra dos itens de concorr6encia
            '(inclusive cota��es, se necess�rio)
            lErro = ItensConcorrencia_Cria_Altera(objItemRC)
            If lErro <> SUCESSO Then gError 62739
        End If
                
    End If
    
    ItemConcorrencia_Exclui_QuantComprar = SUCESSO
    
    Exit Function

Erro_ItemConcorrencia_Exclui_QuantComprar:

    ItemConcorrencia_Exclui_QuantComprar = Err
    
    Select Case gErr
    
        Case 62739, 62785, 62787
        
        Case 62786
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161260)
            
    End Select
    
    Exit Function

End Function

Function ItemConcorrencia_Inclui_QuantComprar(objItemConcorrencia As ClassItemConcorrencia, iItem As Integer, Optional objReqCompra As ClassRequisicaoCompras, Optional objItemRC As ClassItemReqCompras, Optional dQuantidade As Double)

Dim iIndice As Integer
Dim objItemRCItemConc As ClassItemRCItemConcorrencia
Dim objQtSupl As ClassQuantSuplementar
Dim lErro As Long
Dim bAchou As Boolean

On Error GoTo Erro_ItemConcorrencia_Inclui_QuantComprar

    'Se o item j� foi passado atualizado
    If iItem > 0 Then

        'diminui a quantidade a comprar do item de concorrencia vinculado
        objItemConcorrencia.dQuantidade = objItemConcorrencia.dQuantidade + dQuantidade
        
        If Not (objItemRC Is Nothing) Then
            iIndice = 0
            'Atualiza o vinculo entre o item de requisicao e o item de concorrencia
            For Each objItemRCItemConc In objItemConcorrencia.colItemRCItemConcorrencia
                iIndice = iIndice + 1
                If objItemRCItemConc.lItemReqCompra = objItemRC.lNumIntDoc Then
                    objItemRCItemConc.dQuantidade = objItemRCItemConc.dQuantidade + dQuantidade
                    Exit For
                End If
            Next
            
            iIndice = 0
            'Aumenta a quantidade a comprar do correspondente em quant suplementares
            For Each objQtSupl In objItemConcorrencia.colQuantSuplementar
                iIndice = iIndice + 1
                If objQtSupl.iTipoDestino = objReqCompra.iTipoDestino And objQtSupl.iFilialDestino = objReqCompra.iFilialDestino And objQtSupl.lFornCliDestino = objReqCompra.lFornCliDestino Then
                    bAchou = True
                    objQtSupl.dQuantidade = objQtSupl.dQuantidade + dQuantidade
                    objQtSupl.dQuantRequisitada = objQtSupl.dQuantRequisitada + dQuantidade
                    Exit For
                End If
            Next
            'Se n�o h� quant suplementar p\ esse destino
            If Not bAchou Then
                'Cria um registro de quant siplementar novo
                Set objQtSupl = New ClassQuantSuplementar
                
                objQtSupl.dQuantidade = dQuantidade
                objQtSupl.dQuantRequisitada = dQuantidade
                objQtSupl.iFilialDestino = objReqCompra.iFilialDestino
                objQtSupl.iTipoDestino = objReqCompra.iTipoDestino
                objQtSupl.lFornCliDestino = objReqCompra.lFornCliDestino
                
                If objReqCompra.lUrgente = MARCADO Then objItemConcorrencia.dQuantidade = objItemConcorrencia.dQuantidade + dQuantidade
            
                objItemConcorrencia.colQuantSuplementar.Add objQtSupl
            End If
        
        End If
    
    Else
        
        lErro = ItensConcorrencia_Cria_Altera(objItemRC)
        If lErro <> SUCESSO Then gError 62739
    End If
                    
    ItemConcorrencia_Inclui_QuantComprar = SUCESSO
    
    Exit Function
    
Erro_ItemConcorrencia_Inclui_QuantComprar:

    ItemConcorrencia_Inclui_QuantComprar = Err
    
    Select Case gErr
        
        Case 62739
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161261)
        
    End Select

    Exit Function

End Function

Private Sub Localiza_ItemConcorrencia(colItemConcorrencia As Collection, objItemConcorrencia As ClassItemConcorrencia, iItem As Integer, objItemRC As ClassItemReqCompras)
'Devolve os dados do item de concorrecia ligado ao ItemRc passado

Dim objItemConcAux As ClassItemConcorrencia
Dim lForn As Long, iFilForn As Integer
Dim iIndice As Integer, bAchou As Boolean

    iItem = 0
    iIndice = 0
    bAchou = False

    'Busca nos itens de concorrencia
    For Each objItemConcAux In colItemConcorrencia
        iIndice = iIndice + 1
        'Se for exclusivo
        If objItemRC.iExclusivo = MARCADO Then
            lForn = objItemRC.lFornecedor
            iFilForn = objItemRC.iFilial
        Else
            lForn = 0
            iFilForn = 0
        End If
        If objItemConcAux.sProduto = objItemRC.sProduto And objItemConcAux.lFornecedor = lForn And objItemConcAux.iFilial = iFilForn Then
           Set objItemConcorrencia = objItemConcAux
           'encontrou
           bAchou = True
           Exit For
        End If
    Next

    If bAchou Then iItem = iIndice

    Exit Sub

End Sub

Private Function GridProdutos2_Preenche() As Long
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
Dim iCount As Integer

On Error GoTo Erro_GridProdutos2_Preenche
    
    'Limpa o grid de produtos2
    Call Grid_Limpa(objGridProdutos2)
    
    iLinha1 = 0
    iLinha2 = 0
    
    '#######################################################################
    'Inserido por Wagner 25/05/2006
    For Each objItemConc In gcolItemConcorrencia
        iCount = iCount + objItemConc.colQuantSuplementar.Count
    Next
    
    If iCount >= objGridProdutos2.objGrid.Rows Then
        Call Refaz_Grid(objGridProdutos2, iCount)
    End If
    '#######################################################################
    
    'Para cada item de conc
    For Each objItemConc In gcolItemConcorrencia
        iLinha1 = iLinha1 + 1
        If objItemConc.iEscolhido = MARCADO Then
            
            'Para cada quant supl
            For Each objQuantSupl In objItemConc.colQuantSuplementar
            
                iLinha2 = iLinha2 + 1
                'Preenche com os dados do item de conorr�ncia
                GridProdutos2.TextMatrix(iLinha2, iGrid_Produto2_Col) = GridProdutos1.TextMatrix(iLinha1, iGrid_Produto1_Col)
                GridProdutos2.TextMatrix(iLinha2, iGrid_DescProduto2_Col) = objItemConc.sDescricao
                GridProdutos2.TextMatrix(iLinha2, iGrid_UnidadeMed2_Col) = objItemConc.sUM
                GridProdutos2.TextMatrix(iLinha2, iGrid_Quantidade2_Col) = Formata_Estoque(objQuantSupl.dQuantidade)
                  
                If objQuantSupl.iTipoDestino = TIPO_DESTINO_EMPRESA Then
                    
                    Call Busca_Na_Colecao(colFilEmp, objQuantSupl.iFilialDestino, iPosicao)
                    
                    If iPosicao = 0 Then
                    
                        objFilEmp.lCodEmpresa = glEmpresa
                        objFilEmp.iCodFilial = objQuantSupl.iFilialDestino
                                                                
                        lErro = CF("FilialEmpresa_Le", objFilEmp)
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
                        
                        'L� o fornecedor
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
                    
                        'Se n�o encontrou==>Erro
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
        End If
    Next
    
    objGridProdutos2.iLinhasExistentes = iLinha2

    Call Grid_Refresh_Checkbox(objGridProdutos2)
    
    GridProdutos2_Preenche = SUCESSO
    
    Exit Function
    
Erro_GridProdutos2_Preenche:

    GridProdutos2_Preenche = Err
    
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

Function Atualiza_QuantSupl(objItemConcorrencia As ClassItemConcorrencia, dQuantDiferenca As Double, iLinhaProd2 As Integer)
'Atualiza a cole�ao de quantidades suplementares

Dim lErro As Long
Dim objQuantSupl As ClassQuantSuplementar
Dim lForn As Long
Dim iFilial As Integer
Dim iTipo As Integer
Dim objFornecedor As ClassFornecedor

On Error GoTo Erro_Atualiza_QuantSupl

    lForn = 0
    iFilial = Codigo_Extrai(GridProdutos2.TextMatrix(iLinhaProd2, iGrid_FilialDestino_Col))
    
    'Recolhe o tipo de destino
    If GridProdutos2.TextMatrix(iLinhaProd2, iGrid_TipoDestino_Col) = "Empresa" Then
        iTipo = TIPO_DESTINO_EMPRESA
    Else
        iTipo = TIPO_DESTINO_FORNECEDOR
        
        Set objFornecedor = New ClassFornecedor
        
        objFornecedor.sNomeReduzido = GridProdutos2.TextMatrix(iLinhaProd2, iGrid_Destino_Col)
        'L� o fornecdor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 62773
        If lErro <> SUCESSO Then gError 62774
        
        lForn = objFornecedor.lCodigo
    End If
    
    'Localiza o registro de quant supl correspondente
    For Each objQuantSupl In objItemConcorrencia.colQuantSuplementar
        
        If objQuantSupl.iFilialDestino = iFilial And objQuantSupl.lFornCliDestino = lForn And objQuantSupl.iTipoDestino = iTipo Then
            'Atualiza a quantidade
            If (objQuantSupl.dQuantidade + dQuantDiferenca) < objQuantSupl.dQuantRequisitada Then gError 62772
            objQuantSupl.dQuantidade = objQuantSupl.dQuantidade + dQuantDiferenca
        End If
    Next

    Atualiza_QuantSupl = SUCESSO

    Exit Function

Erro_Atualiza_QuantSupl:

    Atualiza_QuantSupl = Err

    Select Case gErr
        
        Case 62772
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTCOMPRAR_MENOR_QUANTCOMPRAR_RC", gErr, (objQuantSupl.dQuantidade + dQuantDiferenca), objQuantSupl.dQuantRequisitada)
        
        Case 62773
        
        Case 62774
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161262)
            
    
    
    End Select
    
    Exit Function
        
End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)
    
    Select Case objControl.Name
    
        Case QuantComprarItemRC.Name
        
            If StrParaInt(GridItensRequisicoes.TextMatrix(iLinha, iGrid_EscolhidoItem_Col)) = DESMARCADO Then
                QuantComprarItemRC.Enabled = False
            Else
                QuantComprarItemRC.Enabled = True
            End If
    
        'MotivoEscolha
        Case MotivoEscolhaCot.Name

            If objControl.Name = MotivoEscolhaCot.Name And _
               GridCotacoes.TextMatrix(iLinha, iGrid_MotivoEscolhaCot_Col) = MOTIVO_EXCLUSIVO_DESCRICAO Then
               objControl.Enabled = False
            Else
               objControl.Enabled = True
            End If
 
        Case Quantidade2.Name
            
            If giPodeAumentarQuant = MARCADO Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
                
    End Select

    Exit Sub

End Sub

Private Sub Localiza_ItemCotacao(objCotItemConc As ClassCotacaoItemConc, iLinha As Integer)
    
Dim sFornecedor As String
Dim sFilial As String
Dim sMotivo As String
Dim sProduto As String
Dim sCondPagto As String
Dim iIndice As Integer
Dim iItemConc As Integer
Dim objItemConcorrencia As ClassItemConcorrencia
Dim objCotItemConc2 As ClassCotacaoItemConc
Dim iMoeda As Integer
    
    'Recolhe os campos que amarram  uma cota��o na tela
    sMotivo = GridCotacoes.TextMatrix(iLinha, iGrid_MotivoEscolhaCot_Col)
    sProduto = GridCotacoes.TextMatrix(iLinha, iGrid_ProdutoCot_Col)
    sCondPagto = GridCotacoes.TextMatrix(iLinha, iGrid_CondPagtoCot_Col)
    sFornecedor = GridCotacoes.TextMatrix(iLinha, iGrid_FornecedorCot_Col)
    sFilial = GridCotacoes.TextMatrix(iLinha, iGrid_FilialFornCot_Col)
    
    For iIndice = 0 To Moeda.ListCount - 1
        If Moeda.List(iIndice) = GridCotacoes.TextMatrix(iLinha, iGrid_Moeda_Col) Then
            iMoeda = Moeda.ItemData(iIndice)
            Exit For
        End If
    Next
    
    'Se for exclusivo
    If sMotivo = MOTIVO_EXCLUSIVO_DESCRICAO Then
        
        'Para cada item de concorrencia
        For iIndice = 1 To objGridProdutos1.iLinhasExistentes
            
            'Busca o item com forn e filial amarrados
            If GridProdutos1.TextMatrix(iIndice, iGrid_Produto1_Col) = sProduto And _
               GridProdutos1.TextMatrix(iIndice, iGrid_Fornecedor1_Col) = sFornecedor And _
               GridProdutos1.TextMatrix(iIndice, iGrid_FilialForn1_Col) = sFilial Then
                
                iItemConc = iIndice
        
            End If
        
        Next
        
    Else
        
        For iIndice = 1 To objGridProdutos1.iLinhasExistentes
            'Busca o item de concorr�ncia ligado a cota��o
            If GridProdutos1.TextMatrix(iIndice, iGrid_Produto1_Col) = sProduto And _
               Len(Trim(GridProdutos1.TextMatrix(iIndice, iGrid_FilialForn1_Col))) = 0 Then
                
                iItemConc = iIndice
        
            End If
        
        Next
    
    End If
    
    'Seleciona o item de concorr�ncia
    Set objItemConcorrencia = gcolItemConcorrencia(iItemConc)
    
    'Busca dentro das cota��es do item de concorr�ncia a cota��o em quest�o
    For Each objCotItemConc2 In objItemConcorrencia.colCotacaoItemConc
        
        If objCotItemConc2.sFornecedor = sFornecedor And _
           objCotItemConc2.sFilial = sFilial And objCotItemConc2.sCondPagto = sCondPagto And _
            objCotItemConc2.iMoeda = iMoeda Then
            
            Set objCotItemConc = objCotItemConc2
            Exit For
        
        End If
    Next
    
End Sub

Function Valida_Quantidade(objItemConcorrencia As ClassItemConcorrencia, iItem As Integer) As Long
'Verifica se os campos da tela foram preenchidos corretamente

Dim lErro As Long
Dim dQuantidade As Double
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim objCotItemConc As ClassCotacaoItemConc
Dim dQuantComprar As Double
Dim iTot As Integer

On Error GoTo Erro_Valida_Quantidade

    If objItemConcorrencia.colCotacaoItemConc.Count = 0 Then gError 63759
    
    iTot = 0

    objProduto.sCodigo = objItemConcorrencia.sProduto

    'L� o produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 23080 Then gError 62712
    If lErro <> SUCESSO Then gError 62713 'n�o encontrou

    'Recolhe a quantidade do grid
    dQuantidade = objItemConcorrencia.dQuantidade

    lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemConcorrencia.sUM, objProduto.sSiglaUMCompra, dFator)
    If lErro <> SUCESSO Then gError 62714

    dQuantidade = dQuantidade * dFator

    dQuantComprar = 0

    'Percorre as cota��es
    For Each objCotItemConc In objItemConcorrencia.colCotacaoItemConc
        objCotItemConc.iSelecionada = MARCADO
        If objCotItemConc.iEscolhido = MARCADO Then
            iTot = iTot + 1
            dQuantComprar = dQuantComprar + objCotItemConc.dQuantidadeComprar
            If objCotItemConc.dPrecoAjustado = 0 Then gError 70498
        End If
    Next
    
    If iTot = 0 Then gError 63759

    If Abs(Formata_Estoque(dQuantComprar - dQuantidade)) >= QTDE_ESTOQUE_DELTA Then gError 63811

    Valida_Quantidade = SUCESSO

    Exit Function

Erro_Valida_Quantidade:

    Valida_Quantidade = gErr

    Select Case gErr

        Case 62712, 62714
        
        Case 62713
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case 63759
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_NAO_VINCULADO_ITEMCOTACAO", gErr, iItem)

        Case 63811
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTCOTACAO_DIFERENTE_QUANTCOMPRAR", gErr, objProduto.sCodigo)

        Case 70498
            Call Rotina_Erro(vbOKOnly, "ERRO_PRECOUNITARIO_ITEMCOTACAO_NAO_PREENCHIDO", gErr, iItem)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161263)

    End Select

    Exit Function
    
End Function

Function ItemConcorrencia_Atualiza_Cotacoes(colItemConcorrencia As Collection, iItem As Integer) As Long
'Atualiza as cota��es para o item passado

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim bPrecisa_Ler As Boolean
Dim objItemConcorrencia As ClassItemConcorrencia
Dim iTipoTributacao As Integer
Dim lItemMaior As Long
Dim lNumIntItem As Long
Dim objCotItemConc As ClassCotacaoItemConc
Dim objItemRC As New ClassItemReqCompras
Dim objReqCompra As New ClassRequisicaoCompras
Dim iIndice As Integer

On Error GoTo Erro_ItemConcorrencia_Atualiza_Cotacoes

    bPrecisa_Ler = True

    'recolhe o Item de concorr�ncia
    Set objItemConcorrencia = gcolItemConcorrencia(iItem)
    
    lItemMaior = 1
    lNumIntItem = objItemConcorrencia.colItemRCItemConcorrencia(1).lItemReqCompra

    For iIndice = 1 To objItemConcorrencia.colItemRCItemConcorrencia.Count
        If objItemConcorrencia.colItemRCItemConcorrencia(iIndice).dQuantidade > objItemConcorrencia.colItemRCItemConcorrencia(lItemMaior).dQuantidade Then
            lItemMaior = iIndice
            lNumIntItem = objItemConcorrencia.colItemRCItemConcorrencia(iIndice).lItemReqCompra
        End If
    Next

    'L� o Produto
    objProduto.sCodigo = objItemConcorrencia.sProduto

    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 23080 Then gError 62791
    If lErro <> SUCESSO Then gError 62792

    If objProduto.iConsideraQuantCotAnt <> PRODUTO_CONSIDERA_QUANT_COTACAO_ANTERIOR And _
       objItemConcorrencia.colCotacaoItemConc.Count > 0 Then bPrecisa_Ler = False

    If bPrecisa_Ler Then
        
        Set objItemConcorrencia.colCotacaoItemConc = New Collection
                
        lErro = CF("Cotacoes_Produto_Le", objItemConcorrencia.colCotacaoItemConc, objProduto, objItemConcorrencia.dQuantidade, gobjGeracaoPedCompraReq.iTipoDestino, gobjGeracaoPedCompraReq.lFornCliDestino, gobjGeracaoPedCompraReq.iFilialDestino, objItemConcorrencia.lFornecedor, objItemConcorrencia.iFilial)
        If lErro <> SUCESSO And lErro <> 63822 Then gError 62793
        
        Call Escolher_Cotacoes(objItemConcorrencia)
    Else
        
        For Each objCotItemConc In objItemConcorrencia.colCotacaoItemConc
            objCotItemConc.dQuantidadeComprar = objItemConcorrencia.dQuantidade
        Next
        
        Call Escolher_Cotacoes(objItemConcorrencia)
    
    End If
    
    Call Localiza_ItemReqCompra(gobjGeracaoPedCompraReq.colRequisicao, lNumIntItem, objItemRC, objReqCompra)
    
    If objItemConcorrencia.colCotacaoItemConc.Count > 0 Then
        For Each objCotItemConc In objItemConcorrencia.colCotacaoItemConc
            objCotItemConc.iTipoTributacao = objItemRC.iTipoTributacao
        Next
    End If
    
    ItemConcorrencia_Atualiza_Cotacoes = SUCESSO

    Exit Function
    
Erro_ItemConcorrencia_Atualiza_Cotacoes:

    ItemConcorrencia_Atualiza_Cotacoes = Err
    
    Select Case gErr
    
        Case 62791, 62793
        
        Case 62792
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161264)
    
    End Select
    
    Exit Function

End Function

Private Sub Adiciona_Codigo(colIndices As Collection, iItem As Integer)
'se o c�digo passado n�o estiver na cole��o ele � adiconado
Dim iIndice As Integer

    For iIndice = 1 To colIndices.Count
        If colIndices(iIndice) = iItem Then Exit Sub
    Next
        
    colIndices.Add iItem

    Exit Sub
    
End Sub

Function colItensCotacao_Adiciona(lItemCotacao As Long, colItensCotacao As Collection) As Long
'Se o Item de cota��o n�o existe na cole��o ele � lido e inclu�do

Dim objItemCotacao As ClassItemCotacao
Dim bAchou As Boolean
Dim lErro As Long

On Error GoTo Erro_colItensCotacao_Adiciona

    bAchou = False
    'Busca o Item de cota��o
    For Each objItemCotacao In colItensCotacao
        If objItemCotacao.lNumIntDoc = lItemCotacao Then
            bAchou = True
            Exit For
        End If
    Next
    
    If Not bAchou Then
        Set objItemCotacao = New ClassItemCotacao
        
        objItemCotacao.lNumIntDoc = lItemCotacao
        'L� o Item cota��o
        lErro = CF("ItemCotacao_Le", objItemCotacao)
        If lErro <> SUCESSO Then gError 62725
        
        'Adiciona na cole��o
        colItensCotacao.Add objItemCotacao, CStr(objItemCotacao.lNumIntDoc)

    End If
    
    colItensCotacao_Adiciona = SUCESSO
    
    Exit Function

Erro_colItensCotacao_Adiciona:

    colItensCotacao_Adiciona = Err
    
    Select Case gErr
    
        Case 62725
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161265)
    
    End Select

    Exit Function

End Function

Function Inclui_Quant_ItemReqCompra(objItemPC As ClassItemPedCompra, objItemConcorrencia As ClassItemConcorrencia, objQuantSupl As ClassQuantSuplementar, colRequisicao As Collection, colProdutos As Collection)

Dim lErro As Long
Dim dQuantidade As Double
Dim objItemReqCompra As ClassItemReqCompras
Dim objItemRCItemConc As ClassItemRCItemConcorrencia
Dim dDiferenca As Double
Dim objItemRC As ClassItemReqCompras
Dim objReqCompra As New ClassRequisicaoCompras
Dim objLocItemPC As ClassLocalizacaoItemPC
Dim bAchou As Boolean, dFatorCOM As Double
Dim objProduto As New ClassProduto

On Error GoTo Erro_Inclui_Quant_ItemReqCompra

    Call Busca_Produto(objItemPC.sProduto, colProdutos, objProduto, bAchou)

    If Not bAchou Then
    
        Set objProduto = New ClassProduto
        
        objProduto.sCodigo = objItemPC.sProduto
    
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 86147
        If lErro <> SUCESSO Then gError 86149
    
        colProdutos.Add objProduto
    
    End If
    
    dQuantidade = objItemPC.dQuantidade

    'Para cada item de req que gerou esse item de concorr�ncia
    For Each objItemRCItemConc In objItemConcorrencia.colItemRCItemConcorrencia

        'Busca os dados do item
        Call Localiza_ItemReqCompra(colRequisicao, objItemRCItemConc.lItemReqCompra, objItemReqCompra, objReqCompra)

        'Se o item acessado � do mesmo tipo de destino do PC
        If objReqCompra.iTipoDestino = objQuantSupl.iTipoDestino And objReqCompra.lFornCliDestino = objQuantSupl.lFornCliDestino And objQuantSupl.iFilialDestino = objReqCompra.iFilialDestino And (objItemReqCompra.dQuantComprar - objItemReqCompra.dQuantNoPedido > 0) Then
                    
            lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemReqCompra.sUM, objItemPC.sUM, dFatorCOM)
            If lErro <> SUCESSO Then gError 86148
            
            'Calcula a diferen�a entre a qt do ItemPC e a qt � associada do ItemRC
            dDiferenca = dQuantidade - ((objItemReqCompra.dQuantComprar - objItemReqCompra.dQuantNoPedido) * dFatorCOM)

            'Cria um objItemRC
            Set objItemRC = New ClassItemReqCompras

            'recolhe alguns dados
            objItemRC.lNumIntDoc = objItemReqCompra.lNumIntDoc
            objItemRC.iAlmoxarifado = objItemReqCompra.iAlmoxarifado
            objItemRC.sProduto = objItemReqCompra.sProduto
            objItemRC.sUM = objItemReqCompra.sUM
            objItemRC.sCcl = objItemReqCompra.sCcl
            objItemRC.sDescProduto = objItemReqCompra.sDescProduto
            objItemRC.sContaContabil = objItemReqCompra.sContaContabil

            'se a diferen�a for positiva
            If dDiferenca >= 0 Then
                'A quantidade do item q n�o est� associada a ItemPC ser� utilizada
                objItemRC.dQuantComprar = objItemReqCompra.dQuantComprar - objItemReqCompra.dQuantNoPedido
                objItemReqCompra.dQuantNoPedido = objItemReqCompra.dQuantComprar
            'se for negativa
            Else
                'Parte da quantidade do item q n�o est� associada a ItemPC ser� utilizada
                objItemRC.dQuantComprar = dQuantidade / dFatorCOM
                objItemReqCompra.dQuantNoPedido = objItemReqCompra.dQuantNoPedido + (dQuantidade / dFatorCOM)
            End If

            If objItemRC.iAlmoxarifado > 0 Then

                bAchou = False
                For Each objLocItemPC In objItemPC.colLocalizacao
                    If objLocItemPC.iAlmoxarifado = objItemRC.iAlmoxarifado Then
                        bAchou = True
                        objLocItemPC.dQuantidade = objLocItemPC.dQuantidade + (objItemRC.dQuantComprar * dFatorCOM)
                    End If
                Next

                If Not bAchou Then
                    Set objLocItemPC = New ClassLocalizacaoItemPC

                    objLocItemPC.dQuantidade = (objItemRC.dQuantComprar * dFatorCOM)
                    objLocItemPC.iAlmoxarifado = objItemRC.iAlmoxarifado
                    objLocItemPC.sCcl = objItemRC.sCcl
                    objLocItemPC.sContaContabil = objItemRC.sContaContabil

                    objItemPC.colLocalizacao.Add objLocItemPC
                End If
            End If

            objItemPC.colItemReqCompras.Add objItemRC
            'Atualiza a quantidade que falta associar a ItemPC
            dQuantidade = dQuantidade - (objItemRC.dQuantComprar * dFatorCOM)

            'Se j� associou toda a quantidade, sai
            If dQuantidade = 0 Then Exit Function

        End If

    Next

    Inclui_Quant_ItemReqCompra = SUCESSO

    Exit Function

Erro_Inclui_Quant_ItemReqCompra:

    Inclui_Quant_ItemReqCompra = gErr

    Select Case gErr
    
        Case 86147, 86148
        
        Case 86149
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161266)

    End Select

    Exit Function

End Function

Private Sub Localiza_ItemReqCompra(colRequisicao As Collection, lItemReqCompra As Long, objItemReqCompra As ClassItemReqCompras, objReqCompra As ClassRequisicaoCompras)
'Localiza o Item de Requisicao com o numero interno passado

Dim iIndice As Integer
Dim objItemRC As ClassItemReqCompras
    
    'Para cada Requsiicao
    For iIndice = 1 To colRequisicao.Count
        Set objReqCompra = colRequisicao(iIndice)
        'Para cada item
        For Each objItemRC In objReqCompra.colItens
            'Se for o item procurado
            If objItemRC.lNumIntDoc = lItemReqCompra Then
                'Devolve o item encontrado
                Set objItemReqCompra = objItemRC
                'Sai a funcao
                Exit Sub
            End If
        Next
    Next

    Exit Sub

End Sub

Function Carrega_Dados_Pedidos(objConcorrencia As ClassConcorrencia, colPedidoCompras As Collection) As Long
'Carrega em colPedidoCompras os Pedidos de Compra gerados a partir de diferentes Fornecedores e FiliaisFornecedores

Dim lErro As Long, bAchou As Boolean
Dim iIndice As Integer, objItemPC As ClassItemPedCompra
Dim dTotalItens As Double, lNumIntOriginal As Long
Dim objFornecedor As New ClassFornecedor
Dim objItemCotacao As ClassItemCotacao
Dim objCotItemConc As ClassCotacaoItemConc
Dim colItensCotacao As New Collection
Dim objQuantSupl As New ClassQuantSuplementar
Dim objPedidoCompra As ClassPedidoCompras
Dim objPedidoCotacao As New ClassPedidoCotacao
Dim colPedCompraGeral As New Collection
Dim colPedCompraExclu As New Collection
Dim objItemConcorrencia As ClassItemConcorrencia
Dim dQuantSupl As Double
Dim colCotItemConcAux As Collection
Dim colProdutos As New Collection

On Error GoTo Erro_Carrega_Dados_Pedidos
        
    Call Inicializa_QuantAssocia_ItenRC(gobjGeracaoPedCompraReq.colRequisicao)
    
    'Para cada item da concorr�ncia
    For Each objItemConcorrencia In objConcorrencia.colItens
        
        If objItemConcorrencia.lFornecedor > 0 And objItemConcorrencia.iFilial > 0 Then
            Set colPedidoCompras = colPedCompraExclu
        Else
            Set colPedidoCompras = colPedCompraGeral
        End If
        
        Call Transfere_Dados_Cotacoes(objItemConcorrencia.colCotacaoItemConc, colCotItemConcAux)
        
        'Para cada destino do item de concorrencia
        For Each objQuantSupl In objItemConcorrencia.colQuantSuplementar
            
            dQuantSupl = objQuantSupl.dQuantidade
        
            For Each objCotItemConc In colCotItemConcAux
                            
                If (objCotItemConc.iEscolhido = MARCADO) And (objCotItemConc.dQuantidadeComprar > 0) Then
                                        
                    'L� o Fornecedor
                    objFornecedor.sNomeReduzido = objCotItemConc.sFornecedor
                    
                    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                    If lErro <> SUCESSO And lErro <> 6681 Then gError 63799
        
                    'Se n�o encontrou ==> erro
                    If lErro = 6681 Then gError 63800
                    
                    iIndice = 0
                    bAchou = False
                      
                    'Verifica se j� foi criado pedido de compra com
                    'o fornecedor, a Filial e a condPagto da cota��o
                    For Each objPedidoCompra In colPedidoCompras
                        iIndice = iIndice + 1
                        
                        If objPedidoCompra.lFornecedor = objFornecedor.lCodigo And _
                           objPedidoCompra.iFilial = Codigo_Extrai(objCotItemConc.sFilial) And _
                           objPedidoCompra.iCondicaoPagto = Codigo_Extrai(objCotItemConc.sCondPagto) And _
                           objPedidoCompra.iTipoDestino = objQuantSupl.iTipoDestino And _
                           objPedidoCompra.lFornCliDestino = objQuantSupl.lFornCliDestino And _
                           objPedidoCompra.iFilialDestino = objQuantSupl.iFilialDestino Then
                           
                            bAchou = True
                            Exit For
                        End If
                    Next
                    
                    'Se j� existe pedido
                    If bAchou Then
                        'seleciona o pedido
                        Set objPedidoCompra = colPedidoCompras(iIndice)
                    'Sen�o
                    Else
                        'Cria um novo Pedido de compras com as caracter�sticas na cota��o
                        Set objPedidoCompra = New ClassPedidoCompras
                        
                        'Guarda o n�mero do pedido de cota��o do item de cota��o
                        objPedidoCompra.lPedCotacao = objCotItemConc.lPedCotacao
                        
                        objPedidoCompra.iFilialEmpresa = giFilialEmpresa
                        objPedidoCompra.dtData = gdtDataAtual
                        objPedidoCompra.dtDataAlteracao = DATA_NULA
                        objPedidoCompra.dtDataBaixa = DATA_NULA
                        objPedidoCompra.dtDataEmissao = DATA_NULA
                        objPedidoCompra.dtDataEnvio = DATA_NULA
                        objPedidoCompra.dValorProdutos = 0
                        objPedidoCompra.dValorTotal = 0
                        objPedidoCompra.iComprador = objConcorrencia.iComprador
                        objPedidoCompra.iCondicaoPagto = Codigo_Extrai(objCotItemConc.sCondPagto)
                        objPedidoCompra.iFilial = Codigo_Extrai(objCotItemConc.sFilial)
                        objPedidoCompra.iFilialDestino = objQuantSupl.iFilialDestino
                        objPedidoCompra.iTipoDestino = objQuantSupl.iTipoDestino
                        objPedidoCompra.lFornCliDestino = objQuantSupl.lFornCliDestino
                        objPedidoCompra.lFornecedor = objFornecedor.lCodigo
                        objPedidoCompra.sTipoFrete = TIPO_FOB
                        objPedidoCompra.iMoeda = objCotItemConc.iMoeda
                        objPedidoCompra.dTaxa = objCotItemConc.dTaxa
                        
                        colPedidoCompras.Add objPedidoCompra
                    End If
              
                    'cria um novo item para o pedido de compras
                    Set objItemPC = New ClassItemPedCompra
                          
                    'Se o pedido de cota��o utilizado no pedido n�o for o mesmo
                    If objPedidoCompra.lPedCotacao <> objCotItemConc.lPedCotacao Then objPedidoCompra.lPedCotacao = 0
          
                    objItemPC.dPrecoUnitario = objCotItemConc.dPrecoAjustado
                    objItemPC.dtDataLimite = objItemConcorrencia.dtDataNecessidade
                    objItemPC.iStatus = ITEM_PED_COMPRAS_ABERTO
                    objItemPC.iTipoOrigem = TIPO_ORIGEM_COTACAOITEMCONC
                    objItemPC.sDescProduto = objItemConcorrencia.sDescricao
                    objItemPC.sProduto = objItemConcorrencia.sProduto
                    objItemPC.sUM = objCotItemConc.sUMCompra
                    objItemPC.lNumIntOrigem = objCotItemConc.lNumIntDoc
                    
                    If dQuantSupl <= objCotItemConc.dQuantidadeComprar Then
                        objItemPC.dQuantidade = dQuantSupl
                        objCotItemConc.dQuantidadeComprar = objCotItemConc.dQuantidadeComprar - dQuantSupl
                        dQuantSupl = 0
                    Else
                        objItemPC.dQuantidade = objCotItemConc.dQuantidadeComprar
                        dQuantSupl = dQuantSupl - objCotItemConc.dQuantidadeComprar
                        objCotItemConc.dQuantidadeComprar = 0
                    End If
                    
                    objPedidoCompra.colItens.Add objItemPC
                    
                    'Vincula qt a comprar de ItensRC do mesmo destino do PC ao ItemPC gerado
                    lErro = Inclui_Quant_ItemReqCompra(objItemPC, objItemConcorrencia, objQuantSupl, gobjGeracaoPedCompraReq.colRequisicao, colProdutos)
                    If lErro <> SUCESSO Then gError 86150
                    
                    'Adiciona o item de cota��o na cole��o de itens de cotacao
                    lErro = colItensCotacao_Adiciona(objCotItemConc.lItemCotacao, colItensCotacao)
                    If lErro <> SUCESSO Then gError 62726
                End If
                If dQuantSupl = 0 Then Exit For
            Next
        Next
    Next
        
    Set colPedidoCompras = New Collection

    'Gera uma �nica colecao de Pedidos de Compra, a partir das colecoes colPedCompraExclu e colPedCompraGeral j� criadas
    lErro = PedidoCompra_Define_Colecao(colPedCompraExclu, colPedCompraGeral, colPedidoCompras)
    If lErro <> SUCESSO Then gError 76246
    
    'Aproveita os valores das cota��es utilizadas
    'caso o pedido tenha sido gerado com itens da mesma cota��o
    lErro = Atualiza_Valores_Pedido(colPedidoCompras, colItensCotacao)
    If lErro <> SUCESSO Then gError 62727
        
    Carrega_Dados_Pedidos = SUCESSO

    Exit Function

Erro_Carrega_Dados_Pedidos:

    Carrega_Dados_Pedidos = gErr

    Select Case gErr

        Case 63799, 70484, 62726, 62727, 86150
            'Erros tratados nas rotinas chamadas

        Case 63800, 70485, 76246
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INEXISTENTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161267)

    End Select

    Exit Function

End Function

Private Sub Transfere_Dados_Cotacoes(colCotacaoItemConc As Collection, colCotItemConcAux As Collection)

Dim objCotItemConc As ClassCotacaoItemConc
Dim objCotItemConcAux As ClassCotacaoItemConc

    Set colCotItemConcAux = New Collection
    
    For Each objCotItemConc In colCotacaoItemConc
        
        If objCotItemConc.iEscolhido = MARCADO Then
        
            Set objCotItemConcAux = New ClassCotacaoItemConc
            
            objCotItemConcAux.dAliquotaICMS = objCotItemConc.dAliquotaICMS
            objCotItemConcAux.dAliquotaIPI = objCotItemConc.dAliquotaIPI
            objCotItemConcAux.dCreditoICMS = objCotItemConc.dCreditoICMS
            objCotItemConcAux.dCreditoIPI = objCotItemConc.dCreditoIPI
            objCotItemConcAux.dPrecoAjustado = objCotItemConc.dPrecoAjustado
            objCotItemConcAux.dPrecoUnitario = objCotItemConc.dPrecoUnitario
            objCotItemConcAux.dPreferencia = objCotItemConc.dPreferencia
            objCotItemConcAux.dQuantCotada = objCotItemConc.dQuantCotada
            objCotItemConcAux.dQuantEntrega = objCotItemConc.dQuantEntrega
            objCotItemConcAux.dQuantidadeComprar = objCotItemConc.dQuantidadeComprar
            objCotItemConcAux.dtDataEntrega = objCotItemConc.dtDataEntrega
            objCotItemConcAux.dtDataValidade = objCotItemConc.dtDataValidade
            objCotItemConcAux.dValorPresente = objCotItemConc.dValorPresente
            objCotItemConcAux.iEscolhido = objCotItemConc.iEscolhido
            objCotItemConcAux.iPrazoEntrega = objCotItemConc.iPrazoEntrega
            objCotItemConcAux.iSelecionada = objCotItemConc.iSelecionada
            objCotItemConcAux.lItemCotacao = objCotItemConc.lItemCotacao
            objCotItemConcAux.lNumIntDoc = objCotItemConc.lNumIntDoc
            objCotItemConcAux.lPedCotacao = objCotItemConc.lPedCotacao
            objCotItemConcAux.sCondPagto = objCotItemConc.sCondPagto
            objCotItemConcAux.sFilial = objCotItemConc.sFilial
            objCotItemConcAux.sFornecedor = objCotItemConc.sFornecedor
            objCotItemConcAux.sMotivoEscolha = objCotItemConc.sMotivoEscolha
            objCotItemConcAux.sUMCompra = objCotItemConc.sUMCompra
            objCotItemConcAux.iMoeda = objCotItemConc.iMoeda
            objCotItemConcAux.dTaxa = objCotItemConc.dTaxa
            
            colCotItemConcAux.Add objCotItemConcAux
        End If
    Next

    Exit Sub
    
End Sub

Function Atualiza_Valores_Pedido(colPedidoCompras As Collection, colItensCotacao As Collection) As Long
'Aproveita os valores das cota��es utilizadas
'caso o pedido tenha sido gerado com itens da mesma cota��o
         
Dim lErro As Long
Dim objItemPC As ClassItemPedCompra
Dim objItemCotacao As ClassItemCotacao
Dim objCotItemConc As ClassCotacaoItemConc
Dim objPedidoCompra As ClassPedidoCompras
Dim objPedidoCotacao As New ClassPedidoCotacao
Dim objItemConcorrencia As ClassItemConcorrencia
    
On Error GoTo Erro_Atualiza_Valores_Pedido

    'Atualiza o valor dos produtos no pedido de venda
    For Each objPedidoCompra In colPedidoCompras

        'Zera os acumuladores dos valores
        objPedidoCompra.dValorDesconto = 0
        objPedidoCompra.dValorFrete = 0
        objPedidoCompra.dValorIPI = 0
        objPedidoCompra.dValorProdutos = 0
        objPedidoCompra.dValorSeguro = 0
        objPedidoCompra.dOutrasDespesas = 0

        'Se o pedido foi gerado com itens de um s� ped Cota��o
        If objPedidoCompra.lPedCotacao <> 0 Then

            objPedidoCotacao.lCodigo = objPedidoCompra.lPedCotacao
            objPedidoCotacao.iFilialEmpresa = giFilialEmpresa
            
            'L� o Pedido de Cotacao
            lErro = CF("PedidoCotacao_Le", objPedidoCotacao)
            If lErro <> SUCESSO And lErro <> 53670 Then gError 62728
            If lErro <> SUCESSO Then gError 62729 'N�o encontrou
            
            objPedidoCompra.sTipoFrete = objPedidoCotacao.iTipoFrete
            
            'Para cada item de pedido de compra
            For Each objItemPC In objPedidoCompra.colItens
                
                'Busca nos itens de concorrencia os dados do item de cota��o
                For Each objItemConcorrencia In gcolItemConcorrencia
                    
                    For Each objCotItemConc In objItemConcorrencia.colCotacaoItemConc
                        
                        'Se a cota��o foi a utilizada pelo item de Pedido de Compras
                        If objItemPC.lNumIntOrigem = objCotItemConc.lNumIntDoc Then

                            'Guarda o n�mero do item de cota��o
                            Set objItemCotacao = colItensCotacao(CStr(objCotItemConc.lItemCotacao))
                                                 
                            objPedidoCompra.dOutrasDespesas = objPedidoCompra.dOutrasDespesas + (objItemCotacao.dOutrasDespesas * (objItemPC.dQuantidade * objItemPC.dPrecoUnitario) / (objItemCotacao.dValorTotal))
                            objPedidoCompra.dValorDesconto = objPedidoCompra.dValorDesconto + (objItemCotacao.dValorDesconto * (objItemPC.dQuantidade * objItemPC.dPrecoUnitario) / (objItemCotacao.dValorTotal))
                            objPedidoCompra.dValorFrete = objPedidoCompra.dValorFrete + (objItemCotacao.dValorFrete * (objItemPC.dQuantidade * objItemPC.dPrecoUnitario) / (objItemCotacao.dValorTotal))
                            objPedidoCompra.dValorSeguro = objPedidoCompra.dValorSeguro + (objItemCotacao.dValorSeguro * (objItemPC.dQuantidade * objItemPC.dPrecoUnitario) / (objItemCotacao.dValorTotal))
                            objItemPC.dAliquotaICMS = objItemCotacao.dAliquotaICMS
                            objItemPC.dAliquotaIPI = objItemCotacao.dAliquotaIPI
                            objItemPC.dValorIPI = (objItemCotacao.dValorIPI * (objItemPC.dQuantidade * objItemPC.dPrecoUnitario) / (objItemCotacao.dValorTotal))
                            objPedidoCompra.dValorIPI = objPedidoCompra.dValorIPI + objItemPC.dValorIPI
                            objItemPC.lObservacao = objItemCotacao.lObservacao
                        End If
                    Next
                Next
            Next
        End If
        
        'Atualiza o valor dos produtos no Pedido de compras
        For Each objItemPC In objPedidoCompra.colItens
            objPedidoCompra.dValorProdutos = objPedidoCompra.dValorProdutos + (objItemPC.dPrecoUnitario * objItemPC.dQuantidade)
        Next
        
        objPedidoCompra.dValorTotal = objPedidoCompra.dValorFrete + objPedidoCompra.dValorIPI + objPedidoCompra.dValorProdutos + objPedidoCompra.dValorSeguro - objPedidoCompra.dValorDesconto
    Next
    
    Atualiza_Valores_Pedido = SUCESSO
    
    Exit Function
    
Erro_Atualiza_Valores_Pedido:

    Atualiza_Valores_Pedido = gErr
    
    Select Case gErr
    
        Case 62728
    
        Case 62729
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOTACAO_NAO_ENCONTRADO", gErr, objPedidoCotacao.lCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161268)
            
    End Select
    
    Exit Function

End Function

Private Sub Inicializa_QuantAssocia_ItenRC(colRequisicao As Collection)
'Zera o campo QuantNoPedido dos Itens de Requisi��o

Dim objReqCompra As ClassRequisicaoCompras
Dim objItemRC As ClassItemReqCompras

    For Each objReqCompra In colRequisicao
        For Each objItemRC In objReqCompra.colItens
            objItemRC.dQuantNoPedido = 0
        Next
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

Function Carrega_TipoTributacao() As Long
'Carrega Tipos de Tributa��o

Dim lErro As Long
Dim colTributacao As New AdmColCodigoNome
Dim iIndice As Integer

On Error GoTo Erro_Carrega_TipoTributacao

    'L� os Tipos de Tributa��o associadas a Compras
    lErro = CF("TiposTributacaoCompras_Le", colTributacao)
    If lErro <> SUCESSO Then gError 66123
           
    'Carrega Tipos de Tributa��o
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161269)
        
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

    'Para cada Requisi��o da tela
    For Each objReqCompra In gobjGeracaoPedCompraReq.colRequisicao
        'se for a req passada
        If objReqCompra.lCodigo = lReqCompra And objReqCompra.iFilialEmpresa = iFilialReq Then
            'Localiza o item procurado
            For Each objItemRC In objReqCompra.colItens
                If objItemRC.iItem = iItem Then
                    
                    objProduto.sCodigo = objItemRC.sProduto
                    'L� o produto
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

    Busca_QuantComprar_ItemReq = Err
    
    Select Case gErr
    
        Case 62796, 62798
        
        Case 62797
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161270)
            
    End Select

    Exit Function

End Function


Private Function Requisicoes_Atualiza() As Long
    
Dim objRequisicao As New ClassRequisicaoCompras
Dim objItemRC As ClassItemReqCompras
Dim lErro As Long
    
On Error GoTo Erro_Requisicoes_Atualiza
    
    'Se a Requisi��o foi selecionada
    If objGridRequisicoes.objGrid.Col = iGrid_EscolhidoReq_Col And objGridRequisicoes.iLinhasExistentes > 0 Then
               
        Set objRequisicao = gobjGeracaoPedCompraReq.colRequisicao(GridRequisicoes.Row)
        
        'Atualiza o campo selecionado na requisicao
        objRequisicao.iSelecionado = GridRequisicoes.TextMatrix(GridRequisicoes.Row, iGrid_EscolhidoReq_Col)
        
        'Para cada Item
        For Each objItemRC In objRequisicao.colItens
        
            If objRequisicao.iSelecionado = MARCADO Then
                If objItemRC.iSelecionado = DESMARCADO Then
                    objItemRC.iSelecionado = MARCADO
                    objItemRC.dQuantComprar = objItemRC.dQuantidade - objItemRC.dQuantCancelada - objItemRC.dQuantPedida - objItemRC.dQuantRecebida
                End If
            End If
            
            'Atualiza os dados do item de concorr�ncia vinculado ao ItemRC
            lErro = ItensConcorrencia_Atualiza(objRequisicao, objItemRC)
            If lErro <> SUCESSO Then gError 62750
        
        Next
        
        'Preenche o grid de itens de requisi��o
        lErro = GridItensReq_Preenche()
        If lErro <> SUCESSO Then gError 62751
        
        'Preenche o grid de produtos e cota��es
        lErro = Grids_Produto_Preenche()
        If lErro <> SUCESSO Then gError 62742

    End If
    
    Requisicoes_Atualiza = SUCESSO
    
    Exit Function
    
Erro_Requisicoes_Atualiza:

    Requisicoes_Atualiza = Err
    
    Select Case gErr
    
        Case 62742, 62750, 62751
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161271)
    
    End Select

    Exit Function

End Function

Function Atualiza_ItensReq() As Long

Dim iIndice1 As Integer
Dim iIndice2 As Integer
Dim objItemRC As ClassItemReqCompras
Dim objReqCompra As ClassRequisicaoCompras
Dim lErro As Long, bAchou As Boolean

On Error GoTo Erro_Atualiza_ItensReq

    'Busca o ItemRc e a Requisi��o correspondente a linha clicada
    For iIndice1 = 1 To gobjGeracaoPedCompraReq.colRequisicao.Count
        
        Set objReqCompra = gobjGeracaoPedCompraReq.colRequisicao(iIndice1)
        For iIndice2 = 1 To objReqCompra.colItens.Count
            
            Set objItemRC = objReqCompra.colItens(iIndice2)
            
            If objItemRC.iItem = StrParaInt(GridItensRequisicoes.TextMatrix(GridItensRequisicoes.Row, iGrid_Item_Col)) And _
               objReqCompra.lCodigo = StrParaLong(GridItensRequisicoes.TextMatrix(GridItensRequisicoes.Row, iGrid_CodigoReqItem_Col)) And _
               objReqCompra.iFilialEmpresa = Codigo_Extrai(GridItensRequisicoes.TextMatrix(GridItensRequisicoes.Row, iGrid_FilialReqItem_Col)) Then
                'Encontrou
                bAchou = True
                Exit For
            End If
        Next
        'Se j� achou sai
        If bAchou Then Exit For
    Next
    
    If objItemRC.iSelecionado = StrParaInt(GridItensRequisicoes.TextMatrix(GridItensRequisicoes.Row, iGrid_EscolhidoItem_Col)) Then Exit Function
    
    'Atualiza a sele��o do Item RC
    objItemRC.iSelecionado = StrParaInt(GridItensRequisicoes.TextMatrix(GridItensRequisicoes.Row, iGrid_EscolhidoItem_Col))

    'Atualiza os dados do item de concorr�ncia ligado ao item RC
    lErro = ItensConcorrencia_Atualiza(objReqCompra, objItemRC)
    If lErro <> SUCESSO Then gError 62743
    
    'Preenche o grid de produtos
    lErro = Grids_Produto_Preenche()
    If lErro <> SUCESSO Then gError 62742

    Atualiza_ItensReq = SUCESSO
    
    Exit Function

Erro_Atualiza_ItensReq:

    Atualiza_ItensReq = Err
    
    Select Case gErr

        Case 62742, 62743

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161272)

    End Select

    Exit Function

End Function


Function PedidoCompra_Define_Colecao(colPedCompraExclu As Collection, colPedCompraGeral As Collection, colPedidoCompras As Collection) As Long
'A partir das colecoes de Pedidos de Compra Exclusivos e de Pedidos de Compra N�o Exclusivos,
'define uma cole��o �nica para todos os Pedidos de Compra criados

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim bProdutoIgual As Boolean
Dim objPCGeral As New ClassPedidoCompras
Dim objPCExclu As New ClassPedidoCompras
Dim objItemPCExclu As New ClassItemPedCompra
Dim objItemPCGeral As New ClassItemPedCompra
Dim objPedidoCompra As New ClassPedidoCompras

On Error GoTo Erro_PedidoCompra_Define_Colecao

    'Verifica se existem Pedidos de Compra nas duas colecoes criadas
    If colPedCompraExclu.Count > 0 And colPedCompraGeral.Count > 0 Then
    
        bProdutoIgual = False
        For iIndice = colPedCompraExclu.Count To 1 Step -1
        
            Set objPCExclu = colPedCompraExclu.Item(iIndice)
            For Each objPCGeral In colPedCompraGeral
            
                'Verifica se os Pedidos tem o mesmo TipoDestino
                If objPCExclu.iTipoDestino = objPCGeral.iTipoDestino And objPCExclu.iFilialDestino = objPCGeral.iFilialDestino And objPCExclu.lFornCliDestino = objPCGeral.lFornCliDestino And objPCExclu.lFornecedor = objPCGeral.lFornecedor And objPCExclu.iFilial = objPCGeral.iFilial And objPCExclu.iCondicaoPagto = objPCGeral.iCondicaoPagto Then
                
                    For iIndice2 = objPCExclu.colItens.Count To 1 Step -1
                        
                        Set objItemPCExclu = objPCExclu.colItens.Item(iIndice2)
                        
                        For Each objItemPCGeral In objPCGeral.colItens
                        
                            'Verifica se o produto do Item Exclusivo est� presente na colecao de Itens nao exclusivos
                            If objItemPCExclu.sProduto = objItemPCGeral.sProduto Then
                                bProdutoIgual = True
                                Exit For
                            End If
                        Next
                    Next
                    'Se nao encontrou produto igual nas colecoes de Itens pesquisadas
                    If bProdutoIgual = False Then
                        
                        For iIndice2 = objPCExclu.colItens.Count To 1 Step -1
                            'Adiciona o item exclusivo na colecao de itens nao exclusivos
                            objPCGeral.colItens.Add objPCExclu.colItens.Item(iIndice2)
                            'Remove o Item
                            objPCExclu.colItens.Remove (iIndice2)
                        Next
                        
                        If objPCExclu.lPedCotacao <> objPCGeral.lPedCotacao Then objPCGeral.lPedCotacao = 0
                        
                        'Remove o Pedido
                        colPedCompraExclu.Remove (iIndice)
                        
                    End If
                End If
            Next
        Next
    End If
    
    'Coloca todos os pedidos em uma �nica cole��o
    For Each objPedidoCompra In colPedCompraExclu
        colPedidoCompras.Add objPedidoCompra
    Next
    For Each objPedidoCompra In colPedCompraGeral
        colPedidoCompras.Add objPedidoCompra
    Next
    
    PedidoCompra_Define_Colecao = SUCESSO
    
    Exit Function
    
Erro_PedidoCompra_Define_Colecao:

    PedidoCompra_Define_Colecao = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161273)
            
    End Select
    
    Exit Function
    
End Function

Public Function Gravar_Registro() As Long
'Grava a Concorrencia

Dim lErro As Long
Dim objConcorrencia As New ClassConcorrencia

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
       
    'Recolhe os dados da tela e armazena em objConcorrencia
    lErro = Move_Concorrencia_Memoria(objConcorrencia)
    If lErro <> SUCESSO Then gError 63761

    'Insere ou Altera uma concorrencia no BD
    lErro = CF("Concorrencia_Grava", objConcorrencia)
    If lErro <> SUCESSO Then gError 63672

    Call Rotina_Aviso(vbOKOnly, "AVISO_CONCORRENCIA_GRAVADA", objConcorrencia.lCodigo)

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function

Erro_Gravar_Registro:

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = gErr
    
    Select Case gErr

        Case 63756

        Case 63761, 63672
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161274)

    End Select

    Exit Function

End Function

Private Sub Busca_Produto(sProduto As String, colProdutos As Collection, objProduto As ClassProduto, bAchou As Boolean)

Dim objProdAux As ClassProduto

    bAchou = False
    
    For Each objProdAux In colProdutos
        
        If objProdAux.sCodigo = sProduto Then
            bAchou = True
            Set objProduto = objProdAux
            Exit For
        End If
    
    Next

    Exit Sub

End Sub

Function Carrega_Moeda() As Long

Dim lErro As Long
Dim objMoeda As ClassMoedas
Dim colMoedas As New Collection
Dim iPosMoedaReal As Integer
Dim iIndice As Integer

On Error GoTo Erro_Carrega_Moeda
    
    lErro = CF("Moedas_Le_Todas", colMoedas)
    If lErro <> SUCESSO Then gError 103371
    
    'se n�o existem moedas cadastradas
    If colMoedas.Count = 0 Then gError 103372
    
    For Each objMoeda In colMoedas
    
        Moeda.AddItem objMoeda.sNome
        Moeda.ItemData(iIndice) = objMoeda.iCodigo
        
        iIndice = iIndice + 1
    
    Next
    
    Moeda.ListIndex = -1

    Carrega_Moeda = SUCESSO
    
    Exit Function
    
Erro_Carrega_Moeda:

    Carrega_Moeda = gErr
    
    Select Case gErr
    
        Case 103371
        
        Case 103372
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDAS_NAO_CADASTRADAS", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161275)
    
    End Select

End Function

Private Sub Indica_Melhores()
'Indica as melhores opcoes

Dim dMenorPreco As Double
Dim objItemCotItemConc As ClassCotacaoItemConc
Dim objItemConcorrencia As New ClassItemConcorrencia
Dim objItemCotItemConcAux As ClassCotacaoItemConc

On Error GoTo Erro_Indica_Melhores

    Call Grid_Refresh_Checkbox_Limpa(objGridCotacoes)
    
    For Each objItemConcorrencia In gcolItemConcorrencia
        
        dMenorPreco = 0
        
        Set objItemCotItemConcAux = New ClassCotacaoItemConc
        
        'Para cada produto da colecao ...
         For Each objItemCotItemConc In objItemConcorrencia.colCotacaoItemConc
            
            'Se for para aparecer no grid ...
            If objItemCotItemConc.iSelecionada = MARCADO Then
            
                'Desmarca.
                objItemCotItemConc.iEscolhido = DESMARCADO
                
                'Caso ainda nao tenhamos um menor preco => Menor = $$ do Primeiro item
                If dMenorPreco = 0 Then
                    
                    dMenorPreco = objItemCotItemConc.dPrecoAjustado
                    
                    Set objItemCotItemConcAux = New ClassCotacaoItemConc
                    Set objItemCotItemConcAux = objItemCotItemConc
                    
                End If
                
                'Se o preco for menor do que o menor preco ja encontrado
                If objItemCotItemConc.dPrecoAjustado < dMenorPreco Then
                    
                    'Guarda o menor preco
                    dMenorPreco = objItemCotItemConc.dPrecoAjustado
                    
                    'Coloca o preco anterior como desmarcado
                    objItemCotItemConcAux.iEscolhido = DESMARCADO
                    
                    'Aponta para o novo candidato
                    Set objItemCotItemConcAux = New ClassCotacaoItemConc
                    Set objItemCotItemConcAux = objItemCotItemConc
                    
                End If
            
            End If
            
        Next
        
        'Seleciona o Menor
        objItemCotItemConcAux.iEscolhido = MARCADO
        
    Next
    
    Call Grid_Refresh_Checkbox(objGridCotacoes)

    Exit Sub

Erro_Indica_Melhores:
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161276)
    
    End Select

End Sub

Private Sub Categoria_Click()
'Preenche os itens da categoria selecionada

Dim lErro As Long
Dim colItensCategoria As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As ClassCategoriaProdutoItem

On Error GoTo Erro_Categoria_Click

    ItensCategoria.Clear
    
    If Len(Trim(Categoria.Text)) > 0 Then
    
        'Preenche o Obj
        objCategoriaProduto.sCategoria = Categoria.List(Categoria.ListIndex)
        
        'Le as categorias do Produto
        lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colItensCategoria)
        If lErro <> SUCESSO And lErro <> 22541 Then gError 108885
        
        For Each objCategoriaProdutoItem In colItensCategoria
            ItensCategoria.AddItem (objCategoriaProdutoItem.sItem)
        Next
        
    End If

    Exit Sub

Erro_Categoria_Click:

    Select Case gErr

         Case 108885
         
         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161277)

    End Select

End Sub

Private Function Carrega_Categorias() As Long

Dim lErro As Long
Dim objCategoria As New ClassCategoriaProduto
Dim colCategorias As New Collection

On Error GoTo Erro_Carrega_Categorias
    
    'Le a categoria
    lErro = CF("CategoriasProduto_Le_Todas", colCategorias)
    If lErro <> SUCESSO And lErro <> 22542 Then gError 108877
    
    '##############################################
    'COMENTADO POR WAGNER - N�O TEM QUE DAR ERRO QUANDO N�O
    'EXISTE CATEGORIA DE PRODUTO
    'Se nao encontrou => Erro
    'If lErro = 22542 Then gError 108878
    '##############################################
    
    Categoria.AddItem ("")
    
    'Carrega as combos de Categorias
    For Each objCategoria In colCategorias
    
        Categoria.AddItem objCategoria.sCategoria
        
    Next
    
    Carrega_Categorias = SUCESSO
    
    Exit Function
    
Erro_Carrega_Categorias:

    Carrega_Categorias = gErr
    
    Select Case gErr
    
        Case 108877
        
        Case 108878
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_NAO_CADASTRADA", gErr)
            '??? N�o existe categoria de produto cadastrada.

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161278)
    
    End Select

End Function

Private Sub ItensCategoria_Click()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub categoria_change()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
End Sub

'##############################################
'Inserido por Wagner
Private Sub Formata_Controles()

    PrecoUnitario.Format = gobjCOM.sFormatoPrecoUnitario
    PrecoUnitarioReal.Format = gobjCOM.sFormatoPrecoUnitario

End Sub
'##############################################

'#########################################################
'Inserido por Wagner
Sub Refaz_Grid(ByVal objGridInt As AdmGrid, ByVal iNumLinhas As Integer)
    objGridInt.objGrid.Rows = iNumLinhas + 1

    'Chama rotina de Inicializa��o do Grid
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
        If lErro <> SUCESSO And lErro <> 30401 Then gError 178876

        If lErro <> SUCESSO Then
        
            lErro = CF("ItensOP_Baixada_Le", objOrdemProducao)
            If lErro <> SUCESSO And lErro <> 178689 Then gError 178877
        
        End If
        
        If lErro = SUCESSO Then
        
            For Each objItemOP In objOrdemProducao.colItens
                
                If objItemOP.lCodPedido <> 0 Then
                    lCodigoPV = objItemOP.lCodPedido
                    Exit For
                End If
                
                If objItemOP.lNumIntDocPai <> 0 Then
                
                    lErro = CF("ItensOP_Le_PV", objItemOP.lNumIntDocPai, lCodigoPV, iFilialPV)
                    If lErro <> SUCESSO And lErro <> 178696 And lErro <> 178697 Then gError 178878
            
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
    
        Case 178876 To 178878
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178879)

    End Select

    Exit Function

End Function

