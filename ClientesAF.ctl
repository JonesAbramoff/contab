VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl Clientes 
   ClientHeight    =   6480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10320
   KeyPreview      =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   10320
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5805
      Index           =   5
      Left            =   60
      TabIndex        =   116
      Top             =   615
      Visible         =   0   'False
      Width           =   10215
      Begin VB.CheckBox IENaoContrib 
         Caption         =   "Não Contribuinte do ICMS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6555
         TabIndex        =   238
         Top             =   4380
         Value           =   1  'Checked
         Width           =   2760
      End
      Begin VB.ComboBox RegimeTributario 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1725
         TabIndex        =   236
         Top             =   4395
         Width           =   4395
      End
      Begin VB.CheckBox IEIsento 
         Caption         =   "Isento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   8700
         TabIndex        =   235
         Top             =   2745
         Value           =   1  'Checked
         Width           =   1530
      End
      Begin VB.Frame SSFrame7 
         Height          =   555
         Index           =   4
         Left            =   165
         TabIndex        =   120
         Top             =   75
         Width           =   8445
         Begin VB.Label ClienteLabel 
            Height          =   210
            Index           =   4
            Left            =   975
            TabIndex        =   122
            Top             =   225
            Width           =   7080
         End
         Begin VB.Label Label30 
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
            Height          =   210
            Index           =   4
            Left            =   210
            TabIndex        =   121
            Top             =   210
            Width           =   630
         End
      End
      Begin VB.ComboBox CondicaoPagto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6585
         TabIndex        =   71
         Top             =   1170
         Width           =   2100
      End
      Begin VB.ComboBox Mensagem 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1605
         TabIndex        =   72
         Top             =   1674
         Width           =   7080
      End
      Begin VB.ComboBox TabelaPreco 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6585
         TabIndex        =   73
         Top             =   2178
         Width           =   2100
      End
      Begin VB.CheckBox Bloqueado 
         Caption         =   "Cliente com crédito bloqueado"
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
         Left            =   2775
         TabIndex        =   69
         Top             =   690
         Width           =   2970
      End
      Begin VB.Frame Frame7 
         Caption         =   "Categorias"
         Height          =   1896
         Left            =   90
         TabIndex        =   117
         Top             =   2100
         Width           =   4515
         Begin VB.ComboBox ComboCategoriaClienteItem 
            Height          =   315
            Left            =   2085
            TabIndex        =   119
            TabStop         =   0   'False
            Top             =   405
            Width           =   1635
         End
         Begin VB.ComboBox ComboCategoriaCliente 
            Height          =   315
            Left            =   570
            TabIndex        =   118
            TabStop         =   0   'False
            Top             =   405
            Width           =   1545
         End
         Begin MSFlexGridLib.MSFlexGrid GridCategoria 
            Height          =   1560
            Left            =   225
            TabIndex        =   77
            Top             =   270
            Width           =   3780
            _ExtentX        =   6668
            _ExtentY        =   2752
            _Version        =   393216
            Rows            =   6
            Cols            =   3
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
         End
      End
      Begin MSMask.MaskEdBox LimiteCredito 
         Height          =   315
         Left            =   1605
         TabIndex        =   70
         Top             =   1170
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desconto 
         Height          =   315
         Left            =   1605
         TabIndex        =   68
         Top             =   675
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox InscricaoEstadual 
         Height          =   315
         Left            =   6585
         TabIndex        =   74
         Top             =   2682
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   18
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox InscricaoMunicipal 
         Height          =   315
         Left            =   6585
         TabIndex        =   75
         Top             =   3186
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   18
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox InscricaoSuframa 
         Height          =   315
         Left            =   6585
         TabIndex        =   76
         Top             =   3690
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   12
         Mask            =   "##.####-##-#"
         PromptChar      =   " "
      End
      Begin VB.Label Label0 
         AutoSize        =   -1  'True
         Caption         =   "Regime Tributário:"
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
         Left            =   75
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   237
         Top             =   4425
         Width           =   1560
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Municipal:"
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
         Left            =   4725
         TabIndex        =   131
         Top             =   3240
         Width           =   1725
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Estadual:"
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
         Left            =   4800
         TabIndex        =   130
         Top             =   2730
         Width           =   1650
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Suframa:"
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
         Left            =   4845
         TabIndex        =   129
         Top             =   3735
         Width           =   1605
      End
      Begin VB.Label Label6 
         Caption         =   "Limite de Crédito:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   15
         TabIndex        =   127
         Top             =   1215
         Width           =   1530
      End
      Begin VB.Label CondicaoPagtoLabel 
         Caption         =   "Condição de Pagamento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4290
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   126
         Top             =   1215
         Width           =   2160
      End
      Begin VB.Label Label8 
         Caption         =   "Desconto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   585
         TabIndex        =   125
         Top             =   735
         Width           =   885
      End
      Begin VB.Label MensagemNFLabel 
         Caption         =   "Mensagem NF:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   124
         Top             =   1740
         Width           =   1410
      End
      Begin VB.Label Label10 
         Caption         =   "Tabela de Preços:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4845
         TabIndex        =   123
         Top             =   2220
         Width           =   1590
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4140
      Index           =   4
      Left            =   660
      TabIndex        =   54
      Top             =   1260
      Visible         =   0   'False
      Width           =   8850
      Begin VB.ComboBox UsuRespCallCenter 
         Height          =   315
         Left            =   3915
         Sorted          =   -1  'True
         TabIndex        =   153
         Top             =   1095
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.ComboBox ComboCobrador 
         Height          =   315
         Left            =   2040
         TabIndex        =   66
         Top             =   3525
         Width           =   2385
      End
      Begin VB.TextBox Guia 
         Height          =   300
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   65
         Top             =   3015
         Width           =   1290
      End
      Begin VB.Frame Frame3 
         Caption         =   "Redespacho"
         Height          =   1155
         Left            =   4470
         TabIndex        =   106
         Top             =   2925
         Width           =   4260
         Begin VB.CheckBox RedespachoCli 
            Caption         =   "por conta do cliente"
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
            Left            =   240
            TabIndex        =   78
            Top             =   750
            Width           =   2100
         End
         Begin VB.ComboBox TranspRedespacho 
            Height          =   315
            Left            =   1665
            TabIndex        =   67
            Top             =   285
            Width           =   2475
         End
         Begin VB.Label TranspRedLabel 
            AutoSize        =   -1  'True
            Caption         =   "Transportadora:"
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
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   107
            Top             =   345
            Width           =   1365
         End
      End
      Begin VB.ComboBox TipoFrete 
         Height          =   315
         ItemData        =   "ClientesAF.ctx":0000
         Left            =   6255
         List            =   "ClientesAF.ctx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   2055
         Width           =   1125
      End
      Begin VB.ComboBox PadraoCobranca 
         Height          =   315
         ItemData        =   "ClientesAF.ctx":0018
         Left            =   2040
         List            =   "ClientesAF.ctx":001A
         TabIndex        =   59
         Top             =   1590
         Width           =   1965
      End
      Begin VB.ComboBox Transportadora 
         Height          =   315
         Left            =   6255
         TabIndex        =   64
         Top             =   2520
         Width           =   2475
      End
      Begin VB.ComboBox Cobrador 
         Height          =   315
         Left            =   6255
         TabIndex        =   58
         Top             =   1125
         Width           =   2475
      End
      Begin VB.ComboBox Regiao 
         Height          =   315
         Left            =   2040
         TabIndex        =   63
         Top             =   2520
         Width           =   2385
      End
      Begin MSMask.MaskEdBox ContaContabil 
         Height          =   315
         Left            =   6255
         TabIndex        =   56
         Top             =   660
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ComissaoVendas 
         Height          =   315
         Left            =   2040
         TabIndex        =   57
         Top             =   1125
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox FreqVisitas 
         Height          =   315
         Left            =   6270
         TabIndex        =   60
         Top             =   1590
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataUltVisita 
         Height          =   315
         Left            =   2040
         TabIndex        =   61
         Top             =   2055
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Vendedor 
         Height          =   315
         Left            =   2040
         TabIndex        =   55
         Top             =   660
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin VB.Frame SSFrame7 
         Height          =   555
         Index           =   1
         Left            =   240
         TabIndex        =   81
         Top             =   30
         Width           =   8445
         Begin VB.Label Label30 
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
            Height          =   210
            Index           =   1
            Left            =   210
            TabIndex        =   89
            Top             =   210
            Width           =   630
         End
         Begin VB.Label ClienteLabel 
            Height          =   210
            Index           =   0
            Left            =   960
            TabIndex        =   90
            Top             =   210
            Width           =   7080
         End
      End
      Begin VB.Label Label74 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Re. Call Center:"
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
         Left            =   2475
         TabIndex        =   154
         Top             =   1125
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Usuário Cobrador:"
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
         Left            =   435
         TabIndex        =   109
         Top             =   3585
         Width           =   1545
      End
      Begin VB.Label Label45 
         Caption         =   "Guia:"
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
         Left            =   1500
         TabIndex        =   108
         Top             =   3045
         Width           =   555
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Frete:"
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
         Left            =   4950
         TabIndex        =   105
         Top             =   2115
         Width           =   1215
      End
      Begin VB.Label PadraoCobrancaLabel 
         AutoSize        =   -1  'True
         Caption         =   "Padrão de Cobrança:"
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
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   91
         Top             =   1650
         Width           =   1815
      End
      Begin VB.Label TransportadoraLabel 
         AutoSize        =   -1  'True
         Caption         =   "Transportadora:"
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
         Left            =   4800
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   92
         Top             =   2580
         Width           =   1365
      End
      Begin VB.Label VendedorLabel 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor:"
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
         Left            =   1080
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   93
         Top             =   720
         Width           =   885
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "Comissão:"
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
         Left            =   1095
         TabIndex        =   94
         Top             =   1185
         Width           =   870
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "Região:"
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
         Left            =   1290
         TabIndex        =   95
         Top             =   2565
         Width           =   675
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "Frequência de Visitas:"
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
         Left            =   4260
         TabIndex        =   96
         Top             =   1650
         Width           =   1905
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "dias"
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
         Left            =   6750
         TabIndex        =   97
         Top             =   1650
         Width           =   360
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "Última Visita:"
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
         Left            =   855
         TabIndex        =   98
         Top             =   2115
         Width           =   1125
      End
      Begin VB.Label ContaContabilLabel 
         AutoSize        =   -1  'True
         Caption         =   "Conta Contábil:"
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
         Left            =   4830
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   99
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label AgenteCobradorLabel 
         AutoSize        =   -1  'True
         Caption         =   "Cobrador:"
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
         Left            =   5310
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   100
         Top             =   1200
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4140
      Index           =   3
      Left            =   930
      TabIndex        =   50
      Top             =   1020
      Visible         =   0   'False
      Width           =   8850
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3555
         Index           =   1
         Left            =   30
         TabIndex        =   170
         Top             =   -375
         Width           =   8595
         Begin TelasFATAF.TabEnderecoAF TabEnd 
            Height          =   1170
            Index           =   1
            Left            =   0
            TabIndex        =   175
            Top             =   0
            Width           =   10275
            _ExtentX        =   18124
            _ExtentY        =   2064
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3555
         Index           =   0
         Left            =   945
         TabIndex        =   84
         Top             =   1380
         Width           =   8595
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   3555
            Index           =   2
            Left            =   -1035
            TabIndex        =   171
            Top             =   -1050
            Width           =   8595
            Begin TelasFATAF.TabEnderecoAF TabEnd 
               Height          =   1170
               Index           =   2
               Left            =   0
               TabIndex        =   176
               Top             =   0
               Width           =   10275
               _ExtentX        =   18124
               _ExtentY        =   2064
            End
         End
      End
      Begin VB.Frame SSFrame5 
         Caption         =   "Endereços"
         Height          =   510
         Left            =   240
         TabIndex        =   82
         Top             =   585
         Width           =   8445
         Begin VB.OptionButton OpcaoEndereco 
            Caption         =   "Principal"
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
            Index           =   0
            Left            =   1440
            TabIndex        =   51
            Top             =   195
            Value           =   -1  'True
            Width           =   1200
         End
         Begin VB.OptionButton OpcaoEndereco 
            Caption         =   "Entrega"
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
            Index           =   1
            Left            =   3915
            TabIndex        =   52
            Top             =   180
            Width           =   1185
         End
         Begin VB.OptionButton OpcaoEndereco 
            Caption         =   "Cobrança"
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
            Index           =   2
            Left            =   6465
            TabIndex        =   53
            Top             =   180
            Width           =   1350
         End
      End
      Begin VB.Frame SSFrame7 
         Height          =   555
         Index           =   2
         Left            =   240
         TabIndex        =   83
         Top             =   30
         Width           =   8445
         Begin VB.Label Label30 
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
            Height          =   210
            Index           =   2
            Left            =   210
            TabIndex        =   101
            Top             =   210
            Width           =   630
         End
         Begin VB.Label ClienteLabel 
            Height          =   210
            Index           =   2
            Left            =   960
            TabIndex        =   102
            Top             =   210
            Width           =   7080
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4140
      Index           =   2
      Left            =   1185
      TabIndex        =   49
      Top             =   1170
      Visible         =   0   'False
      Width           =   8850
      Begin VB.Frame SSFrame7 
         Height          =   480
         Index           =   3
         Left            =   8805
         TabIndex        =   85
         Top             =   -165
         Width           =   8445
         Begin VB.Label Label30 
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
            Height          =   210
            Index           =   3
            Left            =   210
            TabIndex        =   103
            Top             =   165
            Width           =   630
         End
         Begin VB.Label ClienteLabel 
            Height          =   210
            Index           =   3
            Left            =   960
            TabIndex        =   104
            Top             =   165
            Width           =   7080
         End
      End
      Begin MSMask.MaskEdBox NomeReduzido 
         Height          =   315
         Left            =   1605
         TabIndex        =   145
         Top             =   165
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label LabelNomeRed 
         AutoSize        =   -1  'True
         Caption         =   "Nome Reduzido:"
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
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   146
         Top             =   210
         Width           =   1410
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6030
      Index           =   1
      Left            =   75
      TabIndex        =   48
      Top             =   570
      Width           =   10350
      Begin TelasFATAF.TabEnderecoAF TabEnd 
         Height          =   1170
         Index           =   0
         Left            =   -15
         TabIndex        =   11
         Top             =   1245
         Width           =   10275
         _ExtentX        =   18124
         _ExtentY        =   2064
      End
      Begin VB.ComboBox Tipo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5040
         TabIndex        =   3
         Top             =   75
         Width           =   1500
      End
      Begin VB.TextBox ContrSoc 
         Height          =   315
         Left            =   4950
         MaxLength       =   35
         TabIndex        =   35
         Top             =   4395
         Width           =   2670
      End
      Begin VB.TextBox NoBenef 
         Height          =   315
         Left            =   8775
         MaxLength       =   20
         TabIndex        =   32
         Top             =   4035
         Width           =   1410
      End
      Begin VB.TextBox LocalTrabalho 
         Height          =   315
         Left            =   975
         MaxLength       =   80
         TabIndex        =   23
         Top             =   3255
         Width           =   5385
      End
      Begin VB.ComboBox Sexo 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "ClientesAF.ctx":001C
         Left            =   990
         List            =   "ClientesAF.ctx":0026
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2490
         Width           =   765
      End
      Begin VB.TextBox Cargo 
         Height          =   315
         Left            =   975
         MaxLength       =   50
         TabIndex        =   25
         Top             =   3645
         Width           =   2025
      End
      Begin VB.ComboBox TipoApos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   975
         TabIndex        =   29
         Top             =   4035
         Width           =   2040
      End
      Begin VB.CheckBox Benemerito 
         Caption         =   "Benemérito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2250
         TabIndex        =   38
         Top             =   5580
         Width           =   1455
      End
      Begin VB.CheckBox Fundador 
         Caption         =   "Fundador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5055
         TabIndex        =   39
         Top             =   5580
         Width           =   1275
      End
      Begin VB.ComboBox Empresa1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "ClientesAF.ctx":0030
         Left            =   2955
         List            =   "ClientesAF.ctx":0032
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   75
         Width           =   1605
      End
      Begin VB.TextBox Observacao2 
         Height          =   315
         Left            =   975
         MaxLength       =   150
         TabIndex        =   37
         Top             =   5175
         Width           =   9180
      End
      Begin VB.ComboBox Empresa2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   495
         Width           =   1305
      End
      Begin VB.TextBox RazaoSocial 
         Height          =   315
         Left            =   990
         TabIndex        =   12
         Top             =   900
         Width           =   9225
      End
      Begin VB.CheckBox Ativo 
         Caption         =   "Ativo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   9435
         TabIndex        =   6
         Top             =   135
         Width           =   795
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   1635
         Picture         =   "ClientesAF.ctx":0034
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Numeração Automática"
         Top             =   90
         Width           =   300
      End
      Begin VB.TextBox Observacao 
         Height          =   315
         Left            =   975
         MaxLength       =   130
         TabIndex        =   36
         Top             =   4785
         Width           =   9180
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   990
         TabIndex        =   0
         Top             =   75
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MatriculaSec 
         Height          =   315
         Left            =   4035
         TabIndex        =   8
         Top             =   495
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MatriculaAF 
         Height          =   315
         Left            =   6570
         TabIndex        =   9
         Top             =   495
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MatriculaFRG 
         Height          =   315
         Left            =   9195
         TabIndex        =   10
         Top             =   495
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MatriculaPrinc 
         Height          =   315
         Left            =   7995
         TabIndex        =   4
         Top             =   75
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MatriculaPrincDV 
         Height          =   315
         Left            =   9000
         TabIndex        =   5
         Top             =   75
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CGC 
         Height          =   315
         Left            =   975
         TabIndex        =   18
         Top             =   2880
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   14
         Mask            =   "##############"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox RG 
         Height          =   315
         Left            =   2250
         TabIndex        =   14
         Top             =   2490
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox RGOrgaoEmissor 
         Height          =   315
         Left            =   4920
         TabIndex        =   15
         Top             =   2490
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownEmissaoRG 
         Height          =   300
         Left            =   9915
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2490
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataEmissaoRG 
         Height          =   315
         Left            =   8760
         TabIndex        =   16
         Top             =   2490
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownNasc 
         Height          =   300
         Left            =   6075
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2880
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataNasc 
         Height          =   315
         Left            =   4920
         TabIndex        =   19
         Top             =   2880
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownFalecimento 
         Height          =   300
         Left            =   2130
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   4410
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFalecimento 
         Height          =   315
         Left            =   975
         TabIndex        =   33
         Top             =   4410
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataInsc 
         Height          =   315
         Left            =   8760
         TabIndex        =   22
         Top             =   2880
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataAdmAF 
         Height          =   315
         Left            =   8760
         TabIndex        =   24
         Top             =   3255
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownAposINSS 
         Height          =   300
         Left            =   6105
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   3645
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataAposINSS 
         Height          =   315
         Left            =   4950
         TabIndex        =   26
         Top             =   3645
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataAposFRG 
         Height          =   315
         Left            =   8775
         TabIndex        =   28
         Top             =   3645
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownConBenef 
         Height          =   300
         Left            =   6105
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   4035
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataConBenef 
         Height          =   315
         Left            =   4950
         TabIndex        =   30
         Top             =   4035
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownUltAtu 
         Height          =   300
         Left            =   9900
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   5550
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataUltAtu 
         Height          =   315
         Left            =   8745
         TabIndex        =   40
         Top             =   5550
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownInsc 
         Height          =   300
         Left            =   9900
         TabIndex        =   172
         TabStop         =   0   'False
         Top             =   2880
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDownAdmAF 
         Height          =   300
         Left            =   9915
         TabIndex        =   173
         TabStop         =   0   'False
         Top             =   3255
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDownAposFRG 
         Height          =   300
         Left            =   9915
         TabIndex        =   174
         TabStop         =   0   'False
         Top             =   3645
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label TipoClienteLabel 
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
         Left            =   4575
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   152
         Top             =   135
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Últ. Atualização.:"
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
         Index           =   25
         Left            =   7215
         TabIndex        =   151
         Top             =   5610
         Width           =   1485
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Contr.Soc.:"
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
         Index           =   16
         Left            =   3930
         TabIndex        =   150
         Top             =   4455
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Benef.:"
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
         Index           =   15
         Left            =   8070
         TabIndex        =   149
         Top             =   4080
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dt Con. Benef:"
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
         Index           =   26
         Left            =   3630
         TabIndex        =   148
         Top             =   4110
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Local de Trabalho:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   12
         Left            =   75
         TabIndex        =   147
         Top             =   3210
         Width           =   915
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "CPF:"
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
         TabIndex        =   144
         Top             =   2940
         Width           =   420
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "RG:"
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
         Left            =   1875
         TabIndex        =   143
         Top             =   2535
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Orgão:"
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
         Index           =   23
         Left            =   4290
         TabIndex        =   142
         Top             =   2535
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dt. Expedição:"
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
         Index           =   24
         Left            =   7425
         TabIndex        =   141
         Top             =   2550
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sexo:"
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
         Index           =   11
         Left            =   405
         TabIndex        =   140
         Top             =   2535
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Nasc.:"
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
         Index           =   18
         Left            =   3840
         TabIndex        =   139
         Top             =   2940
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Falecim.:"
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
         Index           =   17
         Left            =   105
         TabIndex        =   138
         Top             =   4500
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Inscr.:"
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
         Index           =   19
         Left            =   7680
         TabIndex        =   21
         Top             =   2940
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cargo:"
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
         Index           =   13
         Left            =   315
         TabIndex        =   137
         Top             =   3675
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dt Admissão Furnas:"
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
         Index           =   20
         Left            =   6945
         TabIndex        =   136
         Top             =   3300
         Width           =   1740
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dt Apos.INSS:"
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
         Index           =   22
         Left            =   3660
         TabIndex        =   135
         Top             =   3720
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Apo.:"
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
         Index           =   14
         Left            =   -30
         TabIndex        =   134
         Top             =   4065
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dt Apos.FRG:"
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
         Index           =   21
         Left            =   7515
         TabIndex        =   133
         Top             =   3690
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "-"
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
         Height          =   210
         Index           =   8
         Left            =   9030
         TabIndex        =   132
         Top             =   195
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "OBS 2:"
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
         Index           =   10
         Left            =   270
         TabIndex        =   128
         Top             =   5250
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Matrícula FRG:"
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
         Index           =   1
         Left            =   7785
         TabIndex        =   115
         Top             =   570
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Matrícula AF:"
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
         Index           =   2
         Left            =   5355
         TabIndex        =   114
         Top             =   540
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Matrícula Sec:"
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
         Index           =   3
         Left            =   2670
         TabIndex        =   113
         Top             =   555
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Empresa 2:"
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
         Index           =   4
         Left            =   -15
         TabIndex        =   112
         Top             =   555
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Matrícula Princ:"
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
         Index           =   5
         Left            =   6570
         TabIndex        =   111
         Top             =   135
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Empresa 1:"
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
         Index           =   6
         Left            =   1980
         TabIndex        =   110
         Top             =   135
         Width           =   960
      End
      Begin VB.Label LabelCliente 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
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
         Left            =   285
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   86
         Top             =   135
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
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
         Left            =   375
         TabIndex        =   87
         Top             =   945
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "OBS 1:"
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
         Index           =   7
         Left            =   270
         TabIndex        =   88
         Top             =   4830
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4395
      Index           =   6
      Left            =   750
      TabIndex        =   177
      Top             =   1455
      Visible         =   0   'False
      Width           =   8970
      Begin VB.Frame SSFrame7 
         Height          =   555
         Index           =   0
         Left            =   15
         TabIndex        =   232
         Top             =   -90
         Width           =   8880
         Begin VB.Label Label30 
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
            Height          =   210
            Index           =   0
            Left            =   210
            TabIndex        =   234
            Top             =   210
            Width           =   630
         End
         Begin VB.Label ClienteLabel 
            Height          =   210
            Index           =   1
            Left            =   960
            TabIndex        =   233
            Top             =   210
            Width           =   7080
         End
      End
      Begin VB.Frame SSFrame3 
         Caption         =   "Atraso"
         Height          =   720
         Left            =   30
         TabIndex        =   223
         Top             =   3105
         Width           =   6690
         Begin VB.Label MaiorAtraso 
            Caption         =   "0"
            Height          =   210
            Left            =   825
            TabIndex        =   231
            Top             =   435
            Width           =   720
         End
         Begin VB.Label MediaAtraso 
            Caption         =   "0"
            Height          =   210
            Left            =   825
            TabIndex        =   230
            Top             =   195
            Width           =   720
         End
         Begin VB.Label ValorPagtosAtraso 
            Caption         =   "0,00"
            Height          =   210
            Left            =   4530
            TabIndex        =   229
            Top             =   435
            Width           =   1395
         End
         Begin VB.Label SaldoAtrasados 
            Caption         =   "0,00"
            Height          =   210
            Left            =   4530
            TabIndex        =   228
            Top             =   195
            Width           =   1395
         End
         Begin VB.Label Label27 
            Caption         =   "Maior:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   225
            TabIndex        =   227
            Top             =   435
            Width           =   570
         End
         Begin VB.Label Label26 
            Caption         =   "Média:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   195
            TabIndex        =   226
            Top             =   195
            Width           =   585
         End
         Begin VB.Label Label25 
            Caption         =   "Valor Pagamentos com Atraso:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1860
            TabIndex        =   225
            Top             =   435
            Width           =   2610
         End
         Begin VB.Label Label24 
            Caption         =   "Saldo de Atrasados:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2745
            TabIndex        =   224
            Top             =   195
            Width           =   1725
         End
      End
      Begin VB.Frame SSFrame2 
         Caption         =   "Cheques Devolvidos"
         Height          =   720
         Left            =   6765
         TabIndex        =   218
         Top             =   3105
         Width           =   2115
         Begin VB.Label DataUltChequeDevolvido 
            Caption         =   "  /  /    "
            Height          =   210
            Left            =   765
            TabIndex        =   222
            Top             =   450
            Width           =   1170
         End
         Begin VB.Label NumChequesDevolvidos 
            Caption         =   "0"
            Height          =   210
            Left            =   945
            TabIndex        =   221
            Top             =   180
            Width           =   405
         End
         Begin VB.Label Label29 
            Caption         =   "Número:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   135
            TabIndex        =   220
            Top             =   180
            Width           =   750
         End
         Begin VB.Label Label28 
            Caption         =   "Último:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   135
            TabIndex        =   219
            Top             =   450
            Width           =   600
         End
      End
      Begin VB.Frame SSFrame1 
         Caption         =   "Compras"
         Height          =   1500
         Left            =   4050
         TabIndex        =   207
         Top             =   570
         Width           =   4830
         Begin VB.Label DataUltimaCompra 
            Caption         =   "  /  /    "
            Height          =   210
            Left            =   3255
            TabIndex        =   217
            Top             =   735
            Width           =   1170
         End
         Begin VB.Label DataPrimeiraCompra 
            Caption         =   "  /  /    "
            Height          =   210
            Left            =   3255
            TabIndex        =   216
            Top             =   300
            Width           =   1170
         End
         Begin VB.Label ValorAcumuladoCompras 
            Caption         =   "0,00"
            Height          =   210
            Left            =   1830
            TabIndex        =   215
            Top             =   1170
            Width           =   1575
         End
         Begin VB.Label MediaCompra 
            Caption         =   "0,00"
            Height          =   210
            Left            =   900
            TabIndex        =   214
            Top             =   735
            Width           =   1410
         End
         Begin VB.Label NumeroCompras 
            Caption         =   "0"
            Height          =   210
            Left            =   1050
            TabIndex        =   213
            Top             =   300
            Width           =   585
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
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
            TabIndex        =   212
            Top             =   315
            Width           =   720
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Média:"
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
            TabIndex        =   211
            Top             =   750
            Width           =   585
         End
         Begin VB.Label Label15 
            Caption         =   "Última:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2565
            TabIndex        =   210
            Top             =   735
            Width           =   615
         End
         Begin VB.Label Label16 
            Caption         =   "Primeira:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2415
            TabIndex        =   209
            Top             =   300
            Width           =   765
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Valor Acumulado:"
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
            TabIndex        =   208
            Top             =   1185
            Width           =   1500
         End
      End
      Begin VB.Frame SSFrame4 
         Caption         =   "Saldos"
         Height          =   1515
         Left            =   30
         TabIndex        =   196
         Top             =   570
         Width           =   3705
         Begin VB.Label SaldoTitulos 
            Caption         =   "0,00"
            Height          =   210
            Left            =   1860
            TabIndex        =   201
            Top             =   300
            Width           =   1575
         End
         Begin VB.Label SaldoPedidosLiberados 
            Caption         =   "0,00"
            Height          =   210
            Left            =   1860
            TabIndex        =   199
            Top             =   585
            Width           =   1575
         End
         Begin VB.Label SaldodeCredito 
            Caption         =   "0,00"
            Height          =   210
            Left            =   1860
            TabIndex        =   198
            Top             =   1215
            Width           =   1575
         End
         Begin VB.Label SaldoLimitedeCredito 
            Caption         =   "0,00"
            Height          =   210
            Left            =   1860
            TabIndex        =   197
            Top             =   900
            Width           =   1575
         End
         Begin VB.Label Label42 
            Caption         =   "Saldo de Crédito:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   315
            TabIndex        =   206
            Top             =   1215
            Width           =   1650
         End
         Begin VB.Label Label14 
            Caption         =   "Limite de Crédito:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   285
            TabIndex        =   205
            Top             =   900
            Width           =   1650
         End
         Begin VB.Label Label19 
            Caption         =   "Em Duplicatas:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   495
            TabIndex        =   204
            Top             =   -30
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.Label Label20 
            Caption         =   "Em Títulos:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   780
            TabIndex        =   203
            Top             =   300
            Width           =   1020
         End
         Begin VB.Label Label37 
            Caption         =   "Pedidos Liberados:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   165
            TabIndex        =   202
            Top             =   585
            Width           =   1650
         End
         Begin VB.Label SaldoDuplicatas 
            Caption         =   "0,00"
            Height          =   210
            Left            =   1860
            TabIndex        =   200
            Top             =   -30
            Visible         =   0   'False
            Width           =   1575
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Total em Títulos"
         Height          =   930
         Left            =   30
         TabIndex        =   183
         Top             =   2100
         Width           =   8865
         Begin VB.Label TotalCRComProtesto 
            Caption         =   "0,00"
            Height          =   210
            Left            =   7140
            TabIndex        =   185
            Top             =   270
            Width           =   1575
         End
         Begin VB.Label PercCRComProtesto 
            Caption         =   "0%"
            Height          =   210
            Left            =   7140
            TabIndex        =   184
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label TotalCREmCartorio 
            Caption         =   "0,00"
            Height          =   210
            Left            =   4335
            TabIndex        =   189
            Top             =   285
            Width           =   1575
         End
         Begin VB.Label PercCREmCartorio 
            Caption         =   "0%"
            Height          =   210
            Left            =   4335
            TabIndex        =   188
            Top             =   615
            Width           =   1575
         End
         Begin VB.Label Label9 
            Caption         =   "% do Total:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3300
            TabIndex        =   191
            Top             =   615
            Width           =   1155
         End
         Begin VB.Label Label12 
            Caption         =   "Em Cartório:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3225
            TabIndex        =   190
            Top             =   285
            Width           =   1200
         End
         Begin VB.Label Label43 
            Caption         =   "% do Total:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   6090
            TabIndex        =   187
            Top             =   600
            Width           =   1140
         End
         Begin VB.Label Label46 
            Caption         =   "Com Protesto:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   5880
            TabIndex        =   186
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label PercCREmAberto 
            Caption         =   "0%"
            Height          =   210
            Left            =   1860
            TabIndex        =   195
            Top             =   630
            Width           =   1575
         End
         Begin VB.Label TotalCR 
            Caption         =   "0,00"
            Height          =   210
            Left            =   1860
            TabIndex        =   194
            Top             =   300
            Width           =   1575
         End
         Begin VB.Label Label40 
            Caption         =   "Valor:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1290
            TabIndex        =   193
            Top             =   300
            Width           =   510
         End
         Begin VB.Label Label41 
            Caption         =   "% Aberto:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   975
            TabIndex        =   192
            Top             =   630
            Width           =   840
         End
      End
      Begin VB.CommandButton BotaoTitRec 
         Caption         =   "Todos os Títulos"
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
         Left            =   30
         TabIndex        =   182
         Top             =   3870
         Width           =   1725
      End
      Begin VB.CommandButton BotaoTitRecPgAtrasado 
         Caption         =   "Títulos Pagos com Atraso"
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
         Left            =   1800
         TabIndex        =   181
         Top             =   3870
         Width           =   1725
      End
      Begin VB.CommandButton BotaoTitRecEmCart 
         Caption         =   "Títulos em Cartório"
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
         Left            =   3585
         TabIndex        =   180
         Top             =   3870
         Width           =   1725
      End
      Begin VB.CommandButton BotaoTitRecComProt 
         Caption         =   "Títulos com Protesto"
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
         Left            =   5370
         TabIndex        =   179
         Top             =   3870
         Width           =   1725
      End
      Begin VB.CommandButton BotaoTitRecVenc 
         Caption         =   "Títulos a Receber Vencidos"
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
         Left            =   7140
         TabIndex        =   178
         Top             =   3870
         Width           =   1725
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4185
      Index           =   0
      Left            =   1200
      TabIndex        =   155
      Top             =   1665
      Visible         =   0   'False
      Width           =   8850
      Begin VB.Frame SSFrame7 
         Height          =   555
         Index           =   5
         Left            =   240
         TabIndex        =   167
         Top             =   30
         Width           =   8445
         Begin VB.Label Label30 
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
            Height          =   210
            Index           =   5
            Left            =   210
            TabIndex        =   169
            Top             =   210
            Width           =   630
         End
         Begin VB.Label ClienteLabel 
            Height          =   210
            Index           =   5
            Left            =   960
            TabIndex        =   168
            Top             =   210
            Width           =   7080
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Faturamento"
         Height          =   2385
         Left            =   255
         TabIndex        =   156
         Top             =   765
         Width           =   8430
         Begin VB.Frame Frame8 
            Caption         =   "Faixa de faturamento"
            Height          =   1125
            Left            =   600
            TabIndex        =   162
            Top             =   1035
            Width           =   3525
            Begin MSMask.MaskEdBox PercentMaisReceb 
               Height          =   315
               Left            =   2310
               TabIndex        =   163
               Top             =   300
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   556
               _Version        =   393216
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
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox PercentMenosReceb 
               Height          =   315
               Left            =   2310
               TabIndex        =   164
               Top             =   720
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   556
               _Version        =   393216
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
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin VB.Label Label56 
               AutoSize        =   -1  'True
               Caption         =   "Percentagem a menos:"
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
               Left            =   270
               TabIndex        =   166
               Top             =   780
               Width           =   1950
            End
            Begin VB.Label Label57 
               AutoSize        =   -1  'True
               Caption         =   "Percentagem a mais:"
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
               TabIndex        =   165
               Top             =   375
               Width           =   1785
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Faturamento fora da faixa"
            Height          =   1125
            Left            =   4290
            TabIndex        =   159
            Top             =   1050
            Width           =   3540
            Begin VB.OptionButton RecebForaFaixa 
               Caption         =   "Não aceita"
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
               Height          =   285
               Index           =   0
               Left            =   330
               TabIndex        =   161
               Top             =   300
               Value           =   -1  'True
               Width           =   2415
            End
            Begin VB.OptionButton RecebForaFaixa 
               Caption         =   "Avisa e aceita"
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
               Height          =   285
               Index           =   1
               Left            =   315
               TabIndex        =   160
               Top             =   660
               Width           =   2655
            End
         End
         Begin VB.CheckBox NaoTemFaixaReceb 
            Caption         =   "Aceita qualquer quantidade sem aviso"
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
            Left            =   615
            TabIndex        =   158
            Top             =   720
            Width           =   3585
         End
         Begin VB.CheckBox IgnoraRecebPadrao 
            Caption         =   "Ignora configuração padrão"
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
            Left            =   615
            TabIndex        =   157
            Top             =   345
            Width           =   3585
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   8130
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   30
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "ClientesAF.ctx":011E
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "ClientesAF.ctx":0278
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1110
         Picture         =   "ClientesAF.ctx":0402
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1620
         Picture         =   "ClientesAF.ctx":0934
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoContatos 
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
      Left            =   8160
      Picture         =   "ClientesAF.ctx":0AB2
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   1935
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Filiais 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   520
      Left            =   10755
      Picture         =   "ClientesAF.ctx":2B68
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   60
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   6555
      Left            =   135
      TabIndex        =   80
      Top             =   150
      Visible         =   0   'False
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   11562
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Inscrições"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Endereços"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Vendas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Outros"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Estatísticas"
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
Attribute VB_Name = "Clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Pendencias: diminuir tamanho do Form_Load

Option Explicit

Event Unload()

Private WithEvents objCT As CTClientes
Attribute objCT.VB_VarHelpID = -1

Private Sub Ativo_Click()
     Call objCT.Ativo_Click
End Sub

Private Sub IENaoContrib_Click()
    Call objCT.IENaoContrib_Click
End Sub

Private Sub LabelCliente_Click()
     Call objCT.Label1_Click
End Sub

Private Sub LabelNomeRed_Click()
     Call objCT.Label3_Click
End Sub

Private Sub RedespachoCli_Click()
    Call objCT.RedespachoCli_Click
End Sub

Private Sub TipoFrete_Change()
    Call objCT.TipoFrete_Change
End Sub

Private Sub BotaoProxNum_Click()
     Call objCT.BotaoProxNum_Click
End Sub

Private Sub AgenteCobradorLabel_Click()
     Call objCT.AgenteCobradorLabel_Click
End Sub

Private Sub BotaoExcluir_Click()
     Call objCT.BotaoExcluir_Click
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub CGC_GotFocus()
     Call objCT.CGC_GotFocus
End Sub

Private Sub RG_GotFocus()
     Call objCT.RG_GotFocus
End Sub

Private Sub Cobrador_Click()
     Call objCT.Cobrador_Click
End Sub

Private Sub Cobrador_Validate(Cancel As Boolean)
     Call objCT.Cobrador_Validate(Cancel)
End Sub

Private Sub Codigo_GotFocus()
     Call objCT.Codigo_GotFocus
End Sub

Private Sub Codigo_Validate(Cancel As Boolean)
     Call objCT.Codigo_Validate(Cancel)
End Sub

Private Sub ComboCategoriaCliente_Change()
     Call objCT.ComboCategoriaCliente_Change
End Sub

Private Sub ComboCategoriaCliente_Click()
     Call objCT.ComboCategoriaCliente_Click
End Sub

Private Sub ComboCategoriaCliente_GotFocus()
     Call objCT.ComboCategoriaCliente_GotFocus
End Sub

Private Sub ComboCategoriaCliente_KeyPress(KeyAscii As Integer)
     Call objCT.ComboCategoriaCliente_KeyPress(KeyAscii)
End Sub

Private Sub ComboCategoriaCliente_Validate(Cancel As Boolean)
     Call objCT.ComboCategoriaCliente_Validate(Cancel)
End Sub

Private Sub ComboCategoriaClienteItem_Change()
     Call objCT.ComboCategoriaClienteItem_Change
End Sub

Private Sub ComboCategoriaClienteItem_Click()
     Call objCT.ComboCategoriaClienteItem_Click
End Sub

Private Sub ComboCategoriaClienteItem_GotFocus()
     Call objCT.ComboCategoriaClienteItem_GotFocus
End Sub

Private Sub ComboCategoriaClienteItem_KeyPress(KeyAscii As Integer)
     Call objCT.ComboCategoriaClienteItem_KeyPress(KeyAscii)
End Sub

Private Sub ComboCategoriaClienteItem_Validate(Cancel As Boolean)
     Call objCT.ComboCategoriaClienteItem_Validate(Cancel)
End Sub

Private Sub CondicaoPagto_Click()
     Call objCT.CondicaoPagto_Click
End Sub

Private Sub CondicaoPagto_Validate(Cancel As Boolean)
     Call objCT.CondicaoPagto_Validate(Cancel)
End Sub

Private Sub CondicaoPagtoLabel_Click()
     Call objCT.CondicaoPagtoLabel_Click
End Sub

Private Sub ContaContabilLabel_Click()
     Call objCT.ContaContabilLabel_Click
End Sub

Private Sub DataUltVisita_GotFocus()
     Call objCT.DataUltVisita_GotFocus
End Sub

Private Sub Filiais_Click()
     Call objCT.Filiais_Click
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub BotaoGravar_Click()
     Call objCT.BotaoGravar_Click
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Private Sub CGC_Change()
     Call objCT.CGC_Change
End Sub

Private Sub CGC_Validate(Cancel As Boolean)
     Call objCT.CGC_Validate(Cancel)
End Sub

Private Sub RG_Change()
     Call objCT.RG_Change
End Sub

Private Sub Cobrador_Change()
     Call objCT.Cobrador_Change
End Sub

Private Sub Codigo_Change()
     Call objCT.Codigo_Change
End Sub

Private Sub ComissaoVendas_Change()
     Call objCT.ComissaoVendas_Change
End Sub

Private Sub ComissaoVendas_Validate(Cancel As Boolean)
     Call objCT.ComissaoVendas_Validate(Cancel)
End Sub

Private Sub CondicaoPagto_Change()
     Call objCT.CondicaoPagto_Change
End Sub

Private Sub ContaContabil_Change()
     Call objCT.ContaContabil_Change
End Sub

Private Sub ContaContabil_Validate(Cancel As Boolean)
     Call objCT.ContaContabil_Validate(Cancel)
End Sub

Private Sub DataUltVisita_Change()
     Call objCT.DataUltVisita_Change
End Sub

Private Sub DataUltVisita_Validate(Cancel As Boolean)
     Call objCT.DataUltVisita_Validate(Cancel)
End Sub

Private Sub Desconto_Change()
     Call objCT.Desconto_Change
End Sub

Private Sub Desconto_Validate(Cancel As Boolean)
     Call objCT.Desconto_Validate(Cancel)
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Private Sub FreqVisitas_Change()
     Call objCT.FreqVisitas_Change
End Sub

Private Sub FreqVisitas_GotFocus()
     Call objCT.FreqVisitas_GotFocus
End Sub

Private Sub GridCategoria_Click()
     Call objCT.GridCategoria_Click
End Sub

Private Sub GridCategoria_EnterCell()
     Call objCT.GridCategoria_EnterCell
End Sub

Private Sub GridCategoria_GotFocus()
     Call objCT.GridCategoria_GotFocus
End Sub

Private Sub GridCategoria_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridCategoria_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridCategoria_KeyPress(KeyAscii As Integer)
     Call objCT.GridCategoria_KeyPress(KeyAscii)
End Sub

Private Sub GridCategoria_LeaveCell()
     Call objCT.GridCategoria_LeaveCell
End Sub

Private Sub GridCategoria_Validate(Cancel As Boolean)
     Call objCT.GridCategoria_Validate(Cancel)
End Sub

Private Sub GridCategoria_RowColChange()
     Call objCT.GridCategoria_RowColChange
End Sub

Private Sub GridCategoria_Scroll()
     Call objCT.GridCategoria_Scroll
End Sub

Private Sub InscricaoEstadual_Change()
     Call objCT.InscricaoEstadual_Change
End Sub

Private Sub InscricaoMunicipal_Change()
     Call objCT.InscricaoMunicipal_Change
End Sub

Private Sub Inscricaosuframa_Change()
     Call objCT.Inscricaosuframa_Change
End Sub

Private Sub LimiteCredito_Change()
     Call objCT.LimiteCredito_Change
End Sub

Private Sub LimiteCredito_Validate(Cancel As Boolean)
     Call objCT.LimiteCredito_Validate(Cancel)
End Sub

Private Sub Mensagem_Change()
     Call objCT.Mensagem_Change
End Sub

Private Sub Mensagem_Click()
     Call objCT.Mensagem_Click
End Sub

Private Sub Mensagem_Validate(Cancel As Boolean)
     Call objCT.Mensagem_Validate(Cancel)
End Sub

Private Sub MensagemNFLabel_Click()
     Call objCT.MensagemNFLabel_Click
End Sub

Private Sub NomeReduzido_Change()
     Call objCT.NomeReduzido_Change
End Sub

Private Sub NomeReduzido_Validate(Cancel As Boolean)
     Call objCT.NomeReduzido_Validate(Cancel)
End Sub

Private Sub Observacao_Change()
     Call objCT.Observacao_Change
End Sub

Private Sub Observacao2_Change()
     Call objCT.Observacao2_Change
End Sub

Private Sub Opcao_Click()
     Call objCT.Opcao_Click
End Sub

Private Sub OpcaoEndereco_Click(Index As Integer)
     Call objCT.OpcaoEndereco_Click(Index)
End Sub

Private Sub PadraoCobranca_Change()
     Call objCT.PadraoCobranca_Change
End Sub

Private Sub PadraoCobranca_Click()
     Call objCT.PadraoCobranca_Click
End Sub

Private Sub PadraoCobranca_Validate(Cancel As Boolean)
     Call objCT.PadraoCobranca_Validate(Cancel)
End Sub

Private Sub PadraoCobrancaLabel_Click()
     Call objCT.PadraoCobrancaLabel_Click
End Sub

Private Sub RazaoSocial_Change()
     Call objCT.RazaoSocial_Change
End Sub

Private Sub Regiao_Change()
     Call objCT.Regiao_Change
End Sub

Private Sub Regiao_Click()
     Call objCT.Regiao_Click
End Sub

Private Sub Regiao_Validate(Cancel As Boolean)
     Call objCT.Regiao_Validate(Cancel)
End Sub

Private Sub TabelaPreco_Change()
     Call objCT.TabelaPreco_Change
End Sub

Private Sub TabelaPreco_Click()
     Call objCT.TabelaPreco_Click
End Sub

Private Sub TabelaPreco_Validate(Cancel As Boolean)
     Call objCT.TabelaPreco_Validate(Cancel)
End Sub

Private Sub Tipo_Change()
     Call objCT.Tipo_Change
End Sub

Private Sub Tipo_Click()
     Call objCT.Tipo_Click
End Sub

Private Sub Tipo_Validate(Cancel As Boolean)
     Call objCT.Tipo_Validate(Cancel)
End Sub

Private Sub TipoClienteLabel_Click()
     Call objCT.TipoClienteLabel_Click
End Sub

Private Sub Transportadora_Change()
     Call objCT.Transportadora_Change
End Sub

Private Sub Transportadora_Click()
     Call objCT.Transportadora_Click
End Sub

Private Sub Transportadora_Validate(Cancel As Boolean)
     Call objCT.Transportadora_Validate(Cancel)
End Sub

Private Sub TransportadoraLabel_Click()
     Call objCT.TransportadoraLabel_Click
End Sub

Private Sub UserControl_Initialize()
    Set objCT = New CTClientes
    Set objCT.objUserControl = Me
    
    'AposFurnas
    Set objCT.gobjInfoUsu = New CTClientesVGAF
    Set objCT.gobjInfoUsu.gobjTelaUsu = New CTClientesAF

End Sub

Private Sub Vendedor_Change()
     Call objCT.Vendedor_Change
End Sub

Private Sub Vendedor_Validate(Cancel As Boolean)
     Call objCT.Vendedor_Validate(Cancel)
End Sub

Function Trata_Parametros(Optional objcliente As ClassCliente) As Long
     Trata_Parametros = objCT.Trata_Parametros(objcliente)
End Function

Private Sub VendedorLabel_Click()
     Call objCT.VendedorLabel_Click
End Sub

Public Sub Form_Activate()
     Call objCT.Form_Activate
End Sub

Public Sub Form_Deactivate()
     Call objCT.Form_Deactivate
End Sub

Public Function Form_Load_Ocx() As Object

    Call objCT.Form_Load_Ocx
    Set Form_Load_Ocx = Me

End Function

Public Sub Form_Unload(Cancel As Integer)
    If Not (objCT Is Nothing) Then
        Call objCT.Form_Unload(Cancel)
        Set objCT.objUserControl = Nothing
        Set objCT = Nothing
    End If
End Sub

Private Sub objCT_Unload()
   RaiseEvent Unload
End Sub

Public Function Name() As String
    Name = objCT.Name
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

Private Sub ClienteLabel_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(ClienteLabel(Index), Source, X, Y)
End Sub

Private Sub ClienteLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ClienteLabel(Index), Button, Shift, X, Y)
End Sub

Private Sub Label30_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label30(Index), Source, X, Y)
End Sub

Private Sub Label30_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label30(Index), Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub

Private Sub TipoClienteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoClienteLabel, Source, X, Y)
End Sub

Private Sub TipoClienteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoClienteLabel, Button, Shift, X, Y)
End Sub

Private Sub Label21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label21, Source, X, Y)
End Sub

Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label21, Button, Shift, X, Y)
End Sub

Private Sub Label16_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label16, Source, X, Y)
End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label16, Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label15, Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15, Button, Shift, X, Y)
End Sub

Private Sub Label17_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label17, Source, X, Y)
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label17, Button, Shift, X, Y)
End Sub

Private Sub Label18_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label18, Source, X, Y)
End Sub

Private Sub Label18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label18, Button, Shift, X, Y)
End Sub

Private Sub NumeroCompras_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumeroCompras, Source, X, Y)
End Sub

Private Sub NumeroCompras_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumeroCompras, Button, Shift, X, Y)
End Sub

Private Sub MediaCompra_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(MediaCompra, Source, X, Y)
End Sub

Private Sub MediaCompra_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(MediaCompra, Button, Shift, X, Y)
End Sub

Private Sub ValorAcumuladoCompras_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorAcumuladoCompras, Source, X, Y)
End Sub

Private Sub ValorAcumuladoCompras_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorAcumuladoCompras, Button, Shift, X, Y)
End Sub

Private Sub DataPrimeiraCompra_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataPrimeiraCompra, Source, X, Y)
End Sub

Private Sub DataPrimeiraCompra_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataPrimeiraCompra, Button, Shift, X, Y)
End Sub

Private Sub DataUltimaCompra_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataUltimaCompra, Source, X, Y)
End Sub

Private Sub DataUltimaCompra_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataUltimaCompra, Button, Shift, X, Y)
End Sub

Private Sub Label28_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label28, Source, X, Y)
End Sub

Private Sub Label28_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label28, Button, Shift, X, Y)
End Sub

Private Sub Label29_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label29, Source, X, Y)
End Sub

Private Sub Label29_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label29, Button, Shift, X, Y)
End Sub

Private Sub NumChequesDevolvidos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumChequesDevolvidos, Source, X, Y)
End Sub

Private Sub NumChequesDevolvidos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumChequesDevolvidos, Button, Shift, X, Y)
End Sub

Private Sub DataUltChequeDevolvido_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataUltChequeDevolvido, Source, X, Y)
End Sub

Private Sub DataUltChequeDevolvido_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataUltChequeDevolvido, Button, Shift, X, Y)
End Sub

Private Sub Label24_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label24, Source, X, Y)
End Sub

Private Sub Label24_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label24, Button, Shift, X, Y)
End Sub

Private Sub Label25_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label25, Source, X, Y)
End Sub

Private Sub Label25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label25, Button, Shift, X, Y)
End Sub

Private Sub Label26_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label26, Source, X, Y)
End Sub

Private Sub Label26_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label26, Button, Shift, X, Y)
End Sub

Private Sub Label27_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label27, Source, X, Y)
End Sub

Private Sub Label27_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label27, Button, Shift, X, Y)
End Sub

Private Sub SaldoAtrasados_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SaldoAtrasados, Source, X, Y)
End Sub

Private Sub SaldoAtrasados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SaldoAtrasados, Button, Shift, X, Y)
End Sub

Private Sub ValorPagtosAtraso_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorPagtosAtraso, Source, X, Y)
End Sub

Private Sub ValorPagtosAtraso_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorPagtosAtraso, Button, Shift, X, Y)
End Sub

Private Sub MediaAtraso_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(MediaAtraso, Source, X, Y)
End Sub

Private Sub MediaAtraso_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(MediaAtraso, Button, Shift, X, Y)
End Sub

Private Sub MaiorAtraso_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(MaiorAtraso, Source, X, Y)
End Sub

Private Sub MaiorAtraso_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(MaiorAtraso, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub Label20_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label20, Source, X, Y)
End Sub

Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label20, Button, Shift, X, Y)
End Sub

Private Sub Label19_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label19, Source, X, Y)
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label19, Button, Shift, X, Y)
End Sub

Private Sub SaldoTitulos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SaldoTitulos, Source, X, Y)
End Sub

Private Sub SaldoTitulos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SaldoTitulos, Button, Shift, X, Y)
End Sub

Private Sub SaldoDuplicatas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SaldoDuplicatas, Source, X, Y)
End Sub

Private Sub SaldoDuplicatas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SaldoDuplicatas, Button, Shift, X, Y)
End Sub

Private Sub SaldoPedidosLiberados_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SaldoPedidosLiberados, Source, X, Y)
End Sub

Private Sub SaldoPedidosLiberados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SaldoPedidosLiberados, Button, Shift, X, Y)
End Sub

Private Sub PadraoCobrancaLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PadraoCobrancaLabel, Source, X, Y)
End Sub

Private Sub PadraoCobrancaLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PadraoCobrancaLabel, Button, Shift, X, Y)
End Sub

Private Sub TransportadoraLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TransportadoraLabel, Source, X, Y)
End Sub

Private Sub TransportadoraLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TransportadoraLabel, Button, Shift, X, Y)
End Sub

Private Sub VendedorLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(VendedorLabel, Source, X, Y)
End Sub

Private Sub VendedorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(VendedorLabel, Button, Shift, X, Y)
End Sub

Private Sub Label44_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label44, Source, X, Y)
End Sub

Private Sub Label44_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label44, Button, Shift, X, Y)
End Sub

Private Sub Label33_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label33, Source, X, Y)
End Sub

Private Sub Label33_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label33, Button, Shift, X, Y)
End Sub

Private Sub Label47_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label47, Source, X, Y)
End Sub

Private Sub Label47_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label47, Button, Shift, X, Y)
End Sub

Private Sub Label48_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label48, Source, X, Y)
End Sub

Private Sub Label48_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label48, Button, Shift, X, Y)
End Sub

Private Sub Label49_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label49, Source, X, Y)
End Sub

Private Sub Label49_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label49, Button, Shift, X, Y)
End Sub

Private Sub ContaContabilLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContaContabilLabel, Source, X, Y)
End Sub

Private Sub ContaContabilLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContaContabilLabel, Button, Shift, X, Y)
End Sub

Private Sub AgenteCobradorLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(AgenteCobradorLabel, Source, X, Y)
End Sub

Private Sub AgenteCobradorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(AgenteCobradorLabel, Button, Shift, X, Y)
End Sub

Private Sub Label35_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label35, Source, X, Y)
End Sub

Private Sub Label35_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label35, Button, Shift, X, Y)
End Sub

Private Sub Label34_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label34, Source, X, Y)
End Sub

Private Sub Label34_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label34, Button, Shift, X, Y)
End Sub

Private Sub Label36_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label36, Source, X, Y)
End Sub

Private Sub Label36_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label36, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub MensagemNFLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(MensagemNFLabel, Source, X, Y)
End Sub

Private Sub MensagemNFLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(MensagemNFLabel, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub CondicaoPagtoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CondicaoPagtoLabel, Source, X, Y)
End Sub

Private Sub CondicaoPagtoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CondicaoPagtoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

Private Sub TranspRedespacho_Change()
     Call objCT.TranspRedespacho_Change
End Sub

Private Sub TranspRedespacho_Click()
     Call objCT.TranspRedespacho_Click
End Sub

Private Sub TranspRedespacho_Validate(Cancel As Boolean)
     Call objCT.TranspRedespacho_Validate(Cancel)
End Sub

Private Sub TranspRedLabel_Click()
     Call objCT.TranspRedLabel_Click
End Sub

Private Sub Guia_Change()
    Call objCT.Guia_Change
End Sub

'######################################
'Inserido por Wagner
Private Sub Bloqueado_Click()
     Call objCT.Bloqueado_Click
End Sub
'######################################

Private Sub BotaoContatos_Click()
     Call objCT.BotaoContatos_Click
End Sub

Private Sub ComboCobrador_Click()
    objCT.ComboCobrador_Click
End Sub

Private Sub ComboCobrador_Validate(Cancel As Boolean)
    objCT.ComboCobrador_Validate (Cancel)
End Sub

'PARTE CUSTOMIZADA

Private Sub Empresa1_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.Empresa1_Change(objCT)
End Sub

Private Sub Empresa1_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.Empresa1_Click(objCT)
End Sub

Private Sub MatriculaPrinc_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.MatriculaPrinc_GotFocus(objCT)
End Sub

Private Sub MatriculaPrinc_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.MatriculaPrinc_Change(objCT)
End Sub

Private Sub MatriculaPrinc_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.MatriculaPrinc_Validate(objCT, Cancel)
End Sub

Private Sub Empresa2_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.Empresa2_Change(objCT)
End Sub

Private Sub Empresa2_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.Empresa2_Click(objCT)
End Sub

Private Sub MatriculaSec_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.MatriculaSec_GotFocus(objCT)
End Sub

Private Sub MatriculaSec_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.MatriculaSec_Change(objCT)
End Sub

Private Sub MatriculaSec_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.MatriculaSec_Validate(objCT, Cancel)
End Sub

Private Sub MatriculaAF_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.MatriculaAF_GotFocus(objCT)
End Sub

Private Sub MatriculaAF_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.MatriculaAF_Change(objCT)
End Sub

Private Sub MatriculaAF_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.MatriculaAF_Validate(objCT, Cancel)
End Sub

Private Sub MatriculaFRG_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.MatriculaFRG_GotFocus(objCT)
End Sub

Private Sub MatriculaFRG_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.MatriculaFRG_Change(objCT)
End Sub

Private Sub MatriculaFRG_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.MatriculaFRG_Validate(objCT, Cancel)
End Sub

Private Sub Benemerito_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.Benemerito_Click(objCT)
End Sub

Private Sub Fundador_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.Fundador_Click(objCT)
End Sub

Private Sub RGOrgaoEmissor_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.RGOrgaoEmissor_Change(objCT)
End Sub

Private Sub DataEmissaoRG_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataEmissaoRG_GotFocus(objCT)
End Sub

Private Sub DataEmissaoRG_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataEmissaoRG_Change(objCT)
End Sub

Private Sub DataEmissaoRG_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataEmissaoRG_Validate(objCT, Cancel)
End Sub

Private Sub Sexo_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.Sexo_Change(objCT)
End Sub

Private Sub Sexo_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.Sexo_Click(objCT)
End Sub

Private Sub DataNasc_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataNasc_GotFocus(objCT)
End Sub

Private Sub DataNasc_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataNasc_Change(objCT)
End Sub

Private Sub DataNasc_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataNasc_Validate(objCT, Cancel)
End Sub

Private Sub DataFalecimento_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataFalecimento_GotFocus(objCT)
End Sub

Private Sub DataFalecimento_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataFalecimento_Change(objCT)
End Sub

Private Sub DataFalecimento_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataFalecimento_Validate(objCT, Cancel)
End Sub

Private Sub DataInsc_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataInsc_GotFocus(objCT)
End Sub

Private Sub DataInsc_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataInsc_Change(objCT)
End Sub

Private Sub DataInsc_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataInsc_Validate(objCT, Cancel)
End Sub

Private Sub DataAdmAF_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataAdmAF_GotFocus(objCT)
End Sub

Private Sub DataAdmAF_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataAdmAF_Change(objCT)
End Sub

Private Sub DataAdmAF_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataAdmAF_Validate(objCT, Cancel)
End Sub

Private Sub DataAposINSS_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataAposINSS_GotFocus(objCT)
End Sub

Private Sub DataAposINSS_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataAposINSS_Change(objCT)
End Sub

Private Sub DataAposINSS_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataAposINSS_Validate(objCT, Cancel)
End Sub

Private Sub DataAposFRG_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataAposFRG_GotFocus(objCT)
End Sub

Private Sub DataAposFRG_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataAposFRG_Change(objCT)
End Sub

Private Sub DataAposFRG_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataAposFRG_Validate(objCT, Cancel)
End Sub

Private Sub DataUltAtu_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataUltAtu_GotFocus(objCT)
End Sub

Private Sub DataUltAtu_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataUltAtu_Change(objCT)
End Sub

Private Sub DataUltAtu_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataUltAtu_Validate(objCT, Cancel)
End Sub

Private Sub DataConBenef_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataConBenef_GotFocus(objCT)
End Sub

Private Sub DataConBenef_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataConBenef_Change(objCT)
End Sub

Private Sub DataConBenef_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataConBenef_Validate(objCT, Cancel)
End Sub

Private Sub Cargo_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.Cargo_Change(objCT)
End Sub

Private Sub ContrSoc_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ContrSoc_Change(objCT)
End Sub

Private Sub NoBenef_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.NoBenef_Change(objCT)
End Sub

Private Sub LocalTrabalho_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.LocalTrabalho_Change(objCT)
End Sub

Private Sub TipoApos_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.TipoApos_Change(objCT)
End Sub

Private Sub TipoApos_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.TipoApos_Click(objCT)
End Sub

Private Sub UpDownEmissaoRG_DownClick()
     Call objCT.gobjInfoUsu.gobjTelaUsu.UpDownEmissaoRG_DownClick(objCT)
End Sub

Private Sub UpDownEmissaoRG_UPClick()
     Call objCT.gobjInfoUsu.gobjTelaUsu.UpDownEmissaoRG_UPClick(objCT)
End Sub

Private Sub UpDownNasc_DownClick()
     Call objCT.gobjInfoUsu.gobjTelaUsu.UpDownNasc_DownClick(objCT)
End Sub

Private Sub UpDownNasc_UPClick()
     Call objCT.gobjInfoUsu.gobjTelaUsu.UpDownNasc_UPClick(objCT)
End Sub

Private Sub UpDownFalecimento_DownClick()
     Call objCT.gobjInfoUsu.gobjTelaUsu.UpDownFalecimento_DownClick(objCT)
End Sub

Private Sub UpDownFalecimento_UPClick()
     Call objCT.gobjInfoUsu.gobjTelaUsu.UpDownFalecimento_UPClick(objCT)
End Sub

Private Sub UpDownInsc_DownClick()
     Call objCT.gobjInfoUsu.gobjTelaUsu.UpDownInsc_DownClick(objCT)
End Sub

Private Sub UpDownInsc_UPClick()
     Call objCT.gobjInfoUsu.gobjTelaUsu.UpDownInsc_UPClick(objCT)
End Sub

Private Sub UpDownAdmAF_DownClick()
     Call objCT.gobjInfoUsu.gobjTelaUsu.UpDownAdmAF_DownClick(objCT)
End Sub

Private Sub UpDownAdmAF_UPClick()
     Call objCT.gobjInfoUsu.gobjTelaUsu.UpDownAdmAF_UPClick(objCT)
End Sub

Private Sub UpDownAposINSS_DownClick()
     Call objCT.gobjInfoUsu.gobjTelaUsu.UpDownAposINSS_DownClick(objCT)
End Sub

Private Sub UpDownAposINSS_UPClick()
     Call objCT.gobjInfoUsu.gobjTelaUsu.UpDownAposINSS_UPClick(objCT)
End Sub

Private Sub UpDownAposFRG_DownClick()
     Call objCT.gobjInfoUsu.gobjTelaUsu.UpDownAposFRG_DownClick(objCT)
End Sub

Private Sub UpDownAposFRG_UPClick()
     Call objCT.gobjInfoUsu.gobjTelaUsu.UpDownAposFRG_UPClick(objCT)
End Sub

Private Sub UpDownUltAtu_DownClick()
     Call objCT.gobjInfoUsu.gobjTelaUsu.UpDownUltAtu_DownClick(objCT)
End Sub

Private Sub UpDownUltAtu_UPClick()
     Call objCT.gobjInfoUsu.gobjTelaUsu.UpDownUltAtu_UPClick(objCT)
End Sub

Private Sub UpDownConBenef_DownClick()
     Call objCT.gobjInfoUsu.gobjTelaUsu.UpDownConBenef_DownClick(objCT)
End Sub

Private Sub UpDownConBenef_UPClick()
     Call objCT.gobjInfoUsu.gobjTelaUsu.UpDownConBenef_UPClick(objCT)
End Sub

Private Sub MatriculaPrincDV_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.MatriculaPrincDV_GotFocus(objCT)
End Sub

Private Sub MatriculaPrincDV_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.MatriculaPrincDV_Change(objCT)
End Sub

Private Sub MatriculaPrincDV_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.MatriculaPrincDV_Validate(objCT, Cancel)
End Sub

Private Sub UsuRespCallCenter_Click()
    objCT.UsuRespCallCenter_Click
End Sub

Private Sub UsuRespCallCenter_Validate(Cancel As Boolean)
    objCT.UsuRespCallCenter_Validate (Cancel)
End Sub


Private Sub BotaoTitRec_Click()
    Call objCT.BotaoTitRec_Click
End Sub

Private Sub BotaoTitRecComProt_Click()
    Call objCT.BotaoTitRecComProt_Click
End Sub

Private Sub BotaoTitRecEmCart_Click()
    Call objCT.BotaoTitRecEmCart_Click
End Sub

Private Sub BotaoTitRecPgAtrasado_Click()
    Call objCT.BotaoTitRecPgAtrasado_Click
End Sub

Private Sub BotaoTitRecVenc_Click()
    Call objCT.BotaoTitRecVenc_Click
End Sub

Private Sub IEIsento_Click()
    Call objCT.IEIsento_Click
End Sub

Private Sub InscricaoEstadual_Validate(Cancel As Boolean)
    Call objCT.InscricaoEstadual_Validate(Cancel)
End Sub
