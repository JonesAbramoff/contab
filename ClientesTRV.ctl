VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl Clientes 
   ClientHeight    =   6180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6180
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5115
      Index           =   3
      Left            =   165
      TabIndex        =   66
      Top             =   570
      Visible         =   0   'False
      Width           =   9180
      Begin VB.Frame Frame24 
         Caption         =   "Vouchers Pago no Cartão com valor acima"
         Height          =   645
         Left            =   45
         TabIndex        =   253
         Top             =   3630
         Width           =   3600
         Begin MSMask.MaskEdBox PercFatorDevCMCC 
            Height          =   285
            Left            =   1125
            TabIndex        =   254
            Top             =   270
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Format          =   "##0.#0\%"
            PromptChar      =   " "
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Devolver"
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
            TabIndex        =   256
            Top             =   330
            Width           =   780
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   " junto a CMCC"
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
            Left            =   1890
            TabIndex        =   255
            Top             =   330
            Width           =   1230
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Agência"
         Height          =   3210
         Left            =   45
         TabIndex        =   152
         Top             =   420
         Width           =   3600
         Begin VB.Frame Frame16 
            Caption         =   "Exceções por produto"
            Height          =   2775
            Left            =   105
            TabIndex        =   154
            Top             =   405
            Width           =   3420
            Begin VB.CommandButton BotaoExcAgProduto 
               Caption         =   "Produtos"
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
               Left            =   60
               TabIndex        =   84
               Top             =   2475
               Width           =   1530
            End
            Begin MSMask.MaskEdBox ExcAgProduto 
               Height          =   210
               Left            =   225
               TabIndex        =   155
               Top             =   1005
               Width           =   1100
               _ExtentX        =   1931
               _ExtentY        =   370
               _Version        =   393216
               BorderStyle     =   0
               AllowPrompt     =   -1  'True
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ExcAgPercComis 
               Height          =   210
               Left            =   1575
               TabIndex        =   156
               Top             =   1005
               Width           =   1055
               _ExtentX        =   1852
               _ExtentY        =   370
               _Version        =   393216
               BorderStyle     =   0
               PromptInclude   =   0   'False
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
               Format          =   "0.00##"
               PromptChar      =   " "
            End
            Begin MSFlexGridLib.MSFlexGrid GridExcAg 
               Height          =   675
               Left            =   60
               TabIndex        =   83
               Top             =   195
               Width           =   3330
               _ExtentX        =   5874
               _ExtentY        =   1191
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
         Begin MSMask.MaskEdBox PercComiAg 
            Height          =   275
            Left            =   1140
            TabIndex        =   82
            Top             =   135
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   476
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin VB.Label Label71 
            AutoSize        =   -1  'True
            Caption         =   "% de comis.:"
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
            Left            =   60
            TabIndex        =   153
            Top             =   195
            Width           =   1080
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Promotor"
         Height          =   825
         Left            =   45
         TabIndex        =   121
         Top             =   4260
         Width           =   3600
         Begin MSMask.MaskEdBox Vendedor 
            Height          =   275
            Left            =   1125
            TabIndex        =   85
            Top             =   165
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   476
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox ComissaoVendas 
            Height          =   270
            Left            =   1125
            TabIndex        =   86
            Top             =   480
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   476
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "% de comis.:"
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
            Left            =   60
            TabIndex        =   123
            Top             =   510
            Width           =   1080
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
            Left            =   255
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   122
            Top             =   195
            Width           =   885
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Correntista"
         Height          =   2295
         Left            =   3705
         TabIndex        =   116
         Top             =   2790
         Width           =   5400
         Begin VB.Frame Frame15 
            Caption         =   "Exceções por produto"
            Height          =   1770
            Left            =   75
            TabIndex        =   148
            Top             =   465
            Width           =   5250
            Begin VB.CommandButton BotaoExcCor 
               Caption         =   "Correntistas"
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
               Left            =   3630
               TabIndex        =   96
               Top             =   1455
               Width           =   1530
            End
            Begin VB.CommandButton BotaoExcCorProduto 
               Caption         =   "Produtos"
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
               Left            =   90
               TabIndex        =   95
               Top             =   1455
               Width           =   1530
            End
            Begin MSMask.MaskEdBox ExcCor 
               Height          =   225
               Left            =   405
               TabIndex        =   149
               Top             =   510
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   397
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               AllowPrompt     =   -1  'True
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ExcCorProduto 
               Height          =   210
               Left            =   2070
               TabIndex        =   150
               Top             =   780
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   370
               _Version        =   393216
               BorderStyle     =   0
               AllowPrompt     =   -1  'True
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ExcCorPercComis 
               Height          =   210
               Left            =   2985
               TabIndex        =   151
               Top             =   525
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   370
               _Version        =   393216
               BorderStyle     =   0
               PromptInclude   =   0   'False
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
               Format          =   "0.00##"
               PromptChar      =   " "
            End
            Begin MSFlexGridLib.MSFlexGrid GridExcCor 
               Height          =   1245
               Left            =   60
               TabIndex        =   94
               Top             =   195
               Width           =   5130
               _ExtentX        =   9049
               _ExtentY        =   2196
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
         Begin MSMask.MaskEdBox Correntista 
            Height          =   315
            Left            =   1005
            TabIndex        =   92
            Top             =   165
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox PercComiCorr 
            Height          =   315
            Left            =   4095
            TabIndex        =   93
            Top             =   165
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin VB.Label LabelCorr 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   315
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   118
            Top             =   195
            Width           =   660
         End
         Begin VB.Label Label69 
            AutoSize        =   -1  'True
            Caption         =   "% de comis:"
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
            Left            =   3060
            TabIndex        =   117
            Top             =   195
            Width           =   1020
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Representante"
         Height          =   2355
         Left            =   3705
         TabIndex        =   113
         Top             =   420
         Width           =   5400
         Begin VB.Frame Frame14 
            Caption         =   "Exceções por produto"
            Height          =   1785
            Left            =   75
            TabIndex        =   157
            Top             =   510
            Width           =   5250
            Begin VB.CommandButton BotaoExcRep 
               Caption         =   "Representantes"
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
               Left            =   3600
               TabIndex        =   91
               Top             =   1470
               Width           =   1575
            End
            Begin VB.CommandButton BotaoExcRepProduto 
               Caption         =   "Produtos"
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
               Left            =   90
               TabIndex        =   90
               Top             =   1470
               Width           =   1530
            End
            Begin MSMask.MaskEdBox ExcRepPercComis 
               Height          =   210
               Left            =   2985
               TabIndex        =   158
               Top             =   525
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   370
               _Version        =   393216
               BorderStyle     =   0
               PromptInclude   =   0   'False
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
               Format          =   "0.00##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ExcRepProduto 
               Height          =   210
               Left            =   1980
               TabIndex        =   159
               Top             =   510
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   370
               _Version        =   393216
               BorderStyle     =   0
               AllowPrompt     =   -1  'True
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ExcRep 
               Height          =   225
               Left            =   375
               TabIndex        =   160
               Top             =   795
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   397
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               AllowPrompt     =   -1  'True
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin MSFlexGridLib.MSFlexGrid GridExcRep 
               Height          =   1260
               Left            =   60
               TabIndex        =   89
               Top             =   195
               Width           =   5145
               _ExtentX        =   9075
               _ExtentY        =   2223
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
         Begin MSMask.MaskEdBox Representante 
            Height          =   315
            Left            =   990
            TabIndex        =   87
            Top             =   210
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox PercComiRep 
            Height          =   315
            Left            =   4110
            TabIndex        =   88
            Top             =   210
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin VB.Label Label68 
            AutoSize        =   -1  'True
            Caption         =   "% de comis.:"
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
            Left            =   3030
            TabIndex        =   115
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label LabelRep 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   300
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   114
            Top             =   240
            Width           =   660
         End
      End
      Begin VB.Frame SSFrame7 
         Height          =   465
         Index           =   1
         Left            =   45
         TabIndex        =   69
         Top             =   -45
         Width           =   9075
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
            TabIndex        =   78
            Top             =   150
            Width           =   630
         End
         Begin VB.Label ClienteLabel 
            Height          =   210
            Index           =   0
            Left            =   960
            TabIndex        =   79
            Top             =   150
            Width           =   7080
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5130
      Index           =   1
      Left            =   150
      TabIndex        =   64
      Top             =   555
      Width           =   9180
      Begin VB.Frame Frame21 
         Caption         =   "Considerar os períodos abaixo como vendas do Call Center"
         Height          =   1155
         Left            =   4380
         TabIndex        =   177
         Top             =   3930
         Width           =   4695
         Begin MSMask.MaskEdBox DataCallCenterAte 
            Height          =   255
            Left            =   1965
            TabIndex        =   180
            Top             =   435
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataCallCenterDe 
            Height          =   255
            Left            =   300
            TabIndex        =   179
            Top             =   375
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridCallCenter 
            Height          =   870
            Left            =   135
            TabIndex        =   178
            Top             =   195
            Width           =   4440
            _ExtentX        =   7832
            _ExtentY        =   1535
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
      Begin VB.ComboBox UsuRespCallCenter 
         Height          =   315
         Left            =   1515
         Sorted          =   -1  'True
         TabIndex        =   19
         Top             =   4770
         Width           =   2670
      End
      Begin VB.ComboBox CondicaoPagtoCC 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1515
         TabIndex        =   14
         Top             =   2745
         Width           =   2670
      End
      Begin VB.ComboBox Regiao 
         Height          =   315
         Left            =   1515
         TabIndex        =   17
         Top             =   3960
         Width           =   2670
      End
      Begin VB.ComboBox ComboCobrador 
         Height          =   315
         Left            =   1515
         Sorted          =   -1  'True
         TabIndex        =   18
         Top             =   4365
         Width           =   2670
      End
      Begin VB.ComboBox CondicaoPagto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1515
         TabIndex        =   13
         Top             =   2340
         Width           =   2670
      End
      Begin VB.ComboBox FilialEmpresaFat 
         Height          =   315
         ItemData        =   "ClientesTRV.ctx":0000
         Left            =   1515
         List            =   "ClientesTRV.ctx":0002
         TabIndex        =   10
         Text            =   "FilialEmpresa"
         Top             =   1560
         Width           =   2640
      End
      Begin VB.ComboBox FilialEmpresa 
         Height          =   315
         ItemData        =   "ClientesTRV.ctx":0004
         Left            =   6360
         List            =   "ClientesTRV.ctx":0006
         TabIndex        =   9
         Text            =   "FilialEmpresa"
         Top             =   1185
         Width           =   2640
      End
      Begin VB.ComboBox FilialEmpresaNF 
         Height          =   315
         ItemData        =   "ClientesTRV.ctx":0008
         Left            =   6360
         List            =   "ClientesTRV.ctx":000A
         TabIndex        =   11
         Text            =   "FilialEmpresa"
         Top             =   1560
         Width           =   2640
      End
      Begin VB.CheckBox ConsiderarAporte 
         Caption         =   "Considerar aporte da empresa pai"
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
         Left            =   5625
         TabIndex        =   3
         Top             =   90
         Width           =   3315
      End
      Begin VB.TextBox RazaoSocial 
         Height          =   315
         Left            =   1515
         TabIndex        =   4
         Top             =   435
         Width           =   3720
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
         Height          =   255
         Left            =   3210
         TabIndex        =   2
         Top             =   90
         Width           =   1215
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   2355
         Picture         =   "ClientesTRV.ctx":000C
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Numeração Automática"
         Top             =   75
         Width           =   300
      End
      Begin VB.Frame Frame7 
         Caption         =   "Categorias"
         Height          =   1650
         Left            =   4365
         TabIndex        =   20
         Top             =   2265
         Width           =   4710
         Begin VB.ComboBox ComboCategoriaCliente 
            Height          =   315
            Left            =   435
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   405
            Width           =   1770
         End
         Begin VB.ComboBox ComboCategoriaClienteItem 
            Height          =   315
            Left            =   2175
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   405
            Width           =   1995
         End
         Begin MSFlexGridLib.MSFlexGrid GridCategoria 
            Height          =   1425
            Left            =   150
            TabIndex        =   21
            Top             =   195
            Width           =   4425
            _ExtentX        =   7805
            _ExtentY        =   2514
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
      Begin VB.ComboBox Tipo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1515
         TabIndex        =   8
         Top             =   1185
         Width           =   2790
      End
      Begin VB.TextBox Observacao 
         Height          =   315
         Left            =   1515
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   1935
         Width           =   7485
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   1515
         TabIndex        =   0
         Top             =   60
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NomeReduzido 
         Height          =   315
         Left            =   1515
         TabIndex        =   6
         Top             =   810
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox EmpresaPai 
         Height          =   315
         Left            =   6360
         TabIndex        =   7
         Top             =   810
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Grupo 
         Height          =   315
         Left            =   6360
         TabIndex        =   5
         Top             =   435
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CGC 
         Height          =   315
         Left            =   1515
         TabIndex        =   15
         Top             =   3150
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   14
         Mask            =   "##############"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox InscricaoMunicipal 
         Height          =   315
         Left            =   1515
         TabIndex        =   16
         Top             =   3555
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   18
         PromptChar      =   " "
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
         Left            =   105
         TabIndex        =   161
         Top             =   4800
         Width           =   1365
      End
      Begin VB.Label CondicaoPagtoCCLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Cond. Pagto CC:"
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
         Left            =   -240
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   147
         Top             =   2805
         Width           =   1710
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
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
         Left            =   795
         TabIndex        =   143
         Top             =   3990
         Width           =   675
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Inscr. Municipal:"
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
         TabIndex        =   138
         Top             =   3615
         Width           =   1425
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CNPJ/CPF:"
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
         TabIndex        =   137
         Top             =   3210
         Width           =   990
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Usu. Cobrador:"
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
         Left            =   180
         TabIndex        =   131
         Top             =   4395
         Width           =   1290
      End
      Begin VB.Label CondicaoPagtoLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Cond. Pagto:"
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
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   130
         Top             =   2400
         Width           =   1185
      End
      Begin VB.Label Label73 
         AutoSize        =   -1  'True
         Caption         =   "Grupo:"
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
         Left            =   5730
         TabIndex        =   128
         Top             =   465
         Width           =   585
      End
      Begin VB.Label Label72 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Filial Fatura:"
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
         Left            =   405
         TabIndex        =   127
         Top             =   1620
         Width           =   1065
      End
      Begin VB.Label Label70 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5850
         TabIndex        =   126
         Top             =   1245
         Width           =   465
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "Filial Nota Fiscal:"
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
         Left            =   4830
         TabIndex        =   125
         Top             =   1635
         Width           =   1485
      End
      Begin VB.Label LabelEmpresaPai 
         AutoSize        =   -1  'True
         Caption         =   "Empresa Pai:"
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
         Left            =   5190
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   124
         Top             =   855
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Left            =   810
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   73
         Top             =   105
         Width           =   660
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Left            =   915
         TabIndex        =   74
         Top             =   495
         Width           =   555
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Left            =   60
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   75
         Top             =   855
         Width           =   1410
      End
      Begin VB.Label TipoClienteLabel 
         Alignment       =   1  'Right Justify
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
         Left            =   1020
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   76
         Top             =   1245
         Width           =   450
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Observação:"
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
         TabIndex        =   77
         Top             =   2010
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4965
      Index           =   4
      Left            =   150
      TabIndex        =   97
      Top             =   705
      Visible         =   0   'False
      Width           =   9105
      Begin VB.Frame Frame12 
         Caption         =   "Exceções por produto (Emissor: )"
         Height          =   2010
         Left            =   105
         TabIndex        =   39
         Top             =   2940
         Width           =   8985
         Begin MSMask.MaskEdBox PercComissProd 
            Height          =   240
            Left            =   150
            TabIndex        =   41
            Top             =   780
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
         Begin VB.TextBox DescricaoProduto 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   3555
            MaxLength       =   250
            TabIndex        =   43
            Top             =   750
            Width           =   4605
         End
         Begin MSMask.MaskEdBox Produto 
            Height          =   240
            Left            =   1335
            TabIndex        =   42
            Top             =   705
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.CommandButton BotaoProdutos 
            Caption         =   "Produtos"
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
            Left            =   120
            TabIndex        =   44
            Top             =   1620
            Width           =   1815
         End
         Begin MSFlexGridLib.MSFlexGrid GridExcecoes 
            Height          =   1110
            Left            =   60
            TabIndex        =   40
            Top             =   210
            Width           =   8835
            _ExtentX        =   15584
            _ExtentY        =   1958
            _Version        =   393216
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Emissores"
         Height          =   2355
         Left            =   105
         TabIndex        =   34
         Top             =   570
         Width           =   8985
         Begin MSMask.MaskEdBox EmiPercCI 
            Height          =   315
            Left            =   4950
            TabIndex        =   187
            Top             =   585
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
         Begin MSMask.MaskEdBox EmiCPF 
            Height          =   315
            Left            =   1695
            TabIndex        =   189
            Top             =   615
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   14
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
         Begin VB.TextBox EmiCartao 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   5895
            MaxLength       =   100
            TabIndex        =   188
            Top             =   540
            Width           =   1665
         End
         Begin VB.ComboBox EmiCargo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7515
            Style           =   2  'Dropdown List
            TabIndex        =   186
            Top             =   480
            Width           =   1380
         End
         Begin MSMask.MaskEdBox PercComiss 
            Height          =   315
            Left            =   3855
            TabIndex        =   37
            Top             =   540
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
         Begin VB.TextBox Emissor 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   510
            MaxLength       =   50
            TabIndex        =   36
            Top             =   1185
            Width           =   2025
         End
         Begin VB.CommandButton BotaoEmissores 
            Caption         =   "Emissores"
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
            Left            =   105
            TabIndex        =   38
            Top             =   1965
            Width           =   1815
         End
         Begin MSFlexGridLib.MSFlexGrid GridComissao 
            Height          =   660
            Left            =   90
            TabIndex        =   35
            Top             =   195
            Width           =   8835
            _ExtentX        =   15584
            _ExtentY        =   1164
            _Version        =   393216
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
      End
      Begin VB.Frame SSFrame7 
         Height          =   555
         Index           =   3
         Left            =   90
         TabIndex        =   98
         Top             =   0
         Width           =   9000
         Begin VB.Label ClienteLabel 
            Height          =   210
            Index           =   3
            Left            =   960
            TabIndex        =   100
            Top             =   210
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
            Index           =   5
            Left            =   210
            TabIndex        =   99
            Top             =   210
            Width           =   630
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4950
      Index           =   5
      Left            =   150
      TabIndex        =   101
      Top             =   705
      Visible         =   0   'False
      Width           =   9225
      Begin VB.Frame Frame13 
         Caption         =   "Dados Finaceiros"
         Height          =   1575
         Left            =   75
         TabIndex        =   132
         Top             =   690
         Width           =   5025
         Begin VB.ComboBox Mensagem 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1710
            TabIndex        =   49
            Top             =   1200
            Width           =   3210
         End
         Begin VB.ComboBox TabelaPreco 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1710
            TabIndex        =   48
            Top             =   855
            Width           =   2220
         End
         Begin VB.CheckBox Bloqueado 
            Caption         =   "Crédito bloqueado"
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
            Left            =   3030
            TabIndex        =   47
            Top             =   540
            Width           =   1875
         End
         Begin MSMask.MaskEdBox LimiteCredito 
            Height          =   315
            Left            =   1710
            TabIndex        =   46
            Top             =   510
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto 
            Height          =   315
            Left            =   1710
            TabIndex        =   45
            Top             =   165
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Format          =   "#0.#0\%"
            PromptChar      =   " "
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
            Left            =   120
            TabIndex        =   136
            Top             =   555
            Width           =   1530
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
            Left            =   750
            TabIndex        =   135
            Top             =   210
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
            Left            =   360
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   134
            Top             =   1245
            Width           =   1305
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
            Left            =   75
            TabIndex        =   133
            Top             =   885
            Width           =   1590
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Outros"
         Height          =   4080
         Left            =   5145
         TabIndex        =   119
         Top             =   690
         Width           =   4035
         Begin VB.Frame Frame23 
            Caption         =   "Frame23"
            Height          =   720
            Left            =   180
            TabIndex        =   248
            Top             =   3210
            Visible         =   0   'False
            Width           =   1065
            Begin VB.ComboBox RegimeTributario 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1680
               TabIndex        =   250
               Top             =   0
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
               Left            =   240
               TabIndex        =   249
               Top             =   390
               Value           =   1  'Checked
               Width           =   1530
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
               Left            =   0
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   251
               Top             =   30
               Width           =   1560
            End
         End
         Begin VB.TextBox Observacao2 
            Height          =   1095
            Left            =   1530
            MaxLength       =   150
            MultiLine       =   -1  'True
            TabIndex        =   63
            Top             =   2895
            Width           =   2400
         End
         Begin VB.TextBox Guia 
            Height          =   300
            Left            =   1530
            MaxLength       =   10
            TabIndex        =   57
            Top             =   615
            Width           =   2055
         End
         Begin MSMask.MaskEdBox ContaContabil 
            Height          =   315
            Left            =   1530
            TabIndex        =   56
            Top             =   240
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox InscricaoEstadual 
            Height          =   315
            Left            =   1530
            TabIndex        =   59
            Top             =   1365
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   18
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox InscricaoSuframa 
            Height          =   315
            Left            =   1530
            TabIndex        =   60
            Top             =   1740
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   12
            Mask            =   "##.####-##-#"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox RG 
            Height          =   315
            Left            =   1530
            TabIndex        =   58
            Top             =   990
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FreqVisitas 
            Height          =   315
            Left            =   1530
            TabIndex        =   62
            Top             =   2520
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
            Left            =   1530
            TabIndex        =   61
            Top             =   2130
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
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
            Left            =   180
            TabIndex        =   252
            Top             =   3645
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1155
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
            Left            =   315
            TabIndex        =   146
            Top             =   2190
            Width           =   1125
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
            Left            =   2025
            TabIndex        =   145
            Top             =   2580
            Width           =   360
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            Caption         =   "Freq. de Visitas:"
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
            Left            =   60
            TabIndex        =   144
            Top             =   2610
            Width           =   1395
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Inscr. Estadual:"
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
            Left            =   105
            TabIndex        =   142
            Top             =   1425
            Width           =   1350
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Observação:"
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
            Left            =   315
            TabIndex        =   141
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Inscr. Suframa:"
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
            TabIndex        =   140
            Top             =   1770
            Width           =   1305
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
            Left            =   1095
            TabIndex        =   139
            Top             =   1050
            Width           =   345
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
            Left            =   960
            TabIndex        =   129
            Top             =   645
            Width           =   555
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
            Left            =   105
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   120
            Top             =   300
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Transporte"
         Height          =   1605
         Left            =   90
         TabIndex        =   108
         Top             =   3150
         Width           =   5010
         Begin VB.ComboBox Transportadora 
            Height          =   315
            Left            =   1695
            TabIndex        =   53
            Top             =   480
            Width           =   3105
         End
         Begin VB.ComboBox TipoFrete 
            Height          =   315
            ItemData        =   "ClientesTRV.ctx":00F6
            Left            =   1695
            List            =   "ClientesTRV.ctx":0100
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   120
            Width           =   1260
         End
         Begin VB.Frame Frame3 
            Caption         =   "Redespacho"
            Height          =   795
            Left            =   60
            TabIndex        =   109
            Top             =   750
            Width           =   4830
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
               TabIndex        =   55
               Top             =   495
               Width           =   2100
            End
            Begin VB.ComboBox TranspRedespacho 
               Height          =   315
               Left            =   1635
               TabIndex        =   54
               Top             =   180
               Width           =   3105
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
               Left            =   255
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   110
               Top             =   225
               Width           =   1365
            End
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
            Left            =   240
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   112
            Top             =   540
            Width           =   1365
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
            Left            =   390
            TabIndex        =   111
            Top             =   180
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Cobrança"
         Height          =   900
         Left            =   90
         TabIndex        =   105
         Top             =   2250
         Width           =   5025
         Begin VB.ComboBox Cobrador 
            Height          =   315
            Left            =   1695
            TabIndex        =   51
            Top             =   495
            Width           =   2520
         End
         Begin VB.ComboBox PadraoCobranca 
            Height          =   315
            ItemData        =   "ClientesTRV.ctx":010E
            Left            =   1695
            List            =   "ClientesTRV.ctx":0110
            TabIndex        =   50
            Top             =   150
            Width           =   2520
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
            Left            =   840
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   107
            Top             =   570
            Width           =   840
         End
         Begin VB.Label PadraoCobrancaLabel 
            AutoSize        =   -1  'True
            Caption         =   "Padrão de Cobr.:"
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
            TabIndex        =   106
            Top             =   210
            Width           =   1455
         End
      End
      Begin VB.Frame SSFrame7 
         Height          =   555
         Index           =   4
         Left            =   60
         TabIndex        =   102
         Top             =   0
         Width           =   9120
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
            Index           =   6
            Left            =   210
            TabIndex        =   104
            Top             =   210
            Width           =   630
         End
         Begin VB.Label ClienteLabel 
            Height          =   210
            Index           =   4
            Left            =   960
            TabIndex        =   103
            Top             =   210
            Width           =   7080
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4575
      Index           =   6
      Left            =   270
      TabIndex        =   190
      Top             =   840
      Visible         =   0   'False
      Width           =   8970
      Begin VB.Frame SSFrame7 
         Height          =   555
         Index           =   0
         Left            =   15
         TabIndex        =   245
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
            TabIndex        =   247
            Top             =   210
            Width           =   630
         End
         Begin VB.Label ClienteLabel 
            Height          =   210
            Index           =   1
            Left            =   960
            TabIndex        =   246
            Top             =   210
            Width           =   7080
         End
      End
      Begin VB.Frame SSFrame3 
         Caption         =   "Atraso"
         Height          =   720
         Left            =   30
         TabIndex        =   236
         Top             =   3105
         Width           =   6690
         Begin VB.Label MaiorAtraso 
            Caption         =   "0"
            Height          =   210
            Left            =   825
            TabIndex        =   244
            Top             =   435
            Width           =   720
         End
         Begin VB.Label MediaAtraso 
            Caption         =   "0"
            Height          =   210
            Left            =   825
            TabIndex        =   243
            Top             =   195
            Width           =   720
         End
         Begin VB.Label ValorPagtosAtraso 
            Caption         =   "0,00"
            Height          =   210
            Left            =   4530
            TabIndex        =   242
            Top             =   435
            Width           =   1395
         End
         Begin VB.Label SaldoAtrasados 
            Caption         =   "0,00"
            Height          =   210
            Left            =   4530
            TabIndex        =   241
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
            TabIndex        =   240
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
            TabIndex        =   239
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
            TabIndex        =   238
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
            TabIndex        =   237
            Top             =   195
            Width           =   1725
         End
      End
      Begin VB.Frame SSFrame2 
         Caption         =   "Cheques Devolvidos"
         Height          =   720
         Left            =   6765
         TabIndex        =   231
         Top             =   3105
         Width           =   2115
         Begin VB.Label DataUltChequeDevolvido 
            Caption         =   "  /  /    "
            Height          =   210
            Left            =   765
            TabIndex        =   235
            Top             =   450
            Width           =   1170
         End
         Begin VB.Label NumChequesDevolvidos 
            Caption         =   "0"
            Height          =   210
            Left            =   945
            TabIndex        =   234
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
            TabIndex        =   233
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
            TabIndex        =   232
            Top             =   450
            Width           =   600
         End
      End
      Begin VB.Frame SSFrame1 
         Caption         =   "Compras"
         Height          =   1500
         Left            =   4050
         TabIndex        =   220
         Top             =   570
         Width           =   4830
         Begin VB.Label DataUltimaCompra 
            Caption         =   "  /  /    "
            Height          =   210
            Left            =   3255
            TabIndex        =   230
            Top             =   735
            Width           =   1170
         End
         Begin VB.Label DataPrimeiraCompra 
            Caption         =   "  /  /    "
            Height          =   210
            Left            =   3255
            TabIndex        =   229
            Top             =   300
            Width           =   1170
         End
         Begin VB.Label ValorAcumuladoCompras 
            Caption         =   "0,00"
            Height          =   210
            Left            =   1830
            TabIndex        =   228
            Top             =   1170
            Width           =   1575
         End
         Begin VB.Label MediaCompra 
            Caption         =   "0,00"
            Height          =   210
            Left            =   900
            TabIndex        =   227
            Top             =   735
            Width           =   1410
         End
         Begin VB.Label NumeroCompras 
            Caption         =   "0"
            Height          =   210
            Left            =   1050
            TabIndex        =   226
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
            TabIndex        =   225
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
            TabIndex        =   224
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
            TabIndex        =   223
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
            TabIndex        =   222
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
            TabIndex        =   221
            Top             =   1185
            Width           =   1500
         End
      End
      Begin VB.Frame SSFrame4 
         Caption         =   "Saldos"
         Height          =   1500
         Left            =   30
         TabIndex        =   209
         Top             =   570
         Width           =   3705
         Begin VB.Label SaldoTitulos 
            Caption         =   "0,00"
            Height          =   210
            Left            =   1860
            TabIndex        =   213
            Top             =   300
            Width           =   1575
         End
         Begin VB.Label SaldoPedidosLiberados 
            Caption         =   "0,00"
            Height          =   210
            Left            =   1860
            TabIndex        =   212
            Top             =   585
            Width           =   1575
         End
         Begin VB.Label SaldodeCredito 
            Caption         =   "0,00"
            Height          =   210
            Left            =   1860
            TabIndex        =   211
            Top             =   1215
            Width           =   1575
         End
         Begin VB.Label SaldoLimitedeCredito 
            Caption         =   "0,00"
            Height          =   210
            Left            =   1860
            TabIndex        =   210
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
            TabIndex        =   219
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
            TabIndex        =   218
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
            TabIndex        =   217
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
            TabIndex        =   216
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
            TabIndex        =   215
            Top             =   585
            Width           =   1650
         End
         Begin VB.Label SaldoDuplicatas 
            Caption         =   "0,00"
            Height          =   210
            Left            =   1860
            TabIndex        =   214
            Top             =   -30
            Visible         =   0   'False
            Width           =   1575
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "Total em Títulos"
         Height          =   930
         Left            =   30
         TabIndex        =   196
         Top             =   2100
         Width           =   8865
         Begin VB.Label TotalCRComProtesto 
            Caption         =   "0,00"
            Height          =   210
            Left            =   7140
            TabIndex        =   198
            Top             =   270
            Width           =   1575
         End
         Begin VB.Label PercCRComProtesto 
            Caption         =   "0%"
            Height          =   210
            Left            =   7140
            TabIndex        =   197
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label TotalCREmCartorio 
            Caption         =   "0,00"
            Height          =   210
            Left            =   4335
            TabIndex        =   200
            Top             =   285
            Width           =   1575
         End
         Begin VB.Label PercCREmCartorio 
            Caption         =   "0%"
            Height          =   210
            Left            =   4335
            TabIndex        =   199
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
            TabIndex        =   202
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
            TabIndex        =   201
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
            TabIndex        =   208
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
            TabIndex        =   207
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label PercCREmAberto 
            Caption         =   "0%"
            Height          =   210
            Left            =   1860
            TabIndex        =   206
            Top             =   630
            Width           =   1575
         End
         Begin VB.Label TotalCR 
            Caption         =   "0,00"
            Height          =   210
            Left            =   1860
            TabIndex        =   205
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
            TabIndex        =   204
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
            TabIndex        =   203
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
         TabIndex        =   195
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
         TabIndex        =   194
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
         TabIndex        =   193
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
         TabIndex        =   192
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
         TabIndex        =   191
         Top             =   3870
         Width           =   1725
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5040
      Index           =   2
      Left            =   135
      TabIndex        =   65
      Top             =   615
      Visible         =   0   'False
      Width           =   9150
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3810
         Index           =   1
         Left            =   135
         TabIndex        =   70
         Top             =   1140
         Visible         =   0   'False
         Width           =   8685
         Begin TelasFATTRV.TabEndereco TabEnd 
            Height          =   3675
            Index           =   1
            Left            =   150
            TabIndex        =   181
            Top             =   180
            Width           =   8370
            _ExtentX        =   14764
            _ExtentY        =   6482
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3810
         Index           =   2
         Left            =   135
         TabIndex        =   184
         Top             =   1140
         Visible         =   0   'False
         Width           =   8685
         Begin TelasFATTRV.TabEndereco TabEnd 
            Height          =   3675
            Index           =   2
            Left            =   150
            TabIndex        =   185
            Top             =   180
            Width           =   8370
            _ExtentX        =   14764
            _ExtentY        =   6482
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3810
         Index           =   0
         Left            =   135
         TabIndex        =   182
         Top             =   1140
         Width           =   8685
         Begin TelasFATTRV.TabEndereco TabEnd 
            Height          =   3675
            Index           =   0
            Left            =   150
            TabIndex        =   183
            Top             =   180
            Width           =   8370
            _ExtentX        =   14764
            _ExtentY        =   6482
         End
      End
      Begin VB.Frame SSFrame5 
         Caption         =   "Endereços"
         Height          =   525
         Left            =   240
         TabIndex        =   71
         Top             =   600
         Width           =   8610
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
            TabIndex        =   31
            Top             =   180
            Value           =   -1  'True
            Width           =   1200
         End
         Begin VB.OptionButton OpcaoEndereco 
            Caption         =   "Call Center"
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
            TabIndex        =   32
            Top             =   180
            Width           =   1440
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
            TabIndex        =   33
            Top             =   180
            Width           =   1350
         End
      End
      Begin VB.Frame SSFrame7 
         Height          =   555
         Index           =   2
         Left            =   240
         TabIndex        =   72
         Top             =   0
         Width           =   8610
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
            TabIndex        =   80
            Top             =   210
            Width           =   630
         End
         Begin VB.Label ClienteLabel 
            Height          =   210
            Index           =   2
            Left            =   960
            TabIndex        =   81
            Top             =   210
            Width           =   7080
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4140
      Index           =   99
      Left            =   240
      TabIndex        =   162
      Top             =   1065
      Visible         =   0   'False
      Width           =   8850
      Begin VB.Frame SSFrame7 
         Height          =   555
         Index           =   5
         Left            =   240
         TabIndex        =   174
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
            Index           =   3
            Left            =   210
            TabIndex        =   176
            Top             =   210
            Width           =   630
         End
         Begin VB.Label ClienteLabel 
            Height          =   210
            Index           =   5
            Left            =   960
            TabIndex        =   175
            Top             =   210
            Width           =   7080
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "Faturamento"
         Height          =   2385
         Left            =   255
         TabIndex        =   163
         Top             =   765
         Width           =   8430
         Begin VB.Frame Frame20 
            Caption         =   "Faixa de faturamento"
            Height          =   1125
            Left            =   600
            TabIndex        =   169
            Top             =   1035
            Width           =   3525
            Begin MSMask.MaskEdBox PercentMaisReceb 
               Height          =   315
               Left            =   2310
               TabIndex        =   170
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
               TabIndex        =   171
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
            Begin VB.Label Label78 
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
               TabIndex        =   173
               Top             =   780
               Width           =   1950
            End
            Begin VB.Label Label77 
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
               TabIndex        =   172
               Top             =   375
               Width           =   1785
            End
         End
         Begin VB.Frame Frame19 
            Caption         =   "Faturamento fora da faixa"
            Height          =   1125
            Left            =   4290
            TabIndex        =   166
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
               TabIndex        =   168
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
               TabIndex        =   167
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
            TabIndex        =   165
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
            TabIndex        =   164
            Top             =   345
            Width           =   3585
         End
      End
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
      Height          =   435
      Left            =   7815
      Picture         =   "ClientesTRV.ctx":0112
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5730
      Width           =   1575
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
      Height          =   435
      Left            =   1800
      Picture         =   "ClientesTRV.ctx":0EB4
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5730
      Width           =   1575
   End
   Begin VB.CommandButton BotaoAcordos 
      Caption         =   "Acordos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   24
      Top             =   5730
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7245
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   0
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "ClientesTRV.ctx":2F6A
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "ClientesTRV.ctx":30C4
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1110
         Picture         =   "ClientesTRV.ctx":324E
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1620
         Picture         =   "ClientesTRV.ctx":3780
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5490
      Left            =   120
      TabIndex        =   68
      Top             =   240
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   9684
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Endereços"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Comissões"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Emissores"
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
    
    Set objCT.gobjInfoUsu = New CTClientesVGTRV
    Set objCT.gobjInfoUsu.gobjTelaUsu = New CTClientesTRV

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
    Call objCT.gobjInfoUsu.gobjTelaUsu.UserControl_KeyDown(objCT, KeyCode, Shift)
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

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label1_Click()
     Call objCT.Label1_Click
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label3_Click()
     Call objCT.Label3_Click
End Sub

Private Sub TipoClienteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoClienteLabel, Source, X, Y)
End Sub

Private Sub TipoClienteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoClienteLabel, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

'Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label13, Source, X, Y)
'End Sub
'
'Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
'End Sub

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

Private Sub Label32_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label32, Source, X, Y)
End Sub

Private Sub Label32_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label32, Button, Shift, X, Y)
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

Private Sub GridComissao_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridComissao_Click(objCT)
End Sub

Private Sub GridComissao_EnterCell()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridComissao_EnterCell(objCT)
End Sub

Private Sub GridComissao_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridComissao_GotFocus(objCT)
End Sub

Private Sub GridComissao_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridComissao_KeyDown(objCT, KeyCode, Shift)
End Sub

Private Sub GridComissao_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridComissao_KeyPress(objCT, KeyAscii)
End Sub

Private Sub GridComissao_LeaveCell()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridComissao_LeaveCell(objCT)
End Sub

Private Sub GridComissao_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridComissao_Validate(objCT, Cancel)
End Sub

Private Sub GridComissao_RowColChange()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridComissao_RowColChange(objCT)
End Sub

Private Sub GridComissao_Scroll()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridComissao_Scroll(objCT)
End Sub

Private Sub GridExcecoes_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcecoes_Click(objCT)
End Sub

Private Sub GridExcecoes_EnterCell()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcecoes_EnterCell(objCT)
End Sub

Private Sub GridExcecoes_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcecoes_GotFocus(objCT)
End Sub

Private Sub GridExcecoes_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcecoes_KeyDown(objCT, KeyCode, Shift)
End Sub

Private Sub GridExcecoes_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcecoes_KeyPress(objCT, KeyAscii)
End Sub

Private Sub GridExcecoes_LeaveCell()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcecoes_LeaveCell(objCT)
End Sub

Private Sub GridExcecoes_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcecoes_Validate(objCT, Cancel)
End Sub

Private Sub GridExcecoes_RowColChange()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcecoes_RowColChange(objCT)
End Sub

Private Sub GridExcecoes_Scroll()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcecoes_Scroll(objCT)
End Sub

Private Sub Emissor_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.Emissor_Change(objCT)
End Sub

Private Sub Emissor_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.Emissor_GotFocus(objCT)
End Sub

Private Sub Emissor_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.Emissor_KeyPress(objCT, KeyAscii)
End Sub

Private Sub Emissor_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.Emissor_Validate(objCT, Cancel)
End Sub

Private Sub PercComiss_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.PercComiss_Change(objCT)
End Sub

Private Sub PercComiss_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.PercComiss_GotFocus(objCT)
End Sub

Private Sub PercComiss_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.PercComiss_KeyPress(objCT, KeyAscii)
End Sub

Private Sub PercComiss_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.PercComiss_Validate(objCT, Cancel)
End Sub

Private Sub Produto_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.Produto_Change(objCT)
End Sub

Private Sub Produto_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.Produto_GotFocus(objCT)
End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.Produto_KeyPress(objCT, KeyAscii)
End Sub

Private Sub Produto_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.Produto_Validate(objCT, Cancel)
End Sub

Private Sub PercComissProd_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.PercComissProd_Change(objCT)
End Sub

Private Sub PercComissProd_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.PercComissProd_GotFocus(objCT)
End Sub

Private Sub PercComissProd_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.PercComissProd_KeyPress(objCT, KeyAscii)
End Sub

Private Sub PercComissProd_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.PercComissProd_Validate(objCT, Cancel)
End Sub

Private Sub BotaoAcordos_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.BotaoAcordos_Click(objCT)
End Sub

Private Sub BotaoEmissores_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.BotaoEmissores_Click(objCT)
End Sub

Private Sub BotaoProdutos_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.BotaoProdutos_Click(objCT)
End Sub

Private Sub Representante_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.Representante_Change(objCT)
End Sub

Private Sub Representante_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.Representante_Validate(objCT, Cancel)
End Sub

Private Sub LabelRep_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.LabelRep_Click(objCT)
End Sub

Private Sub PercComiRep_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.PercComiRep_Change(objCT)
End Sub

Private Sub PercComiRep_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.PercComiRep_Validate(objCT, Cancel)
End Sub

Private Sub Correntista_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.Correntista_Change(objCT)
End Sub

Private Sub Correntista_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.Correntista_Validate(objCT, Cancel)
End Sub

Private Sub LabelCorr_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.LabelCorr_Click(objCT)
End Sub

Private Sub PercComiCorr_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.PercComiCorr_Change(objCT)
End Sub

Private Sub PercComiCorr_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.PercComiCorr_Validate(objCT, Cancel)
End Sub

Private Sub PercComiAg_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.PercComiAg_Change(objCT)
End Sub

Private Sub PercComiAg_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.PercComiAg_Validate(objCT, Cancel)
End Sub

Private Sub ConsiderarAporte_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ConsiderarAporte_Click(objCT)
End Sub

Public Sub EmpresaPai_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.EmpresaPai_Change(objCT)
End Sub

Public Sub EmpresaPai_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.EmpresaPai_Validate(objCT, Cancel)
End Sub

Private Sub LabelEmpresaPai_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.LabelEmpresaPai_Click(objCT)
End Sub

Private Sub FilialEmpresa_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.FilialEmpresa_Click(objCT, FilialEmpresa)
End Sub

Private Sub FilialEmpresa_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.FilialEmpresa_Validate(objCT, Cancel, FilialEmpresa)
End Sub

Private Sub FilialEmpresaFat_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.FilialEmpresa_Click(objCT, FilialEmpresaFat)
End Sub

Private Sub FilialEmpresaFat_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.FilialEmpresa_Validate(objCT, Cancel, FilialEmpresaFat)
End Sub

Private Sub FilialEmpresaNF_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.FilialEmpresa_Click(objCT, FilialEmpresaNF)
End Sub

Private Sub FilialEmpresaNF_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.FilialEmpresa_Validate(objCT, Cancel, FilialEmpresaNF)
End Sub

Private Sub Grupo_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.Grupo_Change(objCT)
End Sub

Private Sub GridExcRep_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcRep_Click(objCT)
End Sub

Private Sub GridExcRep_EnterCell()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcRep_EnterCell(objCT)
End Sub

Private Sub GridExcRep_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcRep_GotFocus(objCT)
End Sub

Private Sub GridExcRep_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcRep_KeyDown(objCT, KeyCode, Shift)
End Sub

Private Sub GridExcRep_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcRep_KeyPress(objCT, KeyAscii)
End Sub

Private Sub GridExcRep_LeaveCell()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcRep_LeaveCell(objCT)
End Sub

Private Sub GridExcRep_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcRep_Validate(objCT, Cancel)
End Sub

Private Sub GridExcRep_RowColChange()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcRep_RowColChange(objCT)
End Sub

Private Sub GridExcRep_Scroll()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcRep_Scroll(objCT)
End Sub

Private Sub ExcRep_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcRep_Change(objCT)
End Sub

Private Sub ExcRep_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcRep_GotFocus(objCT)
End Sub

Private Sub ExcRep_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcRep_KeyPress(objCT, KeyAscii)
End Sub

Private Sub ExcRep_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcRep_Validate(objCT, Cancel)
End Sub

Private Sub ExcRepProduto_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcRepProduto_Change(objCT)
End Sub

Private Sub ExcRepProduto_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcRepProduto_GotFocus(objCT)
End Sub

Private Sub ExcRepProduto_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcRepProduto_KeyPress(objCT, KeyAscii)
End Sub

Private Sub ExcRepProduto_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcRepProduto_Validate(objCT, Cancel)
End Sub

Private Sub ExcRepPercComis_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcRepPercComis_Change(objCT)
End Sub

Private Sub ExcRepPercComis_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcRepPercComis_GotFocus(objCT)
End Sub

Private Sub ExcRepPercComis_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcRepPercComis_KeyPress(objCT, KeyAscii)
End Sub

Private Sub ExcRepPercComis_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcRepPercComis_Validate(objCT, Cancel)
End Sub

Private Sub GridExcCor_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcCor_Click(objCT)
End Sub

Private Sub GridExcCor_EnterCell()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcCor_EnterCell(objCT)
End Sub

Private Sub GridExcCor_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcCor_GotFocus(objCT)
End Sub

Private Sub GridExcCor_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcCor_KeyDown(objCT, KeyCode, Shift)
End Sub

Private Sub GridExcCor_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcCor_KeyPress(objCT, KeyAscii)
End Sub

Private Sub GridExcCor_LeaveCell()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcCor_LeaveCell(objCT)
End Sub

Private Sub GridExcCor_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcCor_Validate(objCT, Cancel)
End Sub

Private Sub GridExcCor_RowColChange()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcCor_RowColChange(objCT)
End Sub

Private Sub GridExcCor_Scroll()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcCor_Scroll(objCT)
End Sub

Private Sub ExcCor_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcCor_Change(objCT)
End Sub

Private Sub ExcCor_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcCor_GotFocus(objCT)
End Sub

Private Sub ExcCor_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcCor_KeyPress(objCT, KeyAscii)
End Sub

Private Sub ExcCor_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcCor_Validate(objCT, Cancel)
End Sub

Private Sub ExcCorProduto_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcCorProduto_Change(objCT)
End Sub

Private Sub ExcCorProduto_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcCorProduto_GotFocus(objCT)
End Sub

Private Sub ExcCorProduto_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcCorProduto_KeyPress(objCT, KeyAscii)
End Sub

Private Sub ExcCorProduto_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcCorProduto_Validate(objCT, Cancel)
End Sub

Private Sub ExcCorPercComis_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcCorPercComis_Change(objCT)
End Sub

Private Sub ExcCorPercComis_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcCorPercComis_GotFocus(objCT)
End Sub

Private Sub ExcCorPercComis_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcCorPercComis_KeyPress(objCT, KeyAscii)
End Sub

Private Sub ExcCorPercComis_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcCorPercComis_Validate(objCT, Cancel)
End Sub

Private Sub BotaoExcCorProduto_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.BotaoExcCorProduto_Click(objCT)
End Sub

Private Sub BotaoExcCor_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.BotaoExcCor_Click(objCT)
End Sub

Private Sub BotaoExcRepProduto_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.BotaoExcRepProduto_Click(objCT)
End Sub

Private Sub BotaoExcRep_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.BotaoExcRep_Click(objCT)
End Sub

Private Sub CondicaoPagtoCC_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.CondicaoPagtoCC_Change(objCT)
End Sub

Private Sub CondicaoPagtoCC_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.CondicaoPagtoCC_Click(objCT)
End Sub

Private Sub CondicaoPagtoCC_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.CondicaoPagtoCC_Validate(objCT, Cancel)
End Sub

Private Sub CondicaoPagtoCCLabel_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.CondicaoPagtoCCLabel_Click(objCT)
End Sub

Private Sub GridExcAg_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcAg_Click(objCT)
End Sub

Private Sub GridExcAg_EnterCell()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcAg_EnterCell(objCT)
End Sub

Private Sub GridExcAg_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcAg_GotFocus(objCT)
End Sub

Private Sub GridExcAg_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcAg_KeyDown(objCT, KeyCode, Shift)
End Sub

Private Sub GridExcAg_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcAg_KeyPress(objCT, KeyAscii)
End Sub

Private Sub GridExcAg_LeaveCell()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcAg_LeaveCell(objCT)
End Sub

Private Sub GridExcAg_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcAg_Validate(objCT, Cancel)
End Sub

Private Sub GridExcAg_RowColChange()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcAg_RowColChange(objCT)
End Sub

Private Sub GridExcAg_Scroll()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExcAg_Scroll(objCT)
End Sub

Private Sub ExcAgProduto_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcAgProduto_Change(objCT)
End Sub

Private Sub ExcAgProduto_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcAgProduto_GotFocus(objCT)
End Sub

Private Sub ExcAgProduto_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcAgProduto_KeyPress(objCT, KeyAscii)
End Sub

Private Sub ExcAgProduto_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcAgProduto_Validate(objCT, Cancel)
End Sub

Private Sub ExcAgPercComis_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcAgPercComis_Change(objCT)
End Sub

Private Sub ExcAgPercComis_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcAgPercComis_GotFocus(objCT)
End Sub

Private Sub ExcAgPercComis_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcAgPercComis_KeyPress(objCT, KeyAscii)
End Sub

Private Sub ExcAgPercComis_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcAgPercComis_Validate(objCT, Cancel)
End Sub

Private Sub BotaoExcAgProduto_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.BotaoExcAgProduto_Click(objCT)
End Sub

Private Sub UsuRespCallCenter_Click()
    objCT.UsuRespCallCenter_Click
End Sub

Private Sub UsuRespCallCenter_Validate(Cancel As Boolean)
    objCT.UsuRespCallCenter_Validate (Cancel)
End Sub

Private Sub DataCallCenterDe_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataCallCenterDe_Change(objCT)
End Sub

Private Sub DataCallCenterDe_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataCallCenterDe_GotFocus(objCT)
End Sub

Private Sub DataCallCenterDe_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataCallCenterDe_KeyPress(objCT, KeyAscii)
End Sub

Private Sub DataCallCenterDe_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataCallCenterDe_Validate(objCT, Cancel)
End Sub

Private Sub DataCallCenterAte_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataCallCenterAte_Change(objCT)
End Sub

Private Sub DataCallCenterAte_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataCallCenterAte_GotFocus(objCT)
End Sub

Private Sub DataCallCenterAte_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataCallCenterAte_KeyPress(objCT, KeyAscii)
End Sub

Private Sub DataCallCenterAte_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataCallCenterAte_Validate(objCT, Cancel)
End Sub

Private Sub GridCallCenter_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridCallCenter_Click(objCT)
End Sub

Private Sub GridCallCenter_EnterCell()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridCallCenter_EnterCell(objCT)
End Sub

Private Sub GridCallCenter_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridCallCenter_GotFocus(objCT)
End Sub

Private Sub GridCallCenter_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridCallCenter_KeyDown(objCT, KeyCode, Shift)
End Sub

Private Sub GridCallCenter_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridCallCenter_KeyPress(objCT, KeyAscii)
End Sub

Private Sub GridCallCenter_LeaveCell()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridCallCenter_LeaveCell(objCT)
End Sub

Private Sub GridCallCenter_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridCallCenter_Validate(objCT, Cancel)
End Sub

Private Sub GridCallCenter_RowColChange()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridCallCenter_RowColChange(objCT)
End Sub

Private Sub GridCallCenter_Scroll()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridCallCenter_Scroll(objCT)
End Sub

Private Sub EmiPercCI_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.EmiPercCI_Change(objCT)
End Sub

Private Sub EmiPercCI_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.EmiPercCI_GotFocus(objCT)
End Sub

Private Sub EmiPercCI_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.EmiPercCI_KeyPress(objCT, KeyAscii)
End Sub

Private Sub EmiPercCI_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.EmiPercCI_Validate(objCT, Cancel)
End Sub

Private Sub EmiCargo_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.EmiCargo_Change(objCT)
End Sub

Private Sub EmiCargo_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.EmiCargo_GotFocus(objCT)
End Sub

Private Sub EmiCargo_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.EmiCargo_KeyPress(objCT, KeyAscii)
End Sub

Private Sub EmiCargo_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.EmiCargo_Validate(objCT, Cancel)
End Sub

Private Sub EmiCartao_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.EmiCartao_Change(objCT)
End Sub

Private Sub EmiCartao_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.EmiCartao_GotFocus(objCT)
End Sub

Private Sub EmiCartao_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.EmiCartao_KeyPress(objCT, KeyAscii)
End Sub

Private Sub EmiCartao_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.EmiCartao_Validate(objCT, Cancel)
End Sub

Private Sub EmiCPF_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.EmiCPF_Change(objCT)
End Sub

Private Sub EmiCPF_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.EmiCPF_GotFocus(objCT)
End Sub

Private Sub EmiCPF_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.EmiCPF_KeyPress(objCT, KeyAscii)
End Sub

Private Sub EmiCPF_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.EmiCPF_Validate(objCT, Cancel)
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

Private Sub IENaoContrib_Click()
    Call objCT.IENaoContrib_Click
End Sub

Private Sub InscricaoEstadual_Validate(Cancel As Boolean)
    Call objCT.InscricaoEstadual_Validate(Cancel)
End Sub

Private Sub PercFatorDevCMCC_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.PercFatorDevCMCC_Change(objCT)
End Sub

Private Sub PercFatorDevCMCC_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.PercFatorDevCMCC_Validate(objCT, Cancel)
End Sub
