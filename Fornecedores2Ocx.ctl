VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.UserControl FornecedoresOcx 
   ClientHeight    =   5490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9270
   KeyPreview      =   -1  'True
   ScaleHeight     =   5490
   ScaleWidth      =   9270
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3630
      Index           =   4
      Left            =   180
      TabIndex        =   88
      Top             =   1080
      Visible         =   0   'False
      Width           =   8835
      Begin VB.TextBox IdEstrangeiro 
         Height          =   315
         Left            =   5940
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   114
         ToolTipText     =   "número do passaporte ou outro documento legal para identificar pessoa estrangeira"
         Top             =   2205
         Width           =   2070
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
         Left            =   5250
         TabIndex        =   109
         Top             =   1290
         Value           =   1  'Checked
         Width           =   2760
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
         Left            =   4245
         TabIndex        =   93
         Top             =   1320
         Value           =   1  'Checked
         Width           =   945
      End
      Begin VB.ComboBox RegimeTributario 
         Height          =   315
         ItemData        =   "Fornecedores2Ocx.ctx":0000
         Left            =   1920
         List            =   "Fornecedores2Ocx.ctx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   97
         Top             =   3270
         Width           =   2760
      End
      Begin VB.TextBox Observacao2 
         Height          =   315
         Left            =   1920
         MaxLength       =   150
         MultiLine       =   -1  'True
         TabIndex        =   96
         Top             =   2745
         Width           =   6165
      End
      Begin VB.Frame SSFrame2 
         Height          =   510
         Left            =   180
         TabIndex        =   98
         Top             =   30
         Width           =   8520
         Begin VB.Label Fornecedor 
            Height          =   210
            Index           =   1
            Left            =   1275
            TabIndex        =   100
            Top             =   180
            Width           =   6600
         End
         Begin VB.Label Label9 
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
            Height          =   210
            Left            =   180
            TabIndex        =   99
            Top             =   180
            Width           =   1110
         End
      End
      Begin VB.CheckBox GeraICMS 
         Caption         =   "Pode Gerar Credito ICMS"
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
         Left            =   4245
         TabIndex        =   90
         Top             =   780
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin MSMask.MaskEdBox CGC 
         Height          =   315
         Left            =   1920
         TabIndex        =   89
         Top             =   735
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   14
         Mask            =   "99999999999999"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox InscricaoEstadual 
         Height          =   315
         Left            =   1920
         TabIndex        =   91
         Top             =   1245
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   18
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox InscricaoMunicipal 
         Height          =   315
         Left            =   1920
         TabIndex        =   92
         Top             =   1755
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   18
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox InscricaoINSS 
         Height          =   300
         Left            =   1920
         TabIndex        =   95
         Top             =   2265
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   14
         Mask            =   "999.999.999-99"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox InscricaoSuframa 
         Height          =   315
         Left            =   5970
         TabIndex        =   94
         Top             =   1710
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   12
         Mask            =   "##.####-##-#"
         PromptChar      =   " "
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Ident. Estrangeiro:"
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
         Left            =   4170
         TabIndex        =   115
         Top             =   2250
         Width           =   1650
      End
      Begin VB.Label Label1 
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
         Left            =   300
         TabIndex        =   108
         Top             =   3315
         Width           =   1560
      End
      Begin VB.Label Label3 
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
         Left            =   4245
         TabIndex        =   107
         Top             =   1770
         Width           =   1605
      End
      Begin VB.Label Label32 
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
         Height          =   210
         Left            =   765
         TabIndex        =   105
         Top             =   2775
         Width           =   1125
      End
      Begin VB.Label Label35 
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
         Height          =   210
         Left            =   900
         TabIndex        =   104
         Top             =   780
         Width           =   1005
      End
      Begin VB.Label Label34 
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
         Height          =   210
         Left            =   240
         TabIndex        =   103
         Top             =   1290
         Width           =   1650
      End
      Begin VB.Label Label36 
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
         Height          =   210
         Left            =   150
         TabIndex        =   102
         Top             =   1800
         Width           =   1725
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição INSS:"
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
         TabIndex        =   101
         Top             =   2310
         Width           =   1350
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4155
      Index           =   3
      Left            =   180
      TabIndex        =   61
      Top             =   1050
      Visible         =   0   'False
      Width           =   8865
      Begin VB.Frame Frame4 
         Caption         =   "Contabilidade"
         Height          =   2265
         Left            =   180
         TabIndex        =   81
         Top             =   585
         Width           =   4200
         Begin MSMask.MaskEdBox ContaContabil 
            Height          =   315
            Left            =   2430
            TabIndex        =   82
            Top             =   420
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ContaFornConsig 
            Height          =   315
            Left            =   2430
            TabIndex        =   83
            Top             =   1035
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ContaDespesa 
            Height          =   315
            Left            =   2430
            TabIndex        =   84
            Top             =   1620
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label ContaContabilLabel 
            AutoSize        =   -1  'True
            Caption         =   "Conta de Fornecedores:"
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
            TabIndex        =   87
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label ContaFornConsigLabel 
            AutoSize        =   -1  'True
            Caption         =   "Conta de Fornecedores em Consignação:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   315
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   86
            Top             =   997
            Width           =   2055
            WordWrap        =   -1  'True
         End
         Begin VB.Label LabelContaDespesa 
            AutoSize        =   -1  'True
            Caption         =   "Conta Estoque/Despesa:"
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
            TabIndex        =   85
            Top             =   1680
            Width           =   2145
         End
      End
      Begin VB.Frame SSFrame6 
         Caption         =   "Conta Corrente"
         Height          =   690
         Left            =   180
         TabIndex        =   74
         Top             =   2895
         Width           =   8520
         Begin MSMask.MaskEdBox ContaCorrente 
            Height          =   315
            Left            =   6330
            TabIndex        =   75
            Top             =   225
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   14
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Agencia 
            Height          =   315
            Left            =   3855
            TabIndex        =   76
            Top             =   225
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   7
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Banco 
            Height          =   315
            Left            =   1680
            TabIndex        =   77
            Top             =   225
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   3
            Mask            =   "999"
            PromptChar      =   " "
         End
         Begin VB.Label Label22 
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
            Left            =   5550
            TabIndex        =   80
            Top             =   270
            Width           =   720
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Agência:"
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
            Left            =   3045
            TabIndex        =   79
            Top             =   270
            Width           =   765
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
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
            Left            =   1005
            TabIndex        =   78
            Top             =   270
            Width           =   615
         End
      End
      Begin VB.Frame SSFrame8 
         Height          =   510
         Left            =   180
         TabIndex        =   71
         Top             =   30
         Width           =   8490
         Begin VB.Label Fornecedor 
            Height          =   210
            Index           =   4
            Left            =   1260
            TabIndex        =   73
            Top             =   180
            Width           =   6810
         End
         Begin VB.Label Label24 
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
            Height          =   210
            Left            =   165
            TabIndex        =   72
            Top             =   180
            Width           =   1110
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Pagamento"
         Height          =   2265
         Left            =   4470
         TabIndex        =   62
         Top             =   585
         Width           =   4200
         Begin VB.ComboBox CondicaoPagto 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1140
            TabIndex        =   65
            Top             =   795
            Width           =   2070
         End
         Begin VB.ComboBox TipoCobranca 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   1800
            Width           =   2865
         End
         Begin VB.ComboBox TipoFrete 
            Height          =   315
            ItemData        =   "Fornecedores2Ocx.ctx":0004
            Left            =   1140
            List            =   "Fornecedores2Ocx.ctx":000E
            Style           =   2  'Dropdown List
            TabIndex        =   63
            Top             =   1305
            Width           =   1125
         End
         Begin MSMask.MaskEdBox Desconto 
            Height          =   315
            Left            =   1140
            TabIndex        =   66
            Top             =   285
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin VB.Label CondicaoPagtoLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Condição:"
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
            TabIndex        =   70
            Top             =   855
            Width           =   975
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   225
            TabIndex        =   69
            Top             =   345
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cobrança:"
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
            Index           =   0
            Left            =   225
            TabIndex        =   68
            Top             =   1860
            Width           =   885
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Frete:"
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
            TabIndex        =   67
            Top             =   1365
            Width           =   945
         End
      End
      Begin MSMask.MaskEdBox Natureza 
         Height          =   300
         Left            =   1290
         TabIndex        =   111
         Top             =   3720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label LabelNatureza 
         AutoSize        =   -1  'True
         Caption         =   "Natureza:"
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
         Height          =   180
         Left            =   405
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   113
         Top             =   3735
         Width           =   840
      End
      Begin VB.Label LabelNaturezaDesc 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2325
         TabIndex        =   112
         Top             =   3720
         Width           =   3210
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3525
      Index           =   2
      Left            =   150
      TabIndex        =   7
      Top             =   1020
      Visible         =   0   'False
      Width           =   8805
      Begin VB.Frame Frame7 
         Caption         =   "Categorias"
         Height          =   3120
         Left            =   1290
         TabIndex        =   52
         Top             =   150
         Width           =   6345
         Begin VB.ComboBox ComboCategoriaFornecedorItem 
            Height          =   315
            Left            =   3000
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   540
            Width           =   1635
         End
         Begin VB.ComboBox ComboCategoriaFornecedor 
            Height          =   315
            Left            =   1800
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   540
            Width           =   1545
         End
         Begin MSFlexGridLib.MSFlexGrid GridCategoria 
            Height          =   2520
            Left            =   1140
            TabIndex        =   55
            Top             =   270
            Width           =   3810
            _ExtentX        =   6720
            _ExtentY        =   4445
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
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3825
      Index           =   5
      Left            =   210
      TabIndex        =   8
      Top             =   1050
      Visible         =   0   'False
      Width           =   8775
      Begin VB.Frame SSFrame5 
         Height          =   510
         Left            =   90
         TabIndex        =   15
         Top             =   -60
         Width           =   8610
         Begin VB.Label Label11 
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
            Height          =   210
            Left            =   165
            TabIndex        =   17
            Top             =   180
            Width           =   1110
         End
         Begin VB.Label Fornecedor 
            Height          =   210
            Index           =   2
            Left            =   1275
            TabIndex        =   16
            Top             =   180
            Width           =   3300
         End
      End
      Begin TelasCpr.TabEndereco TabEnd 
         Height          =   3660
         Index           =   0
         Left            =   30
         TabIndex        =   110
         Top             =   480
         Width           =   8445
         _ExtentX        =   14896
         _ExtentY        =   6456
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3525
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   1020
      Width           =   8895
      Begin VB.TextBox RazaoSocial 
         Height          =   300
         Left            =   2385
         MaxLength       =   40
         TabIndex        =   3
         Top             =   1110
         Width           =   4095
      End
      Begin VB.CheckBox Ativo 
         Caption         =   "Ativo"
         Height          =   225
         Left            =   4815
         TabIndex        =   106
         Top             =   510
         Value           =   1  'Checked
         Width           =   930
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   3240
         Picture         =   "Fornecedores2Ocx.ctx":001C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Numeração Automática"
         Top             =   510
         Width           =   300
      End
      Begin VB.TextBox Observacao 
         Height          =   315
         Left            =   2400
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   2850
         Width           =   5160
      End
      Begin VB.ComboBox Tipo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2400
         TabIndex        =   5
         Top             =   2280
         Width           =   2790
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   2400
         TabIndex        =   1
         Top             =   510
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         Mask            =   "99999999"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NomeReduzido 
         Height          =   315
         Left            =   2400
         TabIndex        =   4
         Top             =   1740
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label LabelObservacao 
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
         Left            =   1230
         TabIndex        =   60
         Top             =   2910
         Width           =   1095
      End
      Begin VB.Label TipoFornecedorLabel 
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
         Left            =   1875
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   59
         Top             =   2370
         Width           =   450
      End
      Begin VB.Label LabelNomeReduzido 
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
         Left            =   915
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   58
         Top             =   1800
         Width           =   1410
      End
      Begin VB.Label Label2 
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
         Left            =   1770
         TabIndex        =   57
         Top             =   1170
         Width           =   555
      End
      Begin VB.Label LabelCodigo 
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
         Left            =   1665
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   56
         Top             =   600
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Index           =   6
      Left            =   120
      TabIndex        =   9
      Top             =   1110
      Visible         =   0   'False
      Width           =   8805
      Begin VB.Frame Frame2 
         Caption         =   "Pedidos de Compra"
         Height          =   1365
         Left            =   4830
         TabIndex        =   43
         Top             =   1950
         Width           =   3825
         Begin VB.Frame SSFrame4 
            Caption         =   "Atraso de Entrega"
            Height          =   600
            Left            =   150
            TabIndex        =   44
            Top             =   600
            Width           =   3585
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
               Left            =   465
               TabIndex        =   48
               Top             =   300
               Width           =   585
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
               Left            =   2055
               TabIndex        =   47
               Top             =   300
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label MediaAtraso 
               Caption         =   "0"
               Height          =   210
               Left            =   1110
               TabIndex        =   46
               Top             =   285
               Width           =   510
            End
            Begin VB.Label MaiorAtraso 
               Caption         =   "0"
               Height          =   210
               Left            =   2655
               TabIndex        =   45
               Top             =   300
               Visible         =   0   'False
               Width           =   510
            End
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Em Aberto:"
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
            TabIndex        =   50
            Top             =   300
            Width           =   945
         End
         Begin VB.Label PedidosEmAberto 
            Caption         =   "0"
            Height          =   210
            Left            =   1275
            TabIndex        =   49
            Top             =   315
            Width           =   585
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Devoluções"
         Height          =   1110
         Left            =   4860
         TabIndex        =   22
         Top             =   690
         Width           =   3780
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   1275
            TabIndex        =   28
            Top             =   825
            Width           =   600
         End
         Begin VB.Label Label4 
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
            Left            =   375
            TabIndex        =   27
            Top             =   255
            Width           =   1500
         End
         Begin VB.Label ValorAcumuladoDevolucoes 
            Caption         =   "0,00"
            Height          =   210
            Left            =   1980
            TabIndex        =   26
            Top             =   270
            Width           =   1575
         End
         Begin VB.Label Label19 
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
            Left            =   1155
            TabIndex        =   25
            Top             =   540
            Width           =   720
         End
         Begin VB.Label NumTotalDevolucoes 
            Caption         =   "0"
            Height          =   210
            Left            =   1980
            TabIndex        =   24
            Top             =   540
            Width           =   585
         End
         Begin VB.Label DataUltDevolucao 
            Caption         =   "  /  /    "
            Height          =   210
            Left            =   1980
            TabIndex        =   23
            Top             =   825
            Width           =   1170
         End
      End
      Begin VB.Frame SSFrame7 
         Height          =   510
         Left            =   240
         TabIndex        =   19
         Top             =   15
         Width           =   8520
         Begin VB.Label Fornecedor 
            Height          =   210
            Index           =   3
            Left            =   1260
            TabIndex        =   21
            Top             =   180
            Width           =   6750
         End
         Begin VB.Label Label12 
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
            Height          =   210
            Left            =   180
            TabIndex        =   20
            Top             =   180
            Width           =   1110
         End
      End
      Begin VB.Frame SSFrame1 
         Caption         =   "Compras"
         Height          =   1920
         Left            =   240
         TabIndex        =   29
         Top             =   1410
         Width           =   4200
         Begin VB.Label DataUltimaCompra 
            Caption         =   "  /  /    "
            Height          =   210
            Left            =   1830
            TabIndex        =   39
            Top             =   1530
            Width           =   1170
         End
         Begin VB.Label DataPrimeiraCompra 
            Caption         =   "  /  /    "
            Height          =   210
            Left            =   1830
            TabIndex        =   38
            Top             =   1200
            Width           =   1170
         End
         Begin VB.Label ValorAcumuladoCompras 
            Caption         =   "0,00"
            Height          =   210
            Left            =   1830
            TabIndex        =   37
            Top             =   900
            Width           =   1575
         End
         Begin VB.Label MediaCompra 
            Caption         =   "0,00"
            Height          =   210
            Left            =   1830
            TabIndex        =   36
            Top             =   600
            Width           =   1410
         End
         Begin VB.Label NumeroCompras 
            Caption         =   "0"
            Height          =   210
            Left            =   1830
            TabIndex        =   35
            Top             =   270
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
            Left            =   1035
            TabIndex        =   34
            Top             =   255
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
            Left            =   1170
            TabIndex        =   33
            Top             =   585
            Width           =   585
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   1155
            TabIndex        =   32
            Top             =   1530
            Width           =   600
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   1005
            TabIndex        =   31
            Top             =   1200
            Width           =   750
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
            TabIndex        =   30
            Top             =   900
            Width           =   1500
         End
      End
      Begin VB.Frame SSFrame3 
         Caption         =   "Saldos"
         Height          =   600
         Left            =   240
         TabIndex        =   40
         Top             =   720
         Width           =   4200
         Begin VB.Label SaldoTitulos 
            Caption         =   "0,00"
            Height          =   210
            Left            =   1815
            TabIndex        =   42
            Top             =   300
            Width           =   1575
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   780
            TabIndex        =   41
            Top             =   285
            Width           =   990
         End
      End
   End
   Begin VB.CommandButton Filiais 
      Height          =   525
      Left            =   4860
      Picture         =   "Fornecedores2Ocx.ctx":0106
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   90
      Width           =   1620
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6930
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "Fornecedores2Ocx.ctx":0EA8
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "Fornecedores2Ocx.ctx":1002
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "Fornecedores2Ocx.ctx":118C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "Fornecedores2Ocx.ctx":16BE
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   4620
      Left            =   90
      TabIndex        =   18
      Top             =   735
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   8149
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Complemento"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Financeiros"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Inscrições"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Endereço"
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
Attribute VB_Name = "FornecedoresOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_objUserControl As Object

Public gobjAnotacao As ClassAnotacoes
Public gobjTabEnd As New ClassTabEndereco

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iFrameAtual As Integer
Public iAlterado As Integer

Dim sIEAnt As String

'********** Janaina **********
Dim objGrid2 As AdmGrid
Dim objGridCategoria As AdmGrid
Dim iGrid_Categoria_Col As Integer
Dim iGrid_Valor_Col As Integer
'******************************

'Alteracao Daniel
Private WithEvents objEventoContaDespesa As AdmEvento
Attribute objEventoContaDespesa.VB_VarHelpID = -1
Private WithEvents objEventoNatureza As AdmEvento
Attribute objEventoNatureza.VB_VarHelpID = -1

Private WithEvents objEventoTipoFornecedor As AdmEvento
Attribute objEventoTipoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoCondicaoPagto As AdmEvento
Attribute objEventoCondicaoPagto.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1
Private WithEvents objEventoContaFornConsig As AdmEvento
Attribute objEventoContaFornConsig.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_Complemento = 2
Private Const TAB_DadosFinanceiros = 3
Private Const TAB_Incricoes = 4
Private Const TAB_Endereco = 5
Private Const TAB_Estatisticas = 6

Private Sub Ativo_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LabelNomeReduzido_Click()

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim colSelecao As New Collection

On Error GoTo Erro_LabelNomeReduzido_Click

    'Verifica se o fornecedor foi preenchido
    If Len(NomeReduzido.ClipText) <> 0 Then

        'Preenche o código de objFornecedor
        objFornecedor.sNomeReduzido = NomeReduzido.Text
        
    End If

    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

    Exit Sub

Erro_LabelNomeReduzido_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160575)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Gera código automático do proximo Fornecedor
    lErro = CF("Fornecedor_Automatico", lCodigo)
    If lErro <> SUCESSO Then Error 55708

    'Exibe código automático na Tela
    Codigo.Text = CStr(lCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 55708
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160576)
    
    End Select

    Exit Sub

End Sub

Private Sub Agencia_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Banco_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Banco_GotFocus()

    Call MaskEdBox_TrataGotFocus(Banco, iAlterado)
    
End Sub

Private Sub Banco_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Banco_Validate

    'Verifica se foi preenchido o campo Banco
    If Len(Trim(Banco.Text)) = 0 Then Exit Sub

    'Critica se é do tipo positivo
    lErro = Valor_Positivo_Critica(Banco.Text)
    If lErro <> SUCESSO Then Error 43585

    Exit Sub

Erro_Banco_Validate:

    Cancel = True


    Select Case Err

        Case 43585
          
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160577)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
    
End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica se foi preenchido o campo Codigo Fornecedor
    If Len(Trim(Codigo.Text)) = 0 Then Exit Sub

    'Critica se é do tipo positivo
    lErro = Valor_Positivo_Critica(Codigo.Text)
    If lErro <> SUCESSO Then Error 19301

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case Err

        Case 19301 'Tratado na rotina chamada
          
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160578)

    End Select

    Exit Sub

End Sub

Private Sub ComboCategoriaFornecedor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ComboCategoriaFornecedor_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ComboCategoriaFornecedor_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridCategoria)
End Sub

Private Sub ComboCategoriaFornecedor_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCategoria)
End Sub

Private Sub ComboCategoriaFornecedor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCategoria.objControle = ComboCategoriaFornecedor
    lErro = Grid_Campo_Libera_Foco(objGridCategoria)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ComboCategoriaFornecedorItem_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ComboCategoriaFornecedorItem_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ComboCategoriaFornecedorItem_GotFocus()

Dim lErro As Long

On Error GoTo Erro_ComboCategoriaFornecedorItem_GotFocus

    'Preenche com os ítens relacionados a Categoria correspondente
    Call Trata_ComboCategoriaFornecedorItem

    Call Grid_Campo_Recebe_Foco(objGridCategoria)

    Exit Sub

Erro_ComboCategoriaFornecedorItem_GotFocus:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160579)

    End Select

    Exit Sub

End Sub

Private Function Trata_ComboCategoriaFornecedorItem()

Dim lErro As Long
Dim objCategoriaFornecedor As New ClassCategoriaCliente
Dim iIndice As Integer, sValor As String

On Error GoTo Erro_Trata_ComboCategoriaFornecedorItem

    'Armazena o Item em sValor
    sValor = ComboCategoriaFornecedorItem.Text

    'Limpa a Combo
    ComboCategoriaFornecedorItem.Clear

    'Retorna o Item para a Combo
    ComboCategoriaFornecedorItem.Text = sValor

    'Se alguém estiver selecionado
    If Len(Trim(GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Categoria_Col))) > 0 Then

        'Preencher a Combo de Itens desta Categoria
        objCategoriaFornecedor.sCategoria = GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Categoria_Col)

        lErro = Carrega_ComboCategoriaFornecedorItem()
        If lErro <> SUCESSO Then gError 91169

    End If

    'Seleciona o Item na List da Combo
    For iIndice = 0 To ComboCategoriaFornecedorItem.ListCount - 1
        If ComboCategoriaFornecedorItem.List(iIndice) = GridCategoria.Text Then
            ComboCategoriaFornecedorItem.ListIndex = iIndice
            Exit For
        End If
    Next

    Trata_ComboCategoriaFornecedorItem = SUCESSO

    Exit Function

Erro_Trata_ComboCategoriaFornecedorItem:

    Trata_ComboCategoriaFornecedorItem = gErr

    Select Case gErr

        Case 91169

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160580)

    End Select

End Function

Private Function Carrega_ComboCategoriaFornecedorItem() As Long
'Carrega a Combo CategoriaFornecedorItem

Dim lErro As Long
Dim colItensCategoria As New Collection
Dim objCategoriaFornecedorItem As New ClassCategoriaFornItem

On Error GoTo Erro_Carrega_ComboCategoriaFornecedorItem

    'Guarda no obj a categoria que deverá ter os itens lidos
    objCategoriaFornecedorItem.sCategoria = GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Categoria_Col)
    
    'Lê a tabela CategoriaProdutoItem a partir da Categoria
    lErro = CF("CategoriaFornecedor_Le_Itens", objCategoriaFornecedorItem, colItensCategoria)
    If lErro <> SUCESSO And lErro <> 91180 Then gError 91170

    'Insere na combo CategoriaFornecedorItem
    For Each objCategoriaFornecedorItem In colItensCategoria

        'Insere na combo CategoriaFornecedor
        ComboCategoriaFornecedorItem.AddItem objCategoriaFornecedorItem.sItem

    Next

    Carrega_ComboCategoriaFornecedorItem = SUCESSO

    Exit Function

Erro_Carrega_ComboCategoriaFornecedorItem:

    Carrega_ComboCategoriaFornecedorItem = gErr

    Select Case gErr

        Case 91170

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160581)

    End Select

    Exit Function

End Function

Private Sub ComboCategoriaFornecedorItem_KeyPress(KeyAscii As Integer)
Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCategoria)
End Sub

Private Sub ComboCategoriaFornecedorItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCategoria.objControle = ComboCategoriaFornecedorItem
    lErro = Grid_Campo_Libera_Foco(objGridCategoria)
    If lErro <> SUCESSO Then Cancel = True


End Sub

Private Sub CondicaoPagto_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaCorrente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Filiais_Click()

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filiais_Click

    'Verifica se foi preenchido o Codigo
    If Len(Trim(Codigo.Text)) = 0 Then Error 12862
        
    'Preenche objFornecedor
    objFornecedor.lCodigo = StrParaLong(Codigo.Text)
    'Lê o Fornecedor
    lErro = CF("Fornecedor_Le", objFornecedor)
    If lErro <> SUCESSO And lErro <> 12729 Then Error 43452
    
    'Se não achou o Fornecedor
    If lErro <> SUCESSO Then
            
            'Envia aviso perguntando se deseja cadastrar novo Fornecedor
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FORNECEDOR")
    
            If vbMsgRes = vbYes Then
                lErro = Gravar_Registro()
                If lErro <> SUCESSO Then Error 43453
                'Chama a Tela de Filiais de Fornecedor
                objFilialFornecedor.lCodFornecedor = CLng(Codigo.Text)
                objFilialFornecedor.iCodFilial = 1 'p/começar exibindo a matriz
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If
    Else
        'Chama a Tela de Filiais de Fornecedor
        objFilialFornecedor.lCodFornecedor = CLng(Codigo.Text)
        objFilialFornecedor.iCodFilial = 1 'p/começar exibindo a matriz
        Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
    End If
    
    Exit Sub

Erro_Filiais_Click:

    Select Case Err

        Case 12862
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODFORNECEDOR_NAO_PREENCHIDO", Err)

        Case 43452, 43453
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160582)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colCodigoNome As New AdmCollCodigoNome
Dim objCodigoNome As AdmlCodigoNome
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome
Dim colCodigo As New Collection
Dim vCodigo As Variant
Dim iIndice As Integer
Dim sMascaraConta As String
Dim iTipoPadrao As Integer
Dim objTela As Object

On Error GoTo Erro_Form_Load

    'Implementado pois agora é possível ter constantes cutomizadas em função de tamanhos de campos do BD. AdmLib.ClassConsCust
    RazaoSocial.MaxLength = STRING_FORNECEDOR_RAZAO_SOC
    NomeReduzido.MaxLength = STRING_FORNECEDOR_NOME_REDUZIDO
    
    
   Set objGridCategoria = New AdmGrid

    'Inicializa o Grid de Categoria
    lErro = Inicializa_Grid_Categoria(objGridCategoria)
    If lErro <> SUCESSO Then gError 91171

    'Carrega a combobox de Categoria Cliente
    lErro = Carrega_ComboCategoriaFornecedor()
    If lErro <> SUCESSO And lErro <> 91176 Then gError 91172
    
    If giTipoVersao = VERSAO_LIGHT Then
    
        LabelObservacao.top = TipoFornecedorLabel.top
        TipoFornecedorLabel.left = -20000
        TipoFornecedorLabel.Visible = False
        
        Observacao.top = Tipo.top
        Tipo.left = -20000
        Tipo.Visible = False
    
    End If
    
    iFrameAtual = 1

    Set objEventoTipoFornecedor = New AdmEvento
    Set objEventoCondicaoPagto = New AdmEvento
    Set objEventoFornecedor = New AdmEvento
    Set objEventoNatureza = New AdmEvento

    'Lê cada código e descrição da tabela TiposDeFornecedor
    lErro = CF("Cod_Nomes_Le", "TiposDeFornecedor", "Codigo", "Descricao", STRING_TIPO_FORNECEDOR_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 12718

    'Preenche a ComboBox Tipo com os objetos da coleção colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao

        Tipo.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        Tipo.ItemData(Tipo.NewIndex) = objCodigoDescricao.iCodigo

    Next
    
    lErro = CF("Carrega_CondicaoPagamento", CondicaoPagto, MODULO_CONTASAPAGAR)
    If lErro <> SUCESSO Then gError 12719


'    Set colCodigoDescricao = New AdmColCodigoNome
'
'    'Lê cada código e descrição reduzida da tabela CondicoesPagto
'    lErro = CF("CondicoesPagto_Le_Pagamento", colCodigoDescricao)
'    If lErro <> SUCESSO Then gError 12719
'
'    'Preenche a ComboBox CondicaoPagto com os objetos da coleção colCodigoDescricao
'    For Each objCodigoDescricao In colCodigoDescricao
'
'        CondicaoPagto.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
'        CondicaoPagto.ItemData(CondicaoPagto.NewIndex) = objCodigoDescricao.iCodigo
'
'    Next

    Set objTela = Me
    lErro = gobjTabEnd.Inicializa(objTela, TabEnd(0))
    If lErro <> SUCESSO Then gError 12720

    'Verifica se o modulo de contabilidade esta ativo antes das inicializacoes
    If (gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO) Then
        
        'Conta de Fornecedores
        Set objEventoContaContabil = New AdmEvento
        
        'Inicializa propriedade Mask de ContaContabil
        lErro = MascaraConta(sMascaraConta)
        If lErro <> SUCESSO Then gError 12722
    
        ContaContabil.Mask = sMascaraConta
        
        'Conta de Fornecedores em consignação
        Set objEventoContaFornConsig = New AdmEvento
        
        'Inicializa propriedade Mask de ContaContabil
        lErro = MascaraConta(sMascaraConta)
        If lErro <> SUCESSO Then gError 12722
    
        ContaFornConsig.Mask = sMascaraConta
    
        'Alteracao Daniel
        'Conta de Estoque/Despesa
        Set objEventoContaDespesa = New AdmEvento
        
        'Inicializa propriedade Mask de ContaDespesa
        lErro = MascaraConta(sMascaraConta)
        If lErro <> SUCESSO Then gError 106003
    
        ContaDespesa.Mask = sMascaraConta
        'Fim da Alteracao Daniel
    
    Else
    
       'Incluido a inicialização da máscara para não dar erro na gravação de clientes com conta mas que o módulo de contabilidade foi desabilitado
        lErro = MascaraConta(sMascaraConta)
        If lErro <> SUCESSO Then Error 12722
    
        ContaContabil.Mask = sMascaraConta
        ContaFornConsig.Mask = sMascaraConta
        ContaDespesa.Mask = sMascaraConta

        ContaFornConsig.Enabled = False
        ContaFornConsig.Enabled = False
        
        'Alteracao Daniel
        ContaDespesa.Enabled = False
        LabelContaDespesa.Enabled = False
        'Fim da Alteracao Daniel
        
    End If
    
    'Alteracao Daniel
    'Carrega a List da Combo de Tipos de Cobrança
    lErro = Carrega_TipoCobranca()
    If lErro <> SUCESSO Then gError 106006
    
    Call CF("Fornecedor_Le_Tipo_Padrao", iTipoPadrao)
    
    If iTipoPadrao <> 0 Then
        Tipo.Text = CStr(iTipoPadrao)
        Call Tipo_Validate(bSGECancelDummy)
    End If
    
    'Carrega a Combo TipoFrete
    lErro = Carrega_TipoFrete()
    If lErro <> SUCESSO Then gError 207473
    
    'Prepara as Combos  do Tab de Inscricao
    lErro = Inicializa_Tab_Inscricao()
    If lErro <> SUCESSO Then gError 207398
    
    'Inicializa a mascara de Natureza
    lErro = Inicializa_Mascara_Natureza()
    If lErro <> SUCESSO Then gError 12718
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 12718, 12719, 12720, 12721, 12722, 61394, 91171, 91171, 106003, 106006, 207398, 207473

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160583)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

'******** Janaina ********

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        Select Case GridCategoria.Col

            Case GRID_CATEGORIA_COL

                'Chama SaidaCelula de Categoria
                lErro = Saida_Celula_Categoria(objGridInt)
                If lErro <> SUCESSO Then gError 91173

            Case GRID_VALOR_COL

                'Chama SaidaCelula de Valor
                lErro = Saida_Celula_Valor(objGridInt)
                If lErro <> SUCESSO Then gError 91174

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 91175

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 91173, 91174

        Case 91175
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160585)

    End Select

    Exit Function

End Function

Private Function Carrega_ComboCategoriaFornecedor() As Long
'Carrega as Categorias na Combobox

Dim lErro As Long
Dim colCategorias As New Collection
Dim objCategoriaFornecedor As New ClassCategoriaFornecedor

On Error GoTo Erro_Carrega_ComboCategoriaFornecedor

  
    'Lê o código e a descrição de todas as categorias
    lErro = CF("CategoriaFornecedor_Le_Todos", colCategorias)
    If lErro <> SUCESSO And lErro <> 64686 Then gError 91176

    For Each objCategoriaFornecedor In colCategorias

        'Insere na combo CategoriaCliente
        ComboCategoriaFornecedor.AddItem objCategoriaFornecedor.sCategoria

    Next

    Carrega_ComboCategoriaFornecedor = SUCESSO

    Exit Function

Erro_Carrega_ComboCategoriaFornecedor:

    Carrega_ComboCategoriaFornecedor = gErr

    Select Case gErr

        Case 91176

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160586)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Categoria(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Categoria do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim colCategorias As New Collection
Dim objCategoriaFornecedor As New ClassCategoriaFornecedor

On Error GoTo Erro_Saida_Celula_Categoria

    Set objGridInt.objControle = ComboCategoriaFornecedor

    'Testa se .Text da combo é igual ao TextMatrix "embaixo"
    iIndice = GridCategoria.Row
    
    If UCase(GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL)) <> UCase(ComboCategoriaFornecedor.Text) Then

        If Len(Trim(ComboCategoriaFornecedor)) > 0 Then

            'Preenche o objeto com a Categoria
            objCategoriaFornecedor.sCategoria = ComboCategoriaFornecedor.Text

            'Lê Categoria De Cliente no BD
            lErro = CF("CategoriaFornecedor_Le", objCategoriaFornecedor)
            If lErro <> SUCESSO And lErro <> 90592 Then gError 91182

            'Categoria não está cadastrada
            If lErro = 90592 Then gError 91183

            'Verifica se já existe a categoria no Grid
            If GridCategoria.Row > 0 Then

                iIndice1 = GridCategoria.Row

                For iIndice = 1 To objGridCategoria.iLinhasExistentes

                    If iIndice <> iIndice1 Then

                        'Se já existe a categoria, erro
                        If UCase(GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL)) = UCase(ComboCategoriaFornecedor.Text) Then gError 91184

                    End If

                Next

            End If

            If GridCategoria.Row > 0 Then

                iIndice1 = GridCategoria.Row

                'Limpa o valor (ítem) da categoria
                GridCategoria.TextMatrix(iIndice1, GRID_VALOR_COL) = ""

            End If

        Else

             If GridCategoria.Row > 0 Then

                iIndice1 = GridCategoria.Row

                'Categoria vazia, limpa o Item correspondente
                GridCategoria.TextMatrix(iIndice1, GRID_VALOR_COL) = ""

            End If

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 91185

    If (GridCategoria.Row - GridCategoria.FixedRows = objGridInt.iLinhasExistentes) And Len(Trim(ComboCategoriaFornecedor.Text)) > 0 Then

        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1

    End If

    Saida_Celula_Categoria = SUCESSO

    Exit Function

Erro_Saida_Celula_Categoria:

    Saida_Celula_Categoria = gErr

    Select Case gErr

        Case 91182
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 91183 'Categoria não está cadastrada

            'Perguntar se deseja criar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DESEJA_CRIAR_CATEGORIAFORNECEDOR")

            If vbMsgRes = vbYes Then

                'Chama a Tela "CategoriaFornecedor"
                Call Chama_Tela("CategoriaFornec", objCategoriaFornecedor)
                
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            Else

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 91184
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_JA_SELECIONADA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 91185
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160587)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Valor(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Item do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objCategoriaFornecedor As New ClassCategoriaFornecedor
Dim objCategoriaFornecedorItem As New ClassCategoriaFornItem
Dim colItens As New Collection
Dim colItensCategoria As New Collection

Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridCategoria.objControle = ComboCategoriaFornecedorItem

    If Len(Trim(ComboCategoriaFornecedorItem)) > 0 Then

        'se o campo de categoria estiver vazio ==> erro
        If Len(GridCategoria.TextMatrix(GridCategoria.Row, GRID_CATEGORIA_COL)) = 0 Then gError 91186

        'Tenta selecionar na combo
        lErro = Combo_Item_Igual_CI(ComboCategoriaFornecedorItem)
        If lErro <> SUCESSO Then

            'Preenche o objeto com a Categoria
            objCategoriaFornecedorItem.sCategoria = GridCategoria.TextMatrix(GridCategoria.Row, GRID_CATEGORIA_COL)
            objCategoriaFornecedorItem.sItem = ComboCategoriaFornecedorItem.Text

            'Alterado por Luiz Nogueira em 03/03/04
            'Lê Item De Categoria De Fornecedor no BD
            lErro = CF("CategoriaFornecedorItem_Le", objCategoriaFornecedorItem, colItensCategoria)
            If lErro <> SUCESSO And lErro <> 127089 Then Error 91187

            'Item da Categoria não está cadastrado
            If lErro = 127089 Then gError 91188

        End If

        If GridCategoria.Row - GridCategoria.FixedRows = objGridInt.iLinhasExistentes Then

            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 91189

    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = gErr

    Select Case gErr

        Case 91187
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 91188 'Item da Categoria não está cadastrado

            'Perguntar se deseja criar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DESEJA_CRIAR_CATEGORIAFORNECEDORITEM", objCategoriaFornecedorItem.sItem)

            If vbMsgRes = vbYes Then

                'Preenche o objeto com a Categoria
                objCategoriaFornecedor.sCategoria = GridCategoria.TextMatrix(GridCategoria.Row, GRID_CATEGORIA_COL)

                'Chama a Tela "CategoriaFornecedor"
                Call Chama_Tela("CategoriaFornecedor", objCategoriaFornecedor, objCategoriaFornecedorItem.sItem)
                
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            Else

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 91189
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 91186
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_CATEGORIA_NAO_PREENCHIDA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160588)

    End Select

    Exit Function

End Function

Public Sub GridCategoria_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridCategoria, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCategoria, iAlterado)
    End If

End Sub

Public Sub GridCategoria_EnterCell()
    Call Grid_Entrada_Celula(objGridCategoria, iAlterado)
End Sub

Public Sub GridCategoria_GotFocus()
    Call Grid_Recebe_Foco(objGridCategoria)
End Sub

Public Sub GridCategoria_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridCategoria)
End Sub

Public Sub GridCategoria_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCategoria, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCategoria, iAlterado)
    End If

End Sub

Public Sub GridCategoria_LeaveCell()
    Call Saida_Celula(objGridCategoria)
End Sub

Public Sub GridCategoria_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridCategoria)
End Sub

Public Sub GridCategoria_RowColChange()
    Call Grid_RowColChange(objGridCategoria)
End Sub

Public Sub GridCategoria_Scroll()
    Call Grid_Scroll(objGridCategoria)
End Sub

Private Function Inicializa_Grid_Categoria(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Categoria")
    objGridInt.colColuna.Add ("Item")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ComboCategoriaFornecedor.Name)
    objGridInt.colCampo.Add (ComboCategoriaFornecedorItem.Name)

    'Colunas do Grid
    iGrid_Categoria_Col = 1
    iGrid_Valor_Col = 2

    'Grid do GridInterno
    objGridInt.objGrid = GridCategoria

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 21

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridCategoria.ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Categoria = SUCESSO

    Exit Function

End Function

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o fornecedor foi preenchido
    If Len(Codigo.ClipText) <> 0 Then

        'Preenche o código de objFornecedor
        objFornecedor.lCodigo = StrParaLong(Codigo.Text)
        
    End If

    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160589)

    End Select

    Exit Sub

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objFornecedor As ClassFornecedor
Dim objFornecedorEstatistica As New ClassFilialFornecedorEst

On Error GoTo Erro_objEventoFornecedor_evSelecao

    Set objFornecedor = obj1
    
    lErro = CF("Fornecedor_Le_Estendida", objFornecedor, objFornecedorEstatistica)
    If lErro <> SUCESSO And lErro <> 52701 Then gError 82678

    'Verifica se Fornecedor não está cadastrado
    If lErro <> SUCESSO Then gError 82679

    lErro = Exibe_Dados_Fornecedor(objFornecedor, objFornecedorEstatistica)
    If lErro <> SUCESSO Then gError 82680

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Me.Show
    
    Exit Sub

Erro_objEventoFornecedor_evSelecao:

    Select Case gErr

    Case 82678, 82680

    Case 82679
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INEXISTENTE", gErr)

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160590)

    End Select

    Exit Sub

    

End Sub

Private Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub IdEstrangeiro_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Observacao2_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub RazaoSocial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Tipo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Tipo_Click()

Dim lErro As Long
Dim objTipoFornecedor As New ClassTipoFornecedor
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Tipo_Click

    iAlterado = REGISTRO_ALTERADO
    
    'Se não foi selecionado Tipo, sai
    If Tipo.ListIndex = -1 Then Exit Sub
    
    objTipoFornecedor.iCodigo = Tipo.ItemData(Tipo.ListIndex)

    'Lê o Tipo de Fornecedor no BD
    lErro = CF("TipoFornecedor_Le", objTipoFornecedor)
    If lErro <> SUCESSO And lErro <> 12765 Then Error 12759
    
    'Tipo não está cadastrado
    If lErro <> SUCESSO Then Error 6997
    
    'Exibe na tela o Tipo lido no BD
    lErro = Exibe_Dados_TipoFornecedor(objTipoFornecedor)
    If lErro <> SUCESSO Then Error 12760

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_Tipo_Click:

    Select Case Err

        Case 6997
            'Envia aviso perguntando se deseja cadastrar novo Tipo de Fornecedor
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TIPOFORNECEDOR")
    
            If vbMsgRes = vbYes Then
                Call Chama_Tela("TipoFornecedor", objTipoFornecedor)
            Else
                Tipo.SetFocus
            End If
        
        Case 12759, 12760

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160592)

    End Select

    Exit Sub

End Sub

Private Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTipoFornecedor As New ClassTipoFornecedor
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_Tipo_Validate

    'Verifica se foi preenchida a ComboBox Tipo
    If Len(Trim(Tipo.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o item selecionado na ComboBox Tipo
    If Tipo.Text = Tipo.List(Tipo.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Tipo, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 19283

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTipoFornecedor.iCodigo = iCodigo

        'Tenta ler TipoFornecedor com esse código no BD
        lErro = CF("TipoFornecedor_Le", objTipoFornecedor)
        If lErro <> SUCESSO And lErro <> 12765 Then Error 19284
        If lErro <> SUCESSO Then Error 19285 'Não encontrou Tipo Fornecedor no BD

        'Exibe dados de TipoForncedor na Tela
        Tipo.Text = CStr(iCodigo) & SEPARADOR & objTipoFornecedor.sDescricao
        
        lErro = Exibe_Dados_TipoFornecedor(objTipoFornecedor)
        If lErro <> SUCESSO Then Error 6998

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 19286

    Exit Sub

Erro_Tipo_Validate:

    Cancel = True

    Select Case Err

        Case 6998, 19283, 19284 'Tratados na rotina chamada
    
        Case 19285  'Não encontrou Tipo Fornecedor no BD
    
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TIPOFORNECEDOR")
    
            If vbMsgRes = vbYes Then
                Call Chama_Tela("TipoFornecedor", objTipoFornecedor)
            Else
                'Segura o foco
            End If
    
        Case 19286
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_FORNECEDOR_NAO_ENCONTRADO", Err, Tipo.Text)
    
        Case Else
    
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160593)

    End Select

    Exit Sub

End Sub

Private Sub CondicaoPagto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CondicaoPagto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim iCodigo As Integer

On Error GoTo Erro_CondicaoPagto_Validate

    'Verifica se foi preenchida a ComboBox CondicaoPagto
    If Len(Trim(CondicaoPagto.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox CondicaoPagto
    If CondicaoPagto.Text = CondicaoPagto.List(CondicaoPagto.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(CondicaoPagto, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 33548

    'Não existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objCondicaoPagto.iCodigo = iCodigo

        'Tenta ler CondicaoPagto com esse código no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then Error 33549
        
        If lErro <> SUCESSO Then Error 33550 'Não encontrou CondicaoPagto no BD

        'Encontrou CondicaoPagto no BD e não é de Pagamento
        If objCondicaoPagto.iEmPagamento = 0 Then Error 33551

        'Coloca no Text da Combo
        CondicaoPagto.Text = CondPagto_Traz(objCondicaoPagto)

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 33552

    Exit Sub

Erro_CondicaoPagto_Validate:

    Cancel = True

    Select Case Err

        Case 33551
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_RECEBIMENTO", Err, iCodigo)
    
        Case 33548, 33549
    
        Case 33550  'Não encontrou CondicaoPagto no BD
    
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAO_PAGAMENTO")
            If vbMsgRes = vbYes Then Call Chama_Tela("CondicoesPagto", objCondicaoPagto)
    
        Case 33551
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_RECEBIMENTO", Err, iCodigo)
    
        Case 33552
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", Err, CondicaoPagto.Text)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160594)

    End Select

    Exit Sub

End Sub

Private Sub Desconto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Desconto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sDesconto As String

On Error GoTo Erro_Desconto_Validate

    sDesconto = Desconto.Text

    'Verifica se foi preenchido o Desconto
    If Len(Trim(Desconto.Text)) = 0 Then Exit Sub

    'Critica se é porcentagem
    lErro = Porcentagem_Critica(Desconto.Text)
    If lErro <> SUCESSO Then Error 12776

    Desconto.Text = Format(sDesconto, "Fixed")

    Exit Sub

Erro_Desconto_Validate:

    Cancel = True

    Select Case Err

        Case 12776
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160595)
    
        End Select

    Exit Sub

End Sub

Private Sub ContaContabil_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaFornConsig_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaContabil_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaMascarada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ContaContabil_Validate

    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica_Modulo", ContaContabil.Text, ContaContabil.ClipText, objPlanoConta, MODULO_CONTASAPAGAR)
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then Error 90540
    
    If lErro = SUCESSO Then
    
        sContaFormatada = objPlanoConta.sConta
        
        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)
        
        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 39798
        
        ContaContabil.PromptInclude = False
        ContaContabil.Text = sContaMascarada
        ContaContabil.PromptInclude = True
    
    'se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then
    
        lErro = CF("Conta_Critica", ContaContabil.Text, sContaFormatada, objPlanoConta, MODULO_CONTASAPAGAR)
        If lErro <> SUCESSO And lErro <> 5700 Then Error 12777
        
        If lErro = 5700 Then Error 12778

    End If
    
    Exit Sub

Erro_ContaContabil_Validate:

    Cancel = True

    Select Case Err

    Case 12777, 39797
        
    Case 12778
        lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", Err, ContaContabil.Text)

    Case 39798
        lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160596)

    End Select

    Exit Sub

End Sub

Private Sub ContaFornConsig_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaMascarada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ContaFornConsig_Validate

    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica_Modulo", ContaFornConsig.Text, ContaFornConsig.ClipText, objPlanoConta, MODULO_CONTASAPAGAR)
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 39797
    
    If lErro = SUCESSO Then
    
        sContaFormatada = objPlanoConta.sConta
        
        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)
        
        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then gError 90541
        
        ContaFornConsig.PromptInclude = False
        ContaFornConsig.Text = sContaMascarada
        ContaFornConsig.PromptInclude = True
    
    'se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then
    
        lErro = CF("Conta_Critica", ContaFornConsig.Text, sContaFormatada, objPlanoConta, MODULO_CONTASAPAGAR)
        If lErro <> SUCESSO And lErro <> 5700 Then gError 90542
        
        If lErro = 5700 Then gError 90543

    End If
    
    Exit Sub

Erro_ContaFornConsig_Validate:

    Cancel = True


    Select Case gErr

    Case 90542, 90540
        
    Case 90543
        Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", gErr, ContaFornConsig.Text)

    Case 90541
        Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160597)

    End Select

    Exit Sub

End Sub

Private Sub CGC_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CGC_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CGC, iAlterado)
    
End Sub

Private Sub CGC_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CGC_Validate

    If Len(Trim(CGC.Text)) = 0 Then Exit Sub

    Select Case Len(Trim(CGC.Text))

        Case STRING_CPF 'CPF
    
            lErro = Cpf_Critica(CGC.Text)
            If lErro <> SUCESSO Then Error 12780
    
            CGC.Format = "000\.000\.000-00; ; ; "
            CGC.Text = CGC.Text
    
        Case STRING_CGC  'CGC
    
            lErro = Cgc_Critica(CGC.Text)
            If lErro <> SUCESSO Then Error 12781
    
            CGC.Format = "00\.000\.000\/0000-00; ; ; "
            CGC.Text = CGC.Text
    
    Case Else

        Error 12782

    End Select

    Exit Sub

Erro_CGC_Validate:

    Cancel = True

    Select Case Err

        Case 12780, 12781

        Case 12782
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC_CPF", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160598)

    End Select

    Exit Sub

End Sub

Private Sub InscricaoEstadual_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub InscricaoMunicipal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub


Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava o Fornecedor
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 47890

    'Limpa a Tela
    lErro = Limpa_Tela_Fornecedores
    If lErro <> SUCESSO Then gError 78017

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 47890, 78017

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160601)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 12858

    'Limpa a Tela
    lErro = Limpa_Tela_Fornecedores
    If lErro <> SUCESSO Then gError 78018

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 12858, 78018

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160602)

    End Select

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim colCodNomeFiliais As New AdmColCodigoNome
Dim objFornecedor As New ClassFornecedor
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim lCodigo As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    Parent.MousePointer = vbHourglass
    
    'Verifica se o código foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 12822

    objFornecedor.lCodigo = CLng(Codigo.Text)

    'Lê os dados do Fornecedor a ser excluido
    lErro = CF("Fornecedor_Le", objFornecedor)
    If lErro <> SUCESSO And lErro <> 12729 Then gError 12823

    'Verifica se Fornecedor não está cadastrado
    If lErro <> SUCESSO Then gError 12824
    
    lErro = CF("Fornecedor_Valida", objFornecedor)
    If lErro <> SUCESSO Then gError 12823

    'Envia aviso perguntando se realmente deseja excluir Fornecedor e suas filiais
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_FORNECEDOR", objFornecedor.lCodigo)

    If vbMsgRes = vbYes Then

        'Exclui Fornecedor
        lErro = CF("Fornecedor_Exclui", objFornecedor)
        If lErro <> SUCESSO Then gError 12826

        'Limpa a Tela
        lErro = Limpa_Tela_Fornecedores
        If lErro <> SUCESSO Then gError 78019
        
        iAlterado = 0

    End If

    Parent.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    Parent.MousePointer = vbDefault
    
    Select Case gErr

        Case 12822, 78019
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODFORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 12823, 12826, 91191

        Case 12824
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160603)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

 Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
   
End Sub

Function Limpa_Tela_Fornecedores() As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Tela_Fornecedores

    Call Limpa_Tela(Me)

    Tipo.Text = ""

    CondicaoPagto.Text = ""

    'Inicializa os campos cliente existentes nos frames
    For iIndice = 1 To 4
        Fornecedor(iIndice).Caption = ""
    Next
    
    'cyntia em 14/06/02
    'Limpa os labels
    SaldoTitulos.Caption = Format(0, "standard")
    NumeroCompras.Caption = Format(0, "standard")
    MediaCompra.Caption = Format(0, "standard")
    ValorAcumuladoCompras.Caption = Format(0, "standard")
    DataPrimeiraCompra.Caption = ""
    DataUltimaCompra.Caption = ""
    ValorAcumuladoDevolucoes.Caption = Format(0, "standard")
    NumTotalDevolucoes.Caption = 0
    DataUltDevolucao.Caption = ""
    PedidosEmAberto.Caption = 0
    MediaAtraso.Caption = 0
    MaiorAtraso.Caption = 0
    IEIsento.Value = vbChecked
    IENaoContrib.Value = vbChecked
    
    Ativo.Value = MARCADO
    
    '******** Janaina ********
    Call Grid_Limpa(objGridCategoria)
    '*************************
    
    Set gobjAnotacao = Nothing
    
    Call gobjTabEnd.Limpa_Tela
        
    Natureza.PromptInclude = False
    Natureza.Text = ""
    Natureza.PromptInclude = True
    LabelNaturezaDesc.Caption = ""
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Limpa_Tela_Fornecedores = SUCESSO
    
    Exit Function
    
Erro_Limpa_Tela_Fornecedores:

    Limpa_Tela_Fornecedores = gErr
    
    Select Case gErr
        
        Case 78015
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160604)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objFornecedor As New ClassFornecedor
Dim objCategoriaFornecedor As ClassCategoriaFornecedor
Dim objFornecedorList As New AdmlCodigoNome
Dim objEndereco As New ClassEndereco
Dim colItensCategoria As New Collection
Dim lCodigo As Long
Dim iAchou As Integer
Dim colEnderecos As New Collection

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se foi preenchido o Codigo
    If Len(Trim(Codigo.Text)) = 0 Then gError 12792

    'Verifica se foi preenchida a Razao Social
    If Len(Trim(RazaoSocial.Text)) = 0 Then gError 12793

    'Verifica se foi preenchido o Nome Reduzido
    If Len(Trim(NomeReduzido.Text)) = 0 Then gError 12794
    
    'Se o Tipo de Fornecedor não foi preenchido => erro
    If Len(Trim(Tipo.Text)) = 0 Then gError 62672

    'Lê os dados do Endereço
    lErro = gobjTabEnd.Move_Endereco_Memoria(colEnderecos)
    If lErro <> SUCESSO Then gError 12795
    
    Set objEndereco = colEnderecos.Item(1)

'****** Janaina ******

    'Verifica se para cada Categoria existente no grid de Categorias existe um Valor correspondente
    For iIndice = 1 To objGridCategoria.iLinhasExistentes
        'Verifica se a Categoria foi preenchida
         If Len(Trim(GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL))) <> 0 And Len(Trim(GridCategoria.TextMatrix(iIndice, GRID_VALOR_COL))) = 0 Then gError 91190
    Next
    
'*********************
    
    'Lê os dados da tela do Fornecedor
    lErro = Le_Dados_Fornecedor(objFornecedor)
    If lErro <> SUCESSO Then gError 12798
    
    lErro = CF("InscricaoEstadual_Critica", objFornecedor.sInscricaoEstadual, objEndereco.sSiglaEstado)
    If lErro <> SUCESSO Then gError 189577
    
    lErro = CF("Fornecedor_Valida", objFornecedor)
    If lErro <> SUCESSO Then gError 189577

    lErro = Trata_Alteracao(objFornecedor, objFornecedor.lCodigo)
    If lErro <> SUCESSO Then gError 80453

    'Se o CGC estiver Preenchido
    If Len(Trim(objFornecedor.sCgc)) > 0 Then
        'Verifica se tem outro fornecedor com o mesmo CGC e dá aviso
        lErro = CF("FilialFornecedor_Testa_CGC", objFornecedor.lCodigo, 0, objFornecedor.sCgc)
        If lErro <> SUCESSO Then gError 80217
    End If
    
    lErro = CF("Fornecedor_Grava", objFornecedor, objEndereco)
    If lErro <> SUCESSO Then gError 12799
    
    'gravar anotacao, se houver
    If Not (gobjAnotacao Is Nothing) Then
    
        If Len(Trim(gobjAnotacao.sTextoCompleto)) <> 0 Or Len(Trim(gobjAnotacao.sTitulo)) <> 0 Then
        
            gobjAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_FORNECEDOR
            gobjAnotacao.sID = CStr(objFornecedor.lCodigo)
            gobjAnotacao.dtDataAlteracao = gdtDataHoje
            
            lErro = CF("Anotacoes_Grava", gobjAnotacao)
            If lErro <> SUCESSO Then gError 128149
            
        End If
        
    End If
    
    iAchou = 0

    iAlterado = 0

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = gErr

    Select Case gErr

        Case 12792
            Call Rotina_Erro(vbOKOnly, "ERRO_CODFORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 12793
            Call Rotina_Erro(vbOKOnly, "ERRO_RAZ_SOC_NAO_PREENCHIDA", gErr)

        Case 12794
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_PREENCHIDO", gErr)

        Case 12795, 12798, 80217, 80453, 91191, 12799, 128149, 189577
        
        Case 62672
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TIPO_FORNECEDOR_NAO_PREENCHIDO", gErr)
            
        Case 91190
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAFORNECEDORITEM_ITEM_NAO_SELECIONADO1", gErr, iIndice)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160605)

    End Select

    Exit Function

End Function

Private Function Le_Dados_Fornecedor(objFornecedor As ClassFornecedor) As Long
'Lê os dados que estão na tela de Fornecedores e coloca em objFornecedor

Dim lErro As Long
Dim iContaPreenchida As Integer
Dim objFilialFornCategoria As ClassFilialFornCategoria
Dim iIndice As Integer
Dim sConta As String
Dim sContaFornConsig As String
Dim sNaturezaFormatada As String
Dim iNaturezaPreenchida As Integer

On Error GoTo Erro_Le_Dados_Fornecedor

    'IDENTIFICACAO :

    If Len(Trim(Codigo.Text)) > 0 Then objFornecedor.lCodigo = CLng(Codigo.Text)
    objFornecedor.sRazaoSocial = RazaoSocial.Text
    objFornecedor.sNomeReduzido = NomeReduzido.Text
    If Len(Trim(Tipo.Text)) > 0 Then objFornecedor.iTipo = Codigo_Extrai(Tipo.Text)
    objFornecedor.sObservacao = Observacao.Text
    objFornecedor.iAtivo = Ativo.Value

    sNaturezaFormatada = String(STRING_NATMOVCTA_CODIGO, 0)
    
    'Coloca no formato do BD
    lErro = CF("Item_Formata", SEGMENTO_NATMOVCTA, Natureza.Text, sNaturezaFormatada, iNaturezaPreenchida)
    If lErro <> SUCESSO Then Error 12800
    
    objFornecedor.sNatureza = sNaturezaFormatada

'********** Janaina ***************

    'Preenche uma coleção com todas as linhas "existentes" do grid de categorias
    For iIndice = 1 To objGridCategoria.iLinhasExistentes

        'Verifica se a Categoria foi preenchida
        If Len(Trim(GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL))) <> 0 Then

            Set objFilialFornCategoria = New ClassFilialFornCategoria
           
            objFilialFornCategoria.lFornecedor = objFornecedor.lCodigo
            objFilialFornCategoria.sCategoria = GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL)
            objFilialFornCategoria.sItem = GridCategoria.TextMatrix(iIndice, GRID_VALOR_COL)

            objFornecedor.colCategoriaItem.Add objFilialFornCategoria

        End If

    Next
    
'****************************
    'DADOS FINANCEIROS :

    If Len(Trim(Desconto.Text)) > 0 Then objFornecedor.dDesconto = CDbl(Desconto.Text) / 100

    objFornecedor.iCondicaoPagto = CondPagto_Extrai(CondicaoPagto)

    lErro = CF("Conta_Formata", ContaContabil.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 12800
    If iContaPreenchida = CONTA_VAZIA Then
        objFornecedor.sContaContabil = ""
    Else
        objFornecedor.sContaContabil = sConta
    End If
    
    lErro = CF("Conta_Formata", ContaFornConsig.Text, sContaFornConsig, iContaPreenchida)
    If lErro <> SUCESSO Then gError 90547
    
    If iContaPreenchida = CONTA_VAZIA Then
        objFornecedor.sContaFornConsig = ""
    Else
        objFornecedor.sContaFornConsig = sContaFornConsig
    End If
    
    objFornecedor.sContaCorrente = ContaCorrente.Text
    objFornecedor.sAgencia = Agencia.Text
    If Len(Trim(Banco.Text)) > 0 Then
       If CInt(Banco.Text) < 0 Then Error 12790
        objFornecedor.iBanco = CInt(Banco.Text)
    End If
    
    'INSCRICOES :

    objFornecedor.sCgc = CGC.Text
    objFornecedor.sInscricaoEstadual = InscricaoEstadual.Text
    objFornecedor.sInscricaoMunicipal = InscricaoMunicipal.Text
    objFornecedor.sInscricaoSuframa = Trim(InscricaoSuframa.ClipText)
    objFornecedor.sObservacao2 = Observacao2.Text
    objFornecedor.iRegimeTributario = RegimeTributario.ItemData(RegimeTributario.ListIndex)
    objFornecedor.sIdEstrangeiro = Trim(IdEstrangeiro.Text)

    If IEIsento.Value = vbChecked Then
        objFornecedor.iIEIsento = MARCADO
    Else
        objFornecedor.iIEIsento = DESMARCADO
    End If
    If IENaoContrib.Value = vbChecked Then
        objFornecedor.iIENaoContrib = MARCADO
    Else
        objFornecedor.iIENaoContrib = DESMARCADO
    End If
    
    'Alteracao Daniel
    objFornecedor.sInscricaoINSS = InscricaoINSS.Text
    objFornecedor.iGeraCredICMS = GeraICMS.Value
    
    If TipoFrete.ListIndex <> -1 Then
        objFornecedor.iTipoFrete = TipoFrete.ItemData(TipoFrete.ListIndex)
    Else
        objFornecedor.iTipoFrete = -1
    End If
    
    objFornecedor.iTipoCobranca = Codigo_Extrai(TipoCobranca.Text)
    

    'ESTATISTICAS :

    objFornecedor.dSaldoTitulos = CDbl(SaldoTitulos.Caption)
    objFornecedor.lNumeroCompras = CLng(NumeroCompras.Caption)
    objFornecedor.dMediaCompra = CDbl(MediaCompra.Caption)
    objFornecedor.dValorAcumuladoCompras = CDbl(ValorAcumuladoCompras.Caption)
    If Trim(DataPrimeiraCompra.Caption) = "" Then
        objFornecedor.dtDataPrimeiraCompra = DATA_NULA
    Else
        objFornecedor.dtDataPrimeiraCompra = StrParaDate(DataPrimeiraCompra.Caption)
    End If
    If Trim(DataUltimaCompra.Caption) = "" Then
        objFornecedor.dtDataUltimaCompra = DATA_NULA
    Else
        objFornecedor.dtDataUltimaCompra = StrParaDate(DataUltimaCompra.Caption)
    End If
    objFornecedor.sAgencia = Agencia.Text
    objFornecedor.lMediaAtraso = CLng(MediaAtraso.Caption)
    objFornecedor.lMaiorAtraso = CLng(MaiorAtraso.Caption)
    objFornecedor.dValorAcumuladoDevolucoes = CDbl(ValorAcumuladoDevolucoes.Caption)
    objFornecedor.lNumTotalDevolucoes = CLng(NumTotalDevolucoes.Caption)
    If Trim(DataUltDevolucao.Caption) = "" Then
        objFornecedor.dtDataUltDevolucao = DATA_NULA
    Else
        objFornecedor.dtDataUltDevolucao = StrParaDate(DataUltDevolucao.Caption)
    End If
    
    'Alteracao Daniel
    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaDespesa.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 12966

    If iContaPreenchida = CONTA_VAZIA Then
        objFornecedor.sContaDespesa = ""
    Else
        objFornecedor.sContaDespesa = sConta
    End If
    'Fim da Alteracao Daniel

    Le_Dados_Fornecedor = SUCESSO

    Exit Function

Erro_Le_Dados_Fornecedor:

    Le_Dados_Fornecedor = Err

    Select Case Err

        Case 12800, 19301, 90547
        
        Case 12790
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BANCO_NAO_POSITIVO", Err, Banco.Text)
            Banco.SetFocus
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160607)

    End Select

    Exit Function

End Function

Function Exibe_Dados_TipoFornecedor(objTipoFornecedor As ClassTipoFornecedor) As Long
'Exibe dados do Tipo de Fornecedor na tela

Dim lErro As Long
Dim sContaEnxuta As String

On Error GoTo Erro_Exibe_Dados_TipoFornecedor

    If objTipoFornecedor.iCondicaoPagto = 0 Then
        CondicaoPagto.Text = ""
    Else
        CondicaoPagto.Text = CStr(objTipoFornecedor.iCondicaoPagto)
        Call CondicaoPagto_Validate(bSGECancelDummy)
    End If

    If objTipoFornecedor.dDesconto = 0# Then
        Desconto.Text = ""
    Else
        Desconto.Text = CStr(100 * objTipoFornecedor.dDesconto)
    End If

    Observacao.Text = objTipoFornecedor.sObservacao

    ContaContabil.PromptInclude = False

    If objTipoFornecedor.sContaDespesa = "" Then
        ContaContabil.Text = ""
    Else
        lErro = Mascara_RetornaContaEnxuta(objTipoFornecedor.sContaDespesa, sContaEnxuta)
        If lErro <> SUCESSO Then Error 12766

        ContaContabil.Text = sContaEnxuta
    End If

    ContaContabil.PromptInclude = True

    Exibe_Dados_TipoFornecedor = SUCESSO

Exit Function

Erro_Exibe_Dados_TipoFornecedor:

    Exibe_Dados_TipoFornecedor = Err

    Select Case Err

        Case 12766
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objTipoFornecedor.sContaDespesa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160608)

    End Select

    Exit Function

End Function

Function Trata_Parametros(Optional objFornecedor As ClassFornecedor) As Long

Dim lErro As Long
Dim lCodigo As Long
Dim objFornecedorEstatistica As New ClassFilialFornecedorEst

On Error GoTo Erro_Trata_Parametros

    If Not (objFornecedor Is Nothing) Then

        lErro = CF("Fornecedor_Le_Estendida", objFornecedor, objFornecedorEstatistica)
        If lErro <> SUCESSO And lErro <> 52701 Then gError 12723

        If lErro = SUCESSO Then
            lErro = Exibe_Dados_Fornecedor(objFornecedor, objFornecedorEstatistica)
            If lErro <> SUCESSO Then gError 12724
        Else
            lErro = Limpa_Tela_Fornecedores()
            If lErro <> SUCESSO Then gError 78016
            If objFornecedor.lCodigo <> 0 Then
                Codigo.Text = objFornecedor.lCodigo
            End If
            If objFornecedor.sNomeReduzido <> "" Then NomeReduzido.Text = objFornecedor.sNomeReduzido
        End If
    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 12723, 12724, 12745, 78016

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160609)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Sub NomeReduzido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NomeReduzido_Validate(Cancel As Boolean)

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_NomeReduzido_Validate
    
    'Se está preenchido, testa se começa por letra
    If Len(Trim(NomeReduzido.Text)) > 0 Then
    
        If Not IniciaLetra(NomeReduzido.Text) Then Error 25004
        
    End If

    'Preenche os label Fornecedor existentes nos Frames
    For iIndice = 1 To 4
        Fornecedor(iIndice).Caption = NomeReduzido.Text
    Next
    
    Exit Sub

Erro_NomeReduzido_Validate:

    Cancel = True


    Select Case Err

        Case 25004
            
           lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_COMECA_LETRA", Err, NomeReduzido.Text)
    
        Case Else
    
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160610)

    End Select

End Sub

Private Sub Opcao_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index
            
        Select Case iFrameAtual
        
            Case TAB_Identificacao
                Parent.HelpContextID = IDH_FORNECEDOR_IDENT
            
            Case TAB_Complemento
                Parent.HelpContextID = IDH_FORNECEDOR_DADOS_FIN '??? Criar um IDH para esse tab

            Case TAB_DadosFinanceiros
                Parent.HelpContextID = IDH_FORNECEDOR_DADOS_FIN
                
            Case TAB_Incricoes
                Parent.HelpContextID = IDH_FORNECEDOR_INSCRICOES
            
            Case TAB_Endereco
                Parent.HelpContextID = IDH_FORNECEDOR_ENDERECO
            
            Case TAB_Estatisticas
                Parent.HelpContextID = IDH_FORNECEDOR_ESTATISTICAS
        
        End Select
    End If

End Sub

Function Exibe_Dados_Fornecedor(objFornecedor As ClassFornecedor, objFornecedorEstatistica As ClassFilialFornecedorEst) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iCodigo As Integer
Dim sContaEnxuta As String
Dim objEndereco As New ClassEndereco
Dim colCategoriaItem As New Collection
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objTipoFornecedor As New ClassTipoFornecedor
Dim colEnderecos As New Collection
Dim sNaturezaEnxuta As String

On Error GoTo Erro_Exibe_Dados_Fornecedor

    'pesquisa a primeira compra
    lErro = CF("Pesquisa_Primeira_Compra_Forn", objFornecedorEstatistica)
    If lErro <> SUCESSO Then gError 89402

    'pesquisa a ultima compra
    lErro = CF("Pesquisa_Ultima_Compra_Forn", objFornecedorEstatistica)
    If lErro <> SUCESSO Then gError 89403

    'pesquisa a ultima nota fiscal de Devolução de Devolução da FilialEmpresa (data de emissao mais recente) para o Fornecedor em questão.
    lErro = CF("Pesquisa_Ultima_Devolucao_Forn", objFornecedorEstatistica)
    If lErro <> SUCESSO Then gError 89415

    'IDENTIFICAÇÃO :
    Codigo.Text = CStr(objFornecedor.lCodigo)
    RazaoSocial.Text = objFornecedor.sRazaoSocial
    NomeReduzido.Text = objFornecedor.sNomeReduzido

    Ativo.Value = objFornecedor.iAtivo

    If Len(Trim(objFornecedor.sNatureza)) <> 0 Then
    
        sNaturezaEnxuta = String(STRING_NATMOVCTA_CODIGO, 0)
    
        lErro = Mascara_RetornaItemEnxuto(SEGMENTO_NATMOVCTA, objFornecedor.sNatureza, sNaturezaEnxuta)
        If lErro <> SUCESSO Then gError 89415
    
        Natureza.PromptInclude = False
        Natureza.Text = sNaturezaEnxuta
        Natureza.PromptInclude = True
        
    Else
    
        Natureza.PromptInclude = False
        Natureza.Text = ""
        Natureza.PromptInclude = True
        
    End If
    
    Call Natureza_Validate(bSGECancelDummy)
    
    If objFornecedor.iTipo = 0 Then
        Tipo.Text = ""
    Else
        Tipo.Text = CStr(objFornecedor.iTipo)
         
        'Tenta selecionar o Tipo na Combo
        lErro = Combo_Seleciona(Tipo, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 12355
                    
        If lErro <> SUCESSO Then
                
            objTipoFornecedor.iCodigo = objFornecedor.iTipo
            
            'Tenta ler TipoCliente com esse código no BD
            lErro = CF("TipoFornecedor_Le", objTipoFornecedor)
            If lErro <> SUCESSO And lErro <> 12765 Then gError 12773
            
            'Não encontrou Tipo Fornecedor no BD
            If lErro = 12765 Then gError 12319
            
            'Insere na Combo
            Tipo.AddItem CStr(objTipoFornecedor.iCodigo) & SEPARADOR & objTipoFornecedor.sDescricao
            Tipo.ItemData(Tipo.NewIndex) = objTipoFornecedor.iCodigo
            
            'Chama Combo_Seleciona de novo
            lErro = Combo_Seleciona(Tipo, iCodigo)
            If lErro <> SUCESSO Then gError 12320
            
        End If
        
    End If

    If objFornecedor.sObservacao <> "" Then
        Observacao.Text = objFornecedor.sObservacao
    Else
        Observacao.Text = ""
    End If
    
    'Preenche os campos Fornecedor existentes nos frames
    For iIndice = 1 To 4
        Fornecedor(iIndice).Caption = objFornecedor.sNomeReduzido
    Next

    'DADOS FINANCEIROS :

    If objFornecedor.dDesconto <> 0 Then
        Desconto.Text = CStr(100 * objFornecedor.dDesconto)
    Else
        Desconto.Text = ""
    End If
    
    If objFornecedor.iCondicaoPagto = 0 Then
        CondicaoPagto.Text = ""
    Else
        CondicaoPagto.Text = CStr(objFornecedor.iCondicaoPagto)
        Call CondicaoPagto_Validate(bSGECancelDummy)
    End If

    ContaContabil.PromptInclude = False
    If Len(Trim(objFornecedor.sContaContabil)) = 0 Then
        ContaContabil.Text = ""
    Else
        lErro = Mascara_RetornaContaEnxuta(objFornecedor.sContaContabil, sContaEnxuta)
        If lErro <> SUCESSO Then gError 12819

        ContaContabil.Text = sContaEnxuta
        
    End If
    ContaContabil.PromptInclude = True

    ContaFornConsig.PromptInclude = False
    If Len(Trim(objFornecedor.sContaFornConsig)) = 0 Then
        ContaFornConsig.Text = ""
    Else
        lErro = Mascara_RetornaContaEnxuta(objFornecedor.sContaFornConsig, sContaEnxuta)
        If lErro <> SUCESSO Then gError 90546

        ContaFornConsig.Text = sContaEnxuta
        
    End If
    ContaFornConsig.PromptInclude = True

    ContaCorrente.Text = objFornecedor.sContaCorrente

    Agencia.Text = objFornecedor.sAgencia

    If objFornecedor.iBanco = 0 Then
        Banco.Text = ""
    Else
        Banco.Text = CStr(objFornecedor.iBanco)
    End If

    'INSCRIÇÕES :

    CGC.Text = objFornecedor.sCgc
    Call CGC_Validate(bSGECancelDummy)
    InscricaoEstadual.Text = objFornecedor.sInscricaoEstadual
    Call Trata_IE
    If objFornecedor.iIEIsento = MARCADO Then
        IEIsento.Value = vbChecked
    Else
        IEIsento.Value = vbUnchecked
    End If
    If objFornecedor.iIENaoContrib = MARCADO Then
        IENaoContrib.Value = vbChecked
    Else
        IENaoContrib.Value = vbUnchecked
    End If
    InscricaoMunicipal.Text = objFornecedor.sInscricaoMunicipal
    Observacao2.Text = objFornecedor.sObservacao2
    InscricaoSuframa.Text = objFornecedor.sInscricaoSuframa
    IdEstrangeiro.Text = objFornecedor.sIdEstrangeiro

    'Busca no ItemData da Combo de Tipo de Cobranca o código
    For iIndice = 0 To RegimeTributario.ListCount - 1
        If RegimeTributario.ItemData(iIndice) = objFornecedor.iRegimeTributario Then
            RegimeTributario.ListIndex = iIndice
            Exit For
        End If
    Next


    'ENDEREÇO :

    objEndereco.lCodigo = objFornecedor.lEndereco
    'Lê os dados do endereço do Fornecedor
    lErro = CF("Endereco_Le", objEndereco)
    If lErro <> SUCESSO Then gError 12820

    colEnderecos.Add objEndereco

    lErro = gobjTabEnd.Traz_Endereco_Tela(colEnderecos)
    If lErro <> SUCESSO Then gError 12820

    'ESTATÍSTICAS :

    SaldoTitulos.Caption = Format(objFornecedorEstatistica.dSaldoTitulos, "Standard")
    NumeroCompras.Caption = Format(objFornecedorEstatistica.lNumeroCompras, FORMATO_INTEIRO)
    If objFornecedorEstatistica.lNumeroCompras > 0 Then
        MediaCompra.Caption = Format(objFornecedorEstatistica.dValorAcumuladoCompras / objFornecedorEstatistica.lNumeroCompras, "Standard")
    Else
        MediaCompra.Caption = CStr(0)
    End If
    ValorAcumuladoCompras.Caption = Format(objFornecedorEstatistica.dValorAcumuladoCompras, "Standard")
    
    If objFornecedorEstatistica.dtDataPrimeiraCompra = DATA_NULA Then
        DataPrimeiraCompra.Caption = ""
    Else
        DataPrimeiraCompra.Caption = Format(objFornecedorEstatistica.dtDataPrimeiraCompra, "dd/mm/yyyy")
    End If
    
    If objFornecedorEstatistica.dtDataUltimaCompra = DATA_NULA Then
        DataUltimaCompra.Caption = ""
    Else
        DataUltimaCompra.Caption = Format(objFornecedorEstatistica.dtDataUltimaCompra, "dd/mm/yyyy")
    End If

    If objFornecedorEstatistica.lItensPedidosRecebidos > 0 Then
        MediaAtraso.Caption = Format(objFornecedorEstatistica.lAtrasoAcumulado / objFornecedorEstatistica.lItensPedidosRecebidos, FORMATO_MEDIA_ATRASO)
    Else
        MediaAtraso.Caption = CStr(0)
    End If
    
'    MaiorAtraso.Caption = CStr(objFornecedorEstatistica.lMaiorAtraso)

    ValorAcumuladoDevolucoes.Caption = Format(objFornecedorEstatistica.dValorAcumuladoDevolucoes, "Standard")
    NumTotalDevolucoes.Caption = Format(objFornecedorEstatistica.lNumTotalDevolucoes, FORMATO_INTEIRO)

    If objFornecedorEstatistica.dtDataUltDevolucao = DATA_NULA Then
        DataUltDevolucao.Caption = ""
    Else
        DataUltDevolucao.Caption = Format(objFornecedorEstatistica.dtDataUltDevolucao, "dd/mm/yyyy")
    End If
    
    '********** Janaina ***********
    
   'Lê as categoria deste fornecedor
    objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
    objFilialFornecedor.iCodFilial = FILIAL_MATRIZ
    
    'Alterado em 03/03/04 por Luiz Nogueira
    lErro = CF("FilialFornecedorCategorias_Le_FornFilial", objFilialFornecedor, colCategoriaItem)
    If lErro <> SUCESSO And lErro <> 127085 Then gError 33010

    'Limpa o grid de categorias
    Call Grid_Limpa(objGridCategoria)
    
    'Carrega o Grid Categoria
    If colCategoriaItem.Count > 0 Then
        For iIndice = 1 To colCategoriaItem.Count
            GridCategoria.TextMatrix(iIndice, iGrid_Categoria_Col) = colCategoriaItem.Item(iIndice).sCategoria
            GridCategoria.TextMatrix(iIndice, iGrid_Valor_Col) = colCategoriaItem.Item(iIndice).sItem
        Next
        objGridCategoria.iLinhasExistentes = colCategoriaItem.Count
    End If
'******************************

    'Alteracao Daniel
    ContaDespesa.PromptInclude = False

    If Len(Trim(objFornecedor.sContaDespesa)) = 0 Then
        ContaDespesa.Text = objFornecedor.sContaDespesa
    Else
        lErro = Mascara_RetornaContaEnxuta(objFornecedor.sContaDespesa, sContaEnxuta)
        If lErro <> SUCESSO Then gError 106004

        ContaDespesa.Text = sContaEnxuta
    End If

    ContaDespesa.PromptInclude = True
    'Fim da Alteracao
    
    'Alteracao Daniel
    InscricaoINSS.Text = objFornecedor.sInscricaoINSS
    
    'Alteracao Daniel
    GeraICMS.Value = objFornecedor.iGeraCredICMS
    
    If objFornecedor.iTipoFrete <> -1 Then
        For iIndice = 0 To TipoFrete.ListCount - 1
            If TipoFrete.ItemData(iIndice) = objFornecedor.iTipoFrete Then
                TipoFrete.ListIndex = iIndice
                Exit For
            End If
        Next
    Else
        TipoFrete.ListIndex = -1
    End If
    
    'Busca no ItemData da Combo de Tipo de Cobranca o código
    For iIndice = 0 To TipoCobranca.ListCount - 1
        If TipoCobranca.ItemData(iIndice) = objFornecedor.iTipoCobranca Then
            TipoCobranca.ListIndex = iIndice
            Exit For
        End If
    Next
    'Fim da Alteracao Daniel
    
    iAlterado = 0

    Exibe_Dados_Fornecedor = SUCESSO

Exit Function

Erro_Exibe_Dados_Fornecedor:

    Exibe_Dados_Fornecedor = gErr

    Select Case gErr

        Case 12733, 12820, 12355, 89402, 89403, 89415
        
        Case 12320
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_SELECIONADO", gErr, objTipoFornecedor.iCodigo)
        
        Case 12319
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_FORNECEDOR_NAO_CADASTRADO", gErr, objTipoFornecedor.iCodigo)

        Case 12819
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objFornecedor.sContaContabil)
        
        Case 90546
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objFornecedor.sContaFornConsig)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160611)

    End Select

    Exit Function

End Function

'"""""""""""""""""""""""""""""""""""""""""""
'"      ROTINAS DAS TELAS DE BROWSE        "
'"""""""""""""""""""""""""""""""""""""""""""

'BROWSE TIPO_FORNECEDOR :

Private Sub TipoFornecedorLabel_Click()

Dim objTipoFornecedor As New ClassTipoFornecedor
Dim colSelecao As New Collection

    If Tipo.ListIndex <> -1 Then objTipoFornecedor.iCodigo = Codigo_Extrai(Tipo.Text)

    Call Chama_Tela("TipoFornecedorLista", colSelecao, objTipoFornecedor, objEventoTipoFornecedor)

End Sub

'BROWSE CONDICAO_PAGTO :

Private Sub CondicaoPagtoLabel_Click()

Dim objCondicaoPagto As New ClassCondicaoPagto
Dim colSelecao As New Collection

    objCondicaoPagto.iCodigo = CondPagto_Extrai(CondicaoPagto)

    Call Chama_Tela("CondicaoPagtoCPLista", colSelecao, objCondicaoPagto, objEventoCondicaoPagto)

End Sub

'BROWSE PLANO_CONTA :

Private Sub ContaContabilLabel_Click()

Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim iContaPreenchida As Integer
Dim sConta As String
Dim lErro As Long

On Error GoTo Erro_ContaContabilLabel_Click

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaContabil.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 12774

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    Call Chama_Tela("PlanoContaCPLista", colSelecao, objPlanoConta, objEventoContaContabil)

    Exit Sub

Erro_ContaContabilLabel_Click:

    Select Case Err

    Case 12774

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160612)

    End Select

    Exit Sub

End Sub

Private Sub ContaFornConsigLabel_Click()

Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim iContaPreenchida As Integer
Dim sConta As String
Dim lErro As Long

On Error GoTo Erro_ContaFornConsigLabel_Click

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaFornConsig.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 12774

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    Call Chama_Tela("PlanoContaCPLista", colSelecao, objPlanoConta, objEventoContaFornConsig)

    Exit Sub

Erro_ContaFornConsigLabel_Click:

    Select Case Err

    Case 12774

    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160613)

    End Select

    Exit Sub

End Sub

'RETORNO TIPO_FORNECEDOR :

Private Sub objEventoTipoFornecedor_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTipoFornecedor As ClassTipoFornecedor

On Error GoTo Erro_objEventoTipoFornecedor_evSelecao

    Set objTipoFornecedor = obj1

    If objTipoFornecedor Is Nothing Then
        Tipo.Text = ""
    Else
        Tipo.Text = CStr(objTipoFornecedor.iCodigo)
        lErro = Combo_Item_Seleciona(Tipo)
        If lErro <> SUCESSO Then Error 19114
    End If

    Me.Show

    Exit Sub

Erro_objEventoTipoFornecedor_evSelecao:

    Select Case Err

        Case 19114

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160614)

    End Select

    Exit Sub

End Sub

'RETORNO CONDICAO_PAGTO :

Private Sub objEventoCondicaoPagto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCondicaoPagto As ClassCondicaoPagto

On Error GoTo Erro_objEventoCondicaoPagto_evSelecao

    Set objCondicaoPagto = obj1

    CondicaoPagto.Text = CStr(objCondicaoPagto.iCodigo)
    Call CondicaoPagto_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

Erro_objEventoCondicaoPagto_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160615)

    End Select

    Exit Sub

End Sub

'RETORNO PLANO_CONTA :

Private Sub objEventoContaContabil_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaContabil_evSelecao

    Set objPlanoConta = obj1

    If objPlanoConta.sConta = "" Then

        ContaContabil.Text = ""

    Else

        ContaContabil.PromptInclude = False

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then Error 19108

        ContaContabil.Text = sContaEnxuta

        ContaContabil.PromptInclude = True

    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case Err

        Case 19108
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160616)

    End Select

    Exit Sub

End Sub

Private Sub objEventoContaFornConsig_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaFornConsig_evSelecao

    Set objPlanoConta = obj1

    If objPlanoConta.sConta = "" Then

        ContaFornConsig.Text = ""

    Else

        ContaFornConsig.PromptInclude = False

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then Error 19108

        ContaFornConsig.Text = sContaEnxuta

        ContaFornConsig.PromptInclude = True

    End If

    Me.Show

    Exit Sub

Erro_objEventoContaFornConsig_evSelecao:

    Select Case Err

        Case 19108
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160617)

    End Select

    Exit Sub

End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA
'""""""""""""""""""""""""""""""""""""""""""""""

'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Fornecedores"

    'Lê os dados da Tela Fornecedores
    lErro = Le_Dados_Fornecedor(objFornecedor)
    If lErro <> SUCESSO Then Error 12863

    'Preenche a coleção colCampoValor, com nome do campo,
    'Valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo

    colCampoValor.Add "Codigo", objFornecedor.lCodigo, 0, "Codigo"
    colCampoValor.Add "RazaoSocial", objFornecedor.sRazaoSocial, STRING_FORNECEDOR_RAZAO_SOC, "RazaoSocial"
    colCampoValor.Add "NomeReduzido", objFornecedor.sNomeReduzido, STRING_FORNECEDOR_NOME_REDUZIDO, "NomeReduzido"
    colCampoValor.Add "Tipo", objFornecedor.iTipo, 0, "Tipo"
    colCampoValor.Add "Observacao", objFornecedor.sObservacao, STRING_FORNECEDOR_OBS, "Observacao"
    colCampoValor.Add "CondicaoPagto", objFornecedor.iCondicaoPagto, 0, "condicaopagto"
    colCampoValor.Add "Desconto", objFornecedor.dDesconto, 0, "Desconto"

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 12863

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160618)

    End Select

    Exit Sub

End Sub

'Preenche os campos da tela com os correspondentes do BD
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objFornecedorEstatistica As New ClassFilialFornecedorEst
Dim objFilialFornecedor As New ClassFilialFornecedor

On Error GoTo Erro_Tela_Preenche

    objFornecedor.lCodigo = colCampoValor.Item("Codigo").vValor

    If objFornecedor.lCodigo <> 0 Then

        lErro = CF("Fornecedor_Le_Estendida", objFornecedor, objFornecedorEstatistica)
        If lErro <> SUCESSO And lErro <> 52701 Then Error 19214

        'Verifica se Fornecedor não está cadastrado
        If lErro <> SUCESSO Then Error 19215
        
        '***** Janaina ******
        'Inicializa objFilialFornecedor
        objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
        objFilialFornecedor.iCodFilial = FILIAL_MATRIZ
        
        '*******************

        lErro = Exibe_Dados_Fornecedor(objFornecedor, objFornecedorEstatistica)
        If lErro <> SUCESSO Then Error 19216
        
    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 19215
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INEXISTENTE", Err)

        Case 19214, 19216

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160619)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long

    'Alteracao Daniel
    Set objEventoContaDespesa = Nothing
    Set objEventoNatureza = Nothing
    
    Set objEventoTipoFornecedor = Nothing
    Set objEventoCondicaoPagto = Nothing
    Set objEventoContaContabil = Nothing
    Set objEventoContaFornConsig = Nothing
    Set objEventoFornecedor = Nothing
    
    Set gobjAnotacao = Nothing
    
    Call gobjTabEnd.Finaliza
    Set gobjTabEnd = Nothing
    
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object
    
    Parent.HelpContextID = IDH_FORNECEDOR_IDENT
    Set Form_Load_Ocx = Me
    Caption = "Fornecedores"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Fornecedores"
    
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
   ' Parent.UnloadDoFilho
    
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
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Tipo Then
            Call TipoFornecedorLabel_Click
        ElseIf Me.ActiveControl Is ContaDespesa Then
            Call LabelContaDespesa_Click
        ElseIf Me.ActiveControl Is CondicaoPagto Then
            Call CondicaoPagtoLabel_Click
        ElseIf Me.ActiveControl Is ContaContabil Then
            Call ContaContabilLabel_Click
        ElseIf Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        ElseIf Me.ActiveControl Is NomeReduzido Then
            Call LabelNomeReduzido_Click
        ElseIf Me.ActiveControl Is ContaFornConsig Then
            Call ContaFornConsigLabel_Click
        End If
        
    End If
    
End Sub

Private Sub Fornecedor_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Fornecedor(Index), Source, X, Y)
End Sub

Private Sub Fornecedor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Fornecedor(Index), Button, Shift, X, Y)
End Sub

Private Sub LabelObservacao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelObservacao, Source, X, Y)
End Sub

Private Sub LabelObservacao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelObservacao, Button, Shift, X, Y)
End Sub

Private Sub TipoFornecedorLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoFornecedorLabel, Source, X, Y)
End Sub

Private Sub TipoFornecedorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoFornecedorLabel, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeReduzido_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeReduzido, Source, X, Y)
End Sub

Private Sub LabelNomeReduzido_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeReduzido, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
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

Private Sub Label29_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label29, Source, X, Y)
End Sub

Private Sub Label29_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label29, Button, Shift, X, Y)
End Sub

Private Sub PedidosEmAberto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PedidosEmAberto, Source, X, Y)
End Sub

Private Sub PedidosEmAberto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PedidosEmAberto, Button, Shift, X, Y)
End Sub

Private Sub Label23_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label23, Source, X, Y)
End Sub

Private Sub Label23_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label23, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub ValorAcumuladoDevolucoes_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorAcumuladoDevolucoes, Source, X, Y)
End Sub

Private Sub ValorAcumuladoDevolucoes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorAcumuladoDevolucoes, Button, Shift, X, Y)
End Sub

Private Sub Label19_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label19, Source, X, Y)
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label19, Button, Shift, X, Y)
End Sub

Private Sub NumTotalDevolucoes_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumTotalDevolucoes, Source, X, Y)
End Sub

Private Sub NumTotalDevolucoes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumTotalDevolucoes, Button, Shift, X, Y)
End Sub

Private Sub DataUltDevolucao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataUltDevolucao, Source, X, Y)
End Sub

Private Sub DataUltDevolucao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataUltDevolucao, Button, Shift, X, Y)
End Sub

Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub

Private Sub DataUltimaCompra_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataUltimaCompra, Source, X, Y)
End Sub

Private Sub DataUltimaCompra_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataUltimaCompra, Button, Shift, X, Y)
End Sub

Private Sub DataPrimeiraCompra_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataPrimeiraCompra, Source, X, Y)
End Sub

Private Sub DataPrimeiraCompra_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataPrimeiraCompra, Button, Shift, X, Y)
End Sub

Private Sub ValorAcumuladoCompras_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorAcumuladoCompras, Source, X, Y)
End Sub

Private Sub ValorAcumuladoCompras_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorAcumuladoCompras, Button, Shift, X, Y)
End Sub

Private Sub MediaCompra_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(MediaCompra, Source, X, Y)
End Sub

Private Sub MediaCompra_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(MediaCompra, Button, Shift, X, Y)
End Sub

Private Sub NumeroCompras_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumeroCompras, Source, X, Y)
End Sub

Private Sub NumeroCompras_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumeroCompras, Button, Shift, X, Y)
End Sub

Private Sub Label18_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label18, Source, X, Y)
End Sub

Private Sub Label18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label18, Button, Shift, X, Y)
End Sub

Private Sub Label17_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label17, Source, X, Y)
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label17, Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label15, Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15, Button, Shift, X, Y)
End Sub

Private Sub Label16_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label16, Source, X, Y)
End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label16, Button, Shift, X, Y)
End Sub

Private Sub Label21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label21, Source, X, Y)
End Sub

Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label21, Button, Shift, X, Y)
End Sub

Private Sub SaldoTitulos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SaldoTitulos, Source, X, Y)
End Sub

Private Sub SaldoTitulos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SaldoTitulos, Button, Shift, X, Y)
End Sub

Private Sub Label20_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label20, Source, X, Y)
End Sub

Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label20, Button, Shift, X, Y)
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub
Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label36_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label36, Source, X, Y)
End Sub

Private Sub Label36_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label36, Button, Shift, X, Y)
End Sub

Private Sub Label34_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label34, Source, X, Y)
End Sub

Private Sub Label34_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label34, Button, Shift, X, Y)
End Sub

Private Sub Label35_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label35, Source, X, Y)
End Sub

Private Sub Label35_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label35, Button, Shift, X, Y)
End Sub

Private Sub Label32_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label32, Source, X, Y)
End Sub

Private Sub Label32_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label32, Button, Shift, X, Y)
End Sub

Private Sub Label22_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label22, Source, X, Y)
End Sub

Private Sub Label22_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label22, Button, Shift, X, Y)
End Sub

Private Sub Label25_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label25, Source, X, Y)
End Sub

Private Sub Label25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label25, Button, Shift, X, Y)
End Sub

Private Sub Label28_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label28, Source, X, Y)
End Sub

Private Sub Label28_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label28, Button, Shift, X, Y)
End Sub

Private Sub Label24_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label24, Source, X, Y)
End Sub

Private Sub Label24_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label24, Button, Shift, X, Y)
End Sub

Private Sub CondicaoPagtoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CondicaoPagtoLabel, Source, X, Y)
End Sub

Private Sub CondicaoPagtoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CondicaoPagtoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub ContaContabilLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContaContabilLabel, Source, X, Y)
End Sub

Private Sub ContaFornConsigLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContaFornConsigLabel, Source, X, Y)
End Sub

Private Sub ContaContabilLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContaContabilLabel, Button, Shift, X, Y)
End Sub

Private Sub ContaFornConsigLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContaFornConsigLabel, Button, Shift, X, Y)
End Sub

Public Property Get hWnd() As Long
   hWnd = UserControl.hWnd
End Property

Public Property Get Height() As Long
   Height = UserControl.Height
End Property

Public Property Get Width() As Long
   Width = UserControl.Width
End Property

'Alteracao Daniel
Private Sub GeraICMS_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

'Alteracao Daniel
Private Sub InscricaoINSS_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

'Alteracao Daniel
Private Sub ContaDespesa_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

'Alteracao Daniel
Private Sub TipoCobranca_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

'Alteracao Daniel
Private Sub TipoFrete_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

'Alteracao Daniel
Private Sub LabelContaDespesa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelContaDespesa, Source, X, Y)
End Sub

'Alteracao Daniel
Private Sub LabelContaDespesa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelContaDespesa, Button, Shift, X, Y)
End Sub

'Alteracao Daniel
Private Sub ContaDespesa_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaMascarada As String
Dim sContaFormatada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ContaDespesa_Validate

    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica_Modulo", ContaDespesa.Text, ContaDespesa.ClipText, objPlanoConta, MODULO_CONTASAPAGAR)
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 101499

    If lErro = SUCESSO Then

        sContaFormatada = objPlanoConta.sConta

        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then gError 106000
        
        ContaDespesa.PromptInclude = False
        ContaDespesa.Text = sContaMascarada
        ContaDespesa.PromptInclude = True
        
    'se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then

        lErro = CF("Conta_Critica", ContaDespesa.Text, sContaFormatada, objPlanoConta, MODULO_CONTASAPAGAR)
        If lErro <> SUCESSO And lErro <> 5700 Then gError 106001

        If lErro = 5700 Then gError 106002

    End If

    Exit Sub

Erro_ContaDespesa_Validate:

    Cancel = True

    Select Case gErr

        Case 106001, 101499

        Case 106002
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", gErr, ContaDespesa.Text)

        Case 106000
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160620)

    End Select

End Sub

'Alteracao Daniel
Private Sub objEventoContaDespesa_evSelecao(obj1 As Object)

Dim objPlanoConta As ClassPlanoConta
Dim lErro As Long
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaDespesa_evSelecao

    Set objPlanoConta = obj1

    If objPlanoConta.sConta = "" Then

        ContaDespesa.Text = ""

    Else

        ContaDespesa.PromptInclude = False

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then gError 106005

        ContaDespesa.Text = sContaEnxuta

        ContaDespesa.PromptInclude = True

    End If

    Me.Show

    Exit Sub

Erro_objEventoContaDespesa_evSelecao:

    Select Case gErr

        Case 106005
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160621)

    End Select

End Sub

'Alteracao Daniel
Private Function Carrega_TipoCobranca() As Long
'Carrega na combobox os Tipos de Cobrança

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome

On Error GoTo Erro_Carrega_TipoCobranca

    'Lê o nome e o codigo de todos os Tipos de Cobrança
    lErro = CF("Cod_Nomes_Le", "TiposDeCobranca", "Codigo", "Descricao", STRING_TIPOSDECOBRANCA_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 106007

    'Carrega na combo de Tipos de Cobrança os Tipos que estão em colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao

        TipoCobranca.AddItem CInt(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        TipoCobranca.ItemData(TipoCobranca.NewIndex) = objCodigoDescricao.iCodigo

    Next

    Carrega_TipoCobranca = SUCESSO

    Exit Function

Erro_Carrega_TipoCobranca:

    Carrega_TipoCobranca = gErr

    Select Case gErr

        Case 106007

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160622)

    End Select

End Function

'Alteracao Daniel
Private Sub LabelContaDespesa_Click()

Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim iContaPreenchida As Integer
Dim sConta As String
Dim lErro As Long

On Error GoTo Erro_LabelContaDespesa_Click

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaDespesa.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then gError 106008

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    Call Chama_Tela("PlanoContaCPLista", colSelecao, objPlanoConta, objEventoContaDespesa)

    Exit Sub

Erro_LabelContaDespesa_Click:

    Select Case gErr

    Case 106008

    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160623)

    End Select

End Sub

Public Sub Anotacao_Extrai(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long

On Error GoTo Erro_Anotacao_Extrai

    objAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_FORNECEDOR
    If Len(Trim(Codigo.Text)) > 0 Then
        objAnotacao.sID = Codigo.Text
    Else
        objAnotacao.sID = ""
        If Not (gobjAnotacao Is Nothing) Then
            objAnotacao.sTextoCompleto = gobjAnotacao.sTextoCompleto
            objAnotacao.sTitulo = gobjAnotacao.sTitulo
        End If
    End If
    
    Exit Sub
     
Erro_Anotacao_Extrai:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160626)
     
    End Select
     
    Exit Sub

End Sub

Public Sub Anotacao_Preenche(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long

On Error GoTo Erro_Anotacao_Preenche

    'guarda o texto digitado
    Set gobjAnotacao = objAnotacao
        
    Exit Sub
     
Erro_Anotacao_Preenche:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160627)
     
    End Select
     
    Exit Sub

End Sub

Private Sub Inscricaosuframa_Change()
     iAlterado = REGISTRO_ALTERADO
End Sub

Private Function Carrega_TipoFrete() As Long

Dim lErro As Long
Dim objTipoFrete As ClassTipoFrete
Dim iIndice As Integer
Dim colTipoFrete As Collection

On Error GoTo Erro_Carrega_TipoFrete
    
    'Le Código e Nome de FilialEmpresa
    lErro = CF("TipoFrete_Le", colTipoFrete)
    If lErro <> SUCESSO Then gError 207471

    TipoFrete.Clear

    For Each objTipoFrete In colTipoFrete

        'coloca na combo
        TipoFrete.AddItem objTipoFrete.sDescricao
        TipoFrete.ItemData(TipoFrete.NewIndex) = objTipoFrete.iCodigoNFE

    Next

    Carrega_TipoFrete = SUCESSO

    Exit Function

Erro_Carrega_TipoFrete:

    Carrega_TipoFrete = gErr

    Select Case gErr

        Case 207471

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207472)

    End Select

    Exit Function

End Function

Private Function Inicializa_Tab_Inscricao() As Long

Dim lErro  As Long

On Error GoTo Erro_Inicializa_Tab_Inscricao

    RegimeTributario.AddItem REGIME_TRIBUTARIO_NORMAL_TEXTO
    RegimeTributario.ItemData(RegimeTributario.NewIndex) = REGIME_TRIBUTARIO_NORMAL
    RegimeTributario.AddItem REGIME_TRIBUTARIO_SIMPLES_TEXTO
    RegimeTributario.ItemData(RegimeTributario.NewIndex) = REGIME_TRIBUTARIO_SIMPLES

    RegimeTributario.ListIndex = 0
    
    Inicializa_Tab_Inscricao = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Tab_Inscricao:
    
    Inicializa_Tab_Inscricao = gErr
    
    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207400)
        
    End Select
        
    Exit Function
            
End Function

Private Sub IEIsento_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub IENaoContrib_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Trata_IE()
    If Len(Trim(InscricaoEstadual.Text)) > 0 Then
        IEIsento.Value = vbUnchecked
        IEIsento.Enabled = False
        If InscricaoEstadual.Text <> sIEAnt Then
            IENaoContrib.Value = vbUnchecked
        End If
    Else
        If InscricaoEstadual.Text <> sIEAnt Then
            IEIsento.Value = vbChecked
            IENaoContrib.Value = vbChecked
        End If
        IEIsento.Enabled = True
    End If
    sIEAnt = InscricaoEstadual.Text
End Sub

Private Sub InscricaoEstadual_Validate(Cancel As Boolean)
    Call Trata_IE
End Sub


Private Function Inicializa_Mascara_Natureza() As Long
'inicializa a mascara da Natureza

Dim sMascaraNatureza As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascara_Natureza

    'Inicializa a máscara da Natureza
    sMascaraNatureza = String(STRING_NATMOVCTA_CODIGO, 0)
    
    'Armazena em sMascaraNatureza a mascara a ser a ser exibida no campo Natureza
    lErro = MascaraItem(SEGMENTO_NATMOVCTA, sMascaraNatureza)
    If lErro <> SUCESSO Then gError 122836
    
    'coloca a mascara na tela.
    Natureza.Mask = sMascaraNatureza
    
    Inicializa_Mascara_Natureza = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascara_Natureza:

    Inicializa_Mascara_Natureza = gErr
    
    Select Case gErr
    
        Case 122836
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARAITEM", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
        
    End Select

    Exit Function

End Function

Private Sub objEventoNatureza_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNatMovCta As ClassNatMovCta
Dim sNaturezaEnxuta As String

On Error GoTo Erro_objEventoNatureza_evSelecao

    Set objNatMovCta = obj1

    If objNatMovCta.sCodigo = "" Then
        
        Natureza.PromptInclude = False
        Natureza.Text = ""
        Natureza.PromptInclude = True
    
    Else

        sNaturezaEnxuta = String(STRING_NATMOVCTA_CODIGO, 0)
    
        lErro = Mascara_RetornaItemEnxuto(SEGMENTO_NATMOVCTA, objNatMovCta.sCodigo, sNaturezaEnxuta)
        If lErro <> SUCESSO Then gError 122833

        Natureza.PromptInclude = False
        Natureza.Text = sNaturezaEnxuta
        Natureza.PromptInclude = True
    
    End If

    Call Natureza_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

Erro_objEventoNatureza_evSelecao:

    Select Case gErr

        Case 122833

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub Natureza_Validate(Cancel As Boolean)
     
Dim lErro As Long
Dim sNaturezaFormatada As String
Dim iNaturezaPreenchida As Integer
Dim objNatMovCta As New ClassNatMovCta

On Error GoTo Erro_Natureza_Validate

    If Len(Natureza.ClipText) > 0 Then

        sNaturezaFormatada = String(STRING_NATMOVCTA_CODIGO, 0)

        'critica o formato da Natureza
        lErro = CF("Item_Formata", SEGMENTO_NATMOVCTA, Natureza.Text, sNaturezaFormatada, iNaturezaPreenchida)
        If lErro <> SUCESSO Then gError 122826
        
        'Obj recebe código
        objNatMovCta.sCodigo = sNaturezaFormatada
        
        'Verifica se a Natureza é analítica e se seu Tipo Corresponde a um pagamento
        lErro = CF("Natureza_Critica", objNatMovCta, NATUREZA_TIPO_PAGAMENTO)
        If lErro <> SUCESSO Then gError 122843
        
        'Coloca a Descrição da Natureza na Tela
        LabelNaturezaDesc.Caption = objNatMovCta.sDescricao
        
    Else
    
        LabelNaturezaDesc.Caption = ""
    
    End If
    
    Exit Sub
    
Erro_Natureza_Validate:

    Cancel = True

    Select Case gErr
    
        Case 122826, 122843
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
        
    End Select

    Exit Sub
    
End Sub


Public Sub LabelNatureza_Click()

    Dim objNatMovCta As New ClassNatMovCta
    Dim colSelecao As New Collection

    objNatMovCta.sCodigo = Natureza.ClipText
    
    colSelecao.Add NATUREZA_TIPO_PAGAMENTO
    
    Call Chama_Tela("NatMovCtaLista", colSelecao, objNatMovCta, objEventoNatureza, "Tipo = ?")

End Sub

Public Sub Natureza_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub


