VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.UserControl FiliaisFornecedoresOcx 
   ClientHeight    =   5205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9135
   KeyPreview      =   -1  'True
   ScaleHeight     =   5200
   ScaleMode       =   0  'User
   ScaleWidth      =   9135
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4050
      Index           =   4
      Left            =   135
      TabIndex        =   87
      Top             =   1035
      Visible         =   0   'False
      Width           =   8880
      Begin VB.TextBox IdEstrangeiro 
         Height          =   315
         Left            =   6210
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   111
         ToolTipText     =   "número do passaporte ou outro documento legal para identificar pessoa estrangeira"
         Top             =   2610
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
         Left            =   5460
         TabIndex        =   110
         Top             =   1695
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
         Left            =   4500
         TabIndex        =   96
         Top             =   1740
         Value           =   1  'Checked
         Width           =   960
      End
      Begin VB.ComboBox RegimeTributario 
         Height          =   315
         ItemData        =   "FiliaisFornecedores2Ocx.ctx":0000
         Left            =   2100
         List            =   "FiliaisFornecedores2Ocx.ctx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   95
         Top             =   3585
         Width           =   2760
      End
      Begin VB.Frame SSFrame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   240
         TabIndex        =   97
         Top             =   150
         Width           =   8370
         Begin VB.Label FornecedorLabel 
            Caption         =   "fornecedor"
            Height          =   210
            Index           =   2
            Left            =   1320
            TabIndex        =   101
            Top             =   210
            Width           =   6885
         End
         Begin VB.Label FilialLabel 
            Caption         =   "filial"
            Height          =   210
            Index           =   1
            Left            =   1320
            TabIndex        =   100
            Top             =   510
            Width           =   6945
         End
         Begin VB.Label Label10 
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
            Top             =   195
            Width           =   1020
         End
         Begin VB.Label Label9 
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
            Height          =   210
            Left            =   720
            TabIndex        =   98
            Top             =   495
            Width           =   480
         End
      End
      Begin VB.TextBox Observacao 
         Height          =   300
         Left            =   2085
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   94
         Top             =   3090
         Width           =   6180
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
         Left            =   4500
         TabIndex        =   89
         Top             =   1215
         Value           =   1  'Checked
         Width           =   2580
      End
      Begin MSMask.MaskEdBox CGC 
         Height          =   300
         Left            =   2085
         TabIndex        =   88
         Top             =   1200
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   14
         Mask            =   "99999999999999"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox InscricaoEstadual 
         Height          =   300
         Left            =   2085
         TabIndex        =   90
         Top             =   1665
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   18
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox InscricaoMunicipal 
         Height          =   300
         Left            =   2085
         TabIndex        =   91
         Top             =   2130
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   18
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox InscricaoINSS 
         Height          =   300
         Left            =   2085
         TabIndex        =   93
         Top             =   2610
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
         Left            =   6225
         TabIndex        =   92
         Top             =   2130
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   12
         Mask            =   "##.####-##-#"
         PromptChar      =   " "
      End
      Begin VB.Label Label24 
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
         Left            =   4440
         TabIndex        =   112
         Top             =   2655
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
         Left            =   405
         TabIndex        =   109
         Top             =   3615
         Width           =   1560
      End
      Begin VB.Label Label13 
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
         Left            =   4500
         TabIndex        =   107
         Top             =   2190
         Width           =   1605
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
         Height          =   180
         Left            =   240
         TabIndex        =   106
         Top             =   2190
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
         Left            =   315
         TabIndex        =   105
         Top             =   1695
         Width           =   1650
      End
      Begin VB.Label Label35 
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
         Height          =   180
         Left            =   990
         TabIndex        =   104
         Top             =   1260
         Width           =   1005
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
         Height          =   180
         Left            =   870
         TabIndex        =   103
         Top             =   3105
         Width           =   1095
      End
      Begin VB.Label Label2 
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
         Left            =   600
         TabIndex        =   102
         Top             =   2670
         Width           =   1350
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4065
      Index           =   5
      Left            =   135
      TabIndex        =   5
      Top             =   1050
      Visible         =   0   'False
      Width           =   8880
      Begin VB.Frame SSFrame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   180
         TabIndex        =   17
         Top             =   30
         Width           =   8490
         Begin VB.Label Label5 
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
            Height          =   210
            Left            =   720
            TabIndex        =   21
            Top             =   495
            Width           =   480
         End
         Begin VB.Label Label6 
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
            Top             =   195
            Width           =   1020
         End
         Begin VB.Label FilialLabel 
            Caption         =   "filial"
            Height          =   210
            Index           =   2
            Left            =   1305
            TabIndex        =   19
            Top             =   480
            Width           =   6945
         End
         Begin VB.Label FornecedorLabel 
            Caption         =   "fornecedor"
            Height          =   210
            Index           =   3
            Left            =   1290
            TabIndex        =   18
            Top             =   210
            Width           =   6960
         End
      End
      Begin TelasCpr.TabEndereco TabEnd 
         Height          =   3660
         Index           =   0
         Left            =   195
         TabIndex        =   108
         Top             =   810
         Width           =   8445
         _ExtentX        =   14896
         _ExtentY        =   6456
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3555
      Index           =   1
      Left            =   105
      TabIndex        =   0
      Top             =   1020
      Width           =   8880
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   2430
         Picture         =   "FiliaisFornecedores2Ocx.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Numeração Automática"
         Top             =   1785
         Width           =   300
      End
      Begin MSMask.MaskEdBox CodFilial 
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Top             =   1770
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "9999"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CodFornecedor 
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Top             =   1050
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         Mask            =   "99999999"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Nome 
         Height          =   315
         Left            =   1920
         TabIndex        =   4
         Top             =   2490
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Filial:"
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
         Left            =   570
         TabIndex        =   26
         Top             =   2550
         Width           =   1275
      End
      Begin VB.Label LabelCodFilial 
         AutoSize        =   -1  'True
         Caption         =   "Código Filial:"
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
         Left            =   735
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   25
         Top             =   1815
         Width           =   1110
      End
      Begin VB.Label LabelCodFornecedor 
         AutoSize        =   -1  'True
         Caption         =   "Código Fornecedor:"
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
         Left            =   180
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   24
         Top             =   1095
         Width           =   1680
      End
      Begin VB.Label Label1 
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
         Index           =   0
         Left            =   810
         TabIndex        =   23
         Top             =   390
         Width           =   1035
      End
      Begin VB.Label FornecedorLabel 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   1920
         TabIndex        =   22
         Top             =   345
         Width           =   1710
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3600
      Index           =   3
      Left            =   135
      TabIndex        =   62
      Top             =   990
      Visible         =   0   'False
      Width           =   8880
      Begin VB.Frame SSFrame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   210
         TabIndex        =   82
         Top             =   -30
         Width           =   8415
         Begin VB.Label Label8 
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
            Height          =   210
            Left            =   720
            TabIndex        =   86
            Top             =   495
            Width           =   480
         End
         Begin VB.Label Label7 
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
            TabIndex        =   85
            Top             =   195
            Width           =   1020
         End
         Begin VB.Label FilialLabel 
            Caption         =   "filial"
            Height          =   210
            Index           =   0
            Left            =   1305
            TabIndex        =   84
            Top             =   480
            Width           =   6945
         End
         Begin VB.Label FornecedorLabel 
            Caption         =   "fornecedor"
            Height          =   210
            Index           =   1
            Left            =   1320
            TabIndex        =   83
            Top             =   210
            Width           =   6885
         End
      End
      Begin VB.Frame SSFrame6 
         Caption         =   "Conta Corrente"
         Height          =   735
         Left            =   210
         TabIndex        =   75
         Top             =   2790
         Width           =   8415
         Begin MSMask.MaskEdBox ContaCorrente 
            Height          =   315
            Left            =   6150
            TabIndex        =   76
            Top             =   240
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
            Left            =   3720
            TabIndex        =   77
            Top             =   255
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
            Left            =   1575
            TabIndex        =   78
            Top             =   240
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
            Left            =   5370
            TabIndex        =   81
            Top             =   285
            Width           =   735
         End
         Begin VB.Label Label25 
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
            Height          =   210
            Left            =   2880
            TabIndex        =   80
            Top             =   285
            Width           =   735
         End
         Begin VB.Label Label28 
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
            Height          =   210
            Left            =   915
            TabIndex        =   79
            Top             =   285
            Width           =   630
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Contabilidade"
         Height          =   1230
         Left            =   210
         TabIndex        =   68
         Top             =   810
         Width           =   8415
         Begin MSMask.MaskEdBox ContaContabil 
            Height          =   315
            Left            =   2460
            TabIndex        =   69
            Top             =   270
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
            Left            =   6570
            TabIndex        =   70
            Top             =   270
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
            Left            =   2460
            TabIndex        =   71
            Top             =   765
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
            TabIndex        =   74
            Top             =   330
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
            Left            =   4470
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   73
            Top             =   225
            Width           =   2070
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
            TabIndex        =   72
            Top             =   825
            Width           =   2145
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Pagamento"
         Height          =   690
         Left            =   210
         TabIndex        =   63
         Top             =   2070
         Width           =   8415
         Begin VB.ComboBox TipoFrete 
            Height          =   315
            ItemData        =   "FiliaisFornecedores2Ocx.ctx":00EE
            Left            =   1755
            List            =   "FiliaisFornecedores2Ocx.ctx":00F8
            Style           =   2  'Dropdown List
            TabIndex        =   65
            Top             =   225
            Width           =   1125
         End
         Begin VB.ComboBox TipoCobranca 
            Height          =   315
            Left            =   5100
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   225
            Width           =   2865
         End
         Begin VB.Label Label39 
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
            Left            =   495
            TabIndex        =   67
            Top             =   255
            Width           =   1215
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cobrança:"
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
            Left            =   3465
            TabIndex        =   66
            Top             =   270
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3600
      Index           =   2
      Left            =   105
      TabIndex        =   57
      Top             =   990
      Visible         =   0   'False
      Width           =   8880
      Begin VB.Frame Frame7 
         Caption         =   "Categorias"
         Height          =   3270
         Left            =   1380
         TabIndex        =   58
         Top             =   210
         Width           =   5175
         Begin VB.ComboBox ComboCategoriaFornecedor 
            Height          =   315
            Left            =   1590
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   660
            Width           =   1545
         End
         Begin VB.ComboBox ComboCategoriaFornecedorItem 
            Height          =   315
            Left            =   2670
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   630
            Width           =   1635
         End
         Begin MSFlexGridLib.MSFlexGrid GridCategoria 
            Height          =   2775
            Left            =   810
            TabIndex        =   61
            Top             =   330
            Width           =   3540
            _ExtentX        =   6244
            _ExtentY        =   4895
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
      Height          =   3570
      Index           =   6
      Left            =   135
      TabIndex        =   6
      Top             =   1050
      Visible         =   0   'False
      Width           =   8880
      Begin VB.Frame Frame2 
         Caption         =   "Pedidos de Compra"
         Height          =   1365
         Left            =   4710
         TabIndex        =   49
         Top             =   2115
         Width           =   4020
         Begin VB.Frame SSFrame4 
            Caption         =   "Atraso de Entrega"
            Height          =   600
            Left            =   195
            TabIndex        =   50
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
               TabIndex        =   54
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
               TabIndex        =   53
               Top             =   300
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label MediaAtraso 
               Caption         =   "0"
               Height          =   210
               Left            =   1095
               TabIndex        =   52
               Top             =   300
               Width           =   510
            End
            Begin VB.Label MaiorAtraso 
               Caption         =   "0"
               Height          =   210
               Left            =   2655
               TabIndex        =   51
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
            Left            =   270
            TabIndex        =   56
            Top             =   300
            Width           =   945
         End
         Begin VB.Label PedidosEmAberto 
            Caption         =   "0"
            Height          =   210
            Left            =   1305
            TabIndex        =   55
            Top             =   315
            Width           =   585
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Devoluções"
         Height          =   1110
         Left            =   4710
         TabIndex        =   28
         Top             =   900
         Width           =   4020
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
            TabIndex        =   34
            Top             =   825
            Width           =   600
         End
         Begin VB.Label Label11 
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
            TabIndex        =   33
            Top             =   255
            Width           =   1500
         End
         Begin VB.Label ValorAcumuladoDevolucoes 
            Caption         =   "0,00"
            Height          =   210
            Left            =   1980
            TabIndex        =   32
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
            TabIndex        =   31
            Top             =   540
            Width           =   720
         End
         Begin VB.Label NumTotalDevolucoes 
            Caption         =   "0"
            Height          =   210
            Left            =   1980
            TabIndex        =   30
            Top             =   540
            Width           =   585
         End
         Begin VB.Label DataUltDevolucao 
            Caption         =   "  /  /    "
            Height          =   210
            Left            =   1980
            TabIndex        =   29
            Top             =   825
            Width           =   1170
         End
      End
      Begin VB.Frame SSFrame8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   240
         TabIndex        =   12
         Top             =   0
         Width           =   8490
         Begin VB.Label FornecedorLabel 
            Caption         =   "fornecedor"
            Height          =   210
            Index           =   4
            Left            =   1305
            TabIndex        =   16
            Top             =   195
            Width           =   6960
         End
         Begin VB.Label FilialLabel 
            Caption         =   "filial"
            Height          =   210
            Index           =   3
            Left            =   1305
            TabIndex        =   15
            Top             =   480
            Width           =   6945
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
            TabIndex        =   14
            Top             =   195
            Width           =   1020
         End
         Begin VB.Label Label14 
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
            Height          =   210
            Left            =   720
            TabIndex        =   13
            Top             =   495
            Width           =   480
         End
      End
      Begin VB.Frame SSFrame5 
         Caption         =   "Compras"
         Height          =   1875
         Left            =   240
         TabIndex        =   35
         Top             =   1605
         Width           =   4020
         Begin VB.Label DataUltimaCompra 
            Caption         =   "  /  /    "
            Height          =   210
            Left            =   1830
            TabIndex        =   45
            Top             =   1515
            Width           =   1170
         End
         Begin VB.Label DataPrimeiraCompra 
            Caption         =   "  /  /    "
            Height          =   210
            Left            =   1830
            TabIndex        =   44
            Top             =   1170
            Width           =   1170
         End
         Begin VB.Label ValorAcumuladoCompras 
            Caption         =   "0,00"
            Height          =   210
            Left            =   1830
            TabIndex        =   43
            Top             =   870
            Width           =   1575
         End
         Begin VB.Label MediaCompra 
            Caption         =   "0,00"
            Height          =   210
            Left            =   1830
            TabIndex        =   42
            Top             =   570
            Width           =   1410
         End
         Begin VB.Label NumeroCompras 
            Caption         =   "0"
            Height          =   210
            Left            =   1830
            TabIndex        =   41
            Top             =   240
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
            TabIndex        =   40
            Top             =   225
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
            TabIndex        =   39
            Top             =   555
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
            TabIndex        =   38
            Top             =   1500
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
            TabIndex        =   37
            Top             =   1170
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
            Left            =   255
            TabIndex        =   36
            Top             =   870
            Width           =   1500
         End
      End
      Begin VB.Frame SSFrame7 
         Caption         =   "Saldos"
         Height          =   600
         Left            =   240
         TabIndex        =   46
         Top             =   900
         Width           =   4020
         Begin VB.Label SaldoTitulos 
            Caption         =   "0,00"
            Height          =   210
            Left            =   1815
            TabIndex        =   48
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
            TabIndex        =   47
            Top             =   285
            Width           =   990
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6900
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   75
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "FiliaisFornecedores2Ocx.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "FiliaisFornecedores2Ocx.ctx":0260
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "FiliaisFornecedores2Ocx.ctx":03EA
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "FiliaisFornecedores2Ocx.ctx":091C
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   4500
      Left            =   75
      TabIndex        =   27
      Top             =   660
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   7938
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
Attribute VB_Name = "FiliaisFornecedoresOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'VARIAVEIS GLOBAIS
Public gobjTabEnd As New ClassTabEndereco

Dim iFrameAtual As Integer
Public iAlterado As Integer

Dim sIEAnt As String

'********** Janaina **********
Const STRING_CATEGORIAFORNECEDOR_CATEGORIA = 20
Const STRING_CATEGORIAFORNECEDOR_DESCRICAO = 50
Const STRING_CATEGORIAFORNECEDOR_ITEM = 20
Const STRING_CATEGORIAFORNECEDORITEM_ITEM = 20
Const STRING_CATEGORIAFORNECEDORITEM_DESCRICAO = 50

Dim objGrid2 As AdmGrid
Dim objGridCategoria As AdmGrid
Dim iGrid_Categoria_Col As Integer
Dim iGrid_Valor_Col As Integer
'******************************

Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1
Private WithEvents objEventoContaFornConsig As AdmEvento
Attribute objEventoContaFornConsig.VB_VarHelpID = -1
Private WithEvents objEventoFilialFornecedor As AdmEvento
Attribute objEventoFilialFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1

'Alteracao Daniel
Private WithEvents objEventoContaDespesa As AdmEvento
Attribute objEventoContaDespesa.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_Complemento = 2
Private Const TAB_DadosFinanceiros = 3
Private Const TAB_Inscricoes = 4
Private Const TAB_Endereco = 5
Private Const TAB_Estatisticas = 6

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    'Gera código automático da proxima filialFornecedor
    lErro = CF("FiliaisFornecedores_Automatico", CLng(CodFornecedor.Text), iCodigo)
    If lErro <> SUCESSO Then Error 55709

    CodFilial.Text = CStr(iCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 55709
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160284)
    
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

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objFornecedor As New ClassFornecedor
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_BotaoExcluir_Click

    Parent.MousePointer = vbHourglass
    
    'Verifica se o codigo foi preenchido
    If Len(Trim(CodFornecedor.Text)) = 0 Then Error 12995

    'Verifica se o codigo da filial foi preenchido
    If Len(Trim(CodFilial.Text)) = 0 Then Error 12996

    'Verifica se é matriz
    If Trim(CodFilial.Text) = CStr(FILIAL_MATRIZ) Then Error 12997

    objFornecedor.lCodigo = CLng(CodFornecedor.Text)
    objFilialFornecedor.lCodFornecedor = CLng(CodFornecedor.Text)
    objFilialFornecedor.iCodFilial = CInt(CodFilial.Text)

    'Lê os dados do Fornecedor
    lErro = CF("Fornecedor_Le", objFornecedor)
    If lErro <> SUCESSO And lErro <> 12729 Then Error 12998

    'Verifica se Fornecedor nao esta cadastrado
    If lErro <> SUCESSO Then Error 12999

    'Le os dados da Filial Fornecedor
    lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
    If lErro <> SUCESSO And lErro <> 12929 Then Error 19000

    'Verifica se a Filial Fornecedor não esta cadastrada
    If lErro <> SUCESSO Then Error 19001

    'objFilialFornecedor.iCodFilial = FILIAL_MATRIZ
    objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo

    'envia aviso perguntando se realmente deseja excluir filial Fornecedor
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_FILIAL_FORNECEDOR")

    If vbMsgRes = vbYes Then

        lErro = CF("FiliaisFornecedores_Exclui", objFilialFornecedor)
        If lErro <> SUCESSO Then Error 19002

        'lErro = Exclui_Arvore_Filial(Filiais.Nodes, objFilialFornecedor)
        'If lErro <> SUCESSO Then Error 19003

        Call Limpa_Tela_FiliaisFornecedores

        iAlterado = 0

    End If

    Parent.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    Parent.MousePointer = vbDefault
    
    Select Case Err

        Case 12995
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODFORNECEDOR_NAO_PREENCHIDO", Err)

        Case 12996
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODFILIAL_NAO_PREENCHIDO", Err)

        Case 12997
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_FORNECEDOR_EXCLUSAO_MATRIZ", Err)

        Case 12998, 19000, 19002

        Case 12999
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", Err, objFornecedor.lCodigo)

        Case 19001
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_FORNECEDOR_INEXISTENTE", Err, objFilialFornecedor.iCodFilial, objFilialFornecedor.lCodFornecedor)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160285)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long
Dim iCodigo As Integer
Dim iIndice As Integer

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 19046

    Call Limpa_Tela_FiliaisFornecedores

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 19046 'cancelou operacao de gravacao , continua execucao normal

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160286)

    End Select

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub CodFilial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodFilial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CodFilial, iAlterado)
    
End Sub

Private Sub CodFilial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CodFilial_Validate

    'verifica se foi preenchido o campo Codigo Filial
    If Len(Trim(CodFilial.Text)) = 0 Then Exit Sub

    lErro = Inteiro_Critica(CodFilial.Text)
    If lErro <> SUCESSO Then Error 19302

    Exit Sub

Erro_CodFilial_Validate:

    Cancel = True


    Select Case Err

        Case 19302

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160287)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160288)

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
        If lErro <> SUCESSO Then gError 91239

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

        Case 91239

        Case Else

            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160289)

    End Select

End Function

Private Function Carrega_ComboCategoriaFornecedorItem() As Long
'Carrega a Combo CategoriaFornecedorItem

Dim lErro As Long
Dim colItensCategoria As New Collection
Dim objCategoriaFornecedorItem As ClassCategoriaFornItem

On Error GoTo Erro_Carrega_ComboCategoriaFornecedorItem

    'Lê a tabela CategoriaProdutoItem a partir da Categoria
    lErro = CategoriaFornecedor_Le_Itens(colItensCategoria)
    If lErro <> SUCESSO Then gError 91240

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

        Case 91240

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160290)

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

'"""""""""""""""""""""""""""""""""""""""""""
'"      ROTINAS DAS TELAS DE BROWSE        "
'"""""""""""""""""""""""""""""""""""""""""""

'BROWSE PLANO_CONTA :

Private Sub ContaContabilLabel_Click()

Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim iContaPreenchida As Integer
Dim sConta As String
Dim lErro As Long

On Error GoTo Erro_ContaContabilLabel_DblClick

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaContabil.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 19105

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    Call Chama_Tela("PlanoContaCPLista", colSelecao, objPlanoConta, objEventoContaContabil)

    Exit Sub

Erro_ContaContabilLabel_DblClick:

    Select Case Err

    Case 19105

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160291)

    End Select

    Exit Sub

End Sub

Private Sub ContaFornConsigLabel_Click()
'Função criada em 27/06/2001 por Luiz Gustavo de Freitas Nogueira
'Função criada em função da inclusão do campo ContaFornConsig na tela

Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim iContaPreenchida As Integer
Dim sConta As String
Dim lErro As Long

On Error GoTo Erro_ContaFornConsigLabel_Click

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaFornConsig.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then gError 90550

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    Call Chama_Tela("PlanoContaCPLista", colSelecao, objPlanoConta, objEventoContaFornConsig)

    Exit Sub

Erro_ContaFornConsigLabel_Click:

    Select Case Err

    Case 90550

    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160292)

    End Select

    Exit Sub

End Sub

Private Sub ContaCorrente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

 Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Function Exclui_Arvore_Filial(colNodes As Nodes, objFilialFornecedor As ClassFilialFornecedor) As Long

Dim objNode As Node
Dim sFilialFornecedor As String

    sFilialFornecedor = KEY_CARACTER & CStr(objFilialFornecedor.lCodFornecedor) & SEPARADOR & CStr(objFilialFornecedor.iCodFilial)

    For Each objNode In colNodes
        If objNode.Key = sFilialFornecedor Then
            colNodes.Remove sFilialFornecedor
            Exit For
        End If
    Next

    Exclui_Arvore_Filial = SUCESSO

    Exit Function

End Function

Private Sub Limpa_Tela_FiliaisFornecedores()

Dim lErro As Long
Dim iIndice As Integer
Dim sCodigoFornecedor As String
Dim objFilialEmpresa As New AdmFiliais

    'Guarda o Codigo de Fornecedor
    sCodigoFornecedor = CodFornecedor.Text

    Call Limpa_Tela(Me)

    'Mantém o Código de Fornecedor na tela
    CodFornecedor.Text = sCodigoFornecedor

    'Limpa os campos Filial existentes nos frames
    For iIndice = 0 To 2
        FilialLabel(iIndice).Caption = ""
    Next
    
    IEIsento.Value = vbChecked
    IENaoContrib.Value = vbChecked

    '******** Janaina ********
    Call Grid_Limpa(objGridCategoria)
    '*************************
    
    Call gobjTabEnd.Limpa_Tela

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim lCodigo As Long
Dim iCodigo As Integer
Dim iIndice As Integer

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 12956

    Call Limpa_Tela_FiliaisFornecedores

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 12956

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160293)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objEndereco As New ClassEndereco
Dim iIndice As Integer
Dim colEnderecos As New Collection

On Error GoTo Erro_Gravar_Registro

    Parent.MousePointer = vbHourglass
    
    'verifica se foi preenchido o Codigo do Fornecedor
    If Len(Trim(CodFornecedor.Text)) = 0 Then gError 12958

    'verifica se foi preenchido Codigo da Filial
    If Len(Trim(CodFilial.Text)) = 0 Then gError 12959

    'verifica se foi preenchido o Nome da Filial
    If Len(Trim(Nome.Text)) = 0 Then gError 12961

    'Le os dados da tela respectivos a Filial Fornecedor
    lErro = Le_Dados_FilialFornecedor(objFilialFornecedor)
    If lErro <> SUCESSO Then gError 12962

    lErro = gobjTabEnd.Move_Endereco_Memoria(colEnderecos)
    If lErro <> SUCESSO Then gError 12963
    
    Set objEndereco = colEnderecos.Item(1)
    '****** Janaina ******

    'Verifica se para cada Categoria existente no grid de Categorias existe um Valor correspondente
    For iIndice = 1 To objGridCategoria.iLinhasExistentes
        'Verifica se a Categoria foi preenchida
         If Len(Trim(GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL))) <> 0 And Len(Trim(GridCategoria.TextMatrix(iIndice, GRID_VALOR_COL))) = 0 Then gError 91241
    Next
    
'*********************

    lErro = CF("InscricaoEstadual_Critica", objFilialFornecedor.sInscricaoEstadual, objEndereco.sSiglaEstado)
    If lErro <> SUCESSO Then gError 189578

    lErro = Trata_Alteracao(objFilialFornecedor, objFilialFornecedor.lCodFornecedor, objFilialFornecedor.iCodFilial)
    If lErro <> SUCESSO Then gError 80450
    
    'Se o CGC estiver Preenchido
    If Len(Trim(objFilialFornecedor.sCgc)) > 0 Then
        'Verifica se tem outro Cliente com o mesmo CGC e dá aviso
        lErro = CF("FilialFornecedor_Testa_CGC", objFilialFornecedor.lCodFornecedor, 0, objFilialFornecedor.sCgc)
        If lErro <> SUCESSO Then gError 80218
    End If
    
    lErro = CF("FiliaisFornecedores_Grava", objFilialFornecedor, objEndereco)
    If lErro <> SUCESSO Then gError 12964

    'alterar a definição da Filial Fornecedor na TreeView
    'lErro = Alterar_Arvore_FilialFornecedor(Filiais.Nodes, objFilialFornecedor)
    'If lErro <> SUCESSO And lErro <> 12994 Then gError 12965

    'Se a filial não estava cadastrada na arvore
'    If lErro = 12994 Then
'        'inserir a filial na arvore
'        Call Inserir_Arvore_FilialFornecedor(Filiais.Nodes, objFilialFornecedor)
'
'    End If

    iAlterado = 0

    Parent.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Parent.MousePointer = vbDefault
    
    Gravar_Registro = gErr

    Select Case gErr

        Case 12958
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODFORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 12959
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODFILIAL_NAO_PREENCHIDO", gErr)

        Case 12961
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOMEFILIAL_NAO_PREENCHIDO", gErr)

        Case 12962, 12963, 12964, 80218, 80450, 91241, 189578

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160294)

    End Select

    Exit Function

End Function

Function Alterar_Arvore_FilialFornecedor(colNodes As Nodes, objFilialFornecedor As ClassFilialFornecedor) As Long

Dim lErro As Long
Dim objNode As Node
Dim sKeyFilial As String
Dim sMatriz As String
Dim iAchou As Integer

On Error GoTo Erro_Alterar_Arvore_FilialFornecedor

    sKeyFilial = KEY_CARACTER & CStr(objFilialFornecedor.lCodFornecedor) & SEPARADOR & CStr(objFilialFornecedor.iCodFilial)
    sMatriz = KEY_CARACTER & CStr(objFilialFornecedor.lCodFornecedor)

    iAchou = 0

    For Each objNode In colNodes

        If objNode.Key = sKeyFilial Then

            objNode.Text = objFilialFornecedor.sNome

            iAchou = 1

            Exit For

        End If

    Next

    'se não achou a filial na arvore
    If iAchou = 0 Then Error 12994

    Alterar_Arvore_FilialFornecedor = SUCESSO

    Exit Function

Erro_Alterar_Arvore_FilialFornecedor:

    Alterar_Arvore_FilialFornecedor = Err

    Select Case Err

        Case 12994

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160295)

    End Select

    Exit Function

End Function

Private Function Inserir_Arvore_FilialFornecedor(colNodes As Nodes, objFilialFornecedor As ClassFilialFornecedor) As Long
'insere a Filial Fornecedor na lista de Fornecedores

Dim objNode As Node
Dim lErro As Long
Dim sKeyFilial As String, objFornecedor As New ClassFornecedor
Dim sMatriz As String, iAchou As Integer

On Error GoTo Erro_Inserir_Arvore_FilialFornecedor

    sMatriz = KEY_CARACTER & CStr(objFilialFornecedor.lCodFornecedor)
    
    iAchou = 0

    For Each objNode In colNodes

        If objNode.Key = sMatriz Then

            iAchou = 1
            Exit For

        End If

    Next

    'se não achou a matriz na arvore
    If iAchou = 0 Then
    
        'obtem o nome do fornecedor
        objFornecedor.lCodigo = objFilialFornecedor.lCodFornecedor
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12729 Then Error 56714
        If lErro <> SUCESSO Then Error 56715
        
        'inclui o fornecedor na arvore
        Set objNode = colNodes.Add(, tvwLast, sMatriz)
        objNode.Text = objFornecedor.sNomeReduzido
    
    End If
    
    sKeyFilial = KEY_CARACTER & CStr(objFilialFornecedor.lCodFornecedor) & SEPARADOR & CStr(objFilialFornecedor.iCodFilial)
    
    Set objNode = colNodes.Add(colNodes.Item(sMatriz), tvwChild, sKeyFilial)

    objNode.Text = objFilialFornecedor.sNome

    Inserir_Arvore_FilialFornecedor = SUCESSO
     
    Exit Function
    
Erro_Inserir_Arvore_FilialFornecedor:

    Inserir_Arvore_FilialFornecedor = Err
     
    Select Case Err
          
        Case 56714
        
        Case 56715
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", Err, objFornecedor.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160296)
     
    End Select
     
    Exit Function
    
End Function

Private Sub CodFornecedor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodFornecedor_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodFornecedor, iAlterado)

End Sub

Private Sub CodFornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objFornecedor As New ClassFornecedor
Dim iIndice As Integer

On Error GoTo Erro_CodFornecedor_Validate

    'verifica se foi preenchido o campo Codigo Fornecedor
    If Len(Trim(CodFornecedor.Text)) = 0 Then Exit Sub

    lErro = Long_Critica(CodFornecedor.Text)
    If lErro <> SUCESSO Then Error 12942

    objFornecedor.lCodigo = CLng(CodFornecedor.Text)

    'Le o Fornecedor
    lErro = CF("Fornecedor_Le", objFornecedor)
    If lErro <> SUCESSO And lErro <> 12729 Then Error 12941

    For iIndice = 0 To 3
        FornecedorLabel(iIndice).Caption = objFornecedor.sNomeReduzido
    Next

    'Verifica se existe o Fornecedor
    If lErro <> SUCESSO Then

        'envia aviso perguntando se deseja cadastrar novo Fornecedor
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FORNECEDOR")

        If vbMsgRes = vbYes Then
            Call Chama_Tela("Fornecedores", objFornecedor)
        Else
            CodFornecedor.SetFocus
        End If

    End If

    Exit Sub

Erro_CodFornecedor_Validate:

    Cancel = True


    Select Case Err

    Case 12941, 12942

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160297)

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

'   Na primeira vez que ocorre Validate de CGC/CPF não formata com
'os pontos e os traços;***

Dim lErro As Long

On Error GoTo Erro_CGC_Validate

    If Len(Trim(CGC.Text)) = 0 Then Exit Sub

    Select Case Len(Trim(CGC.Text))

        Case STRING_CPF 'CPF

            lErro = Cpf_Critica(CGC.Text)
            If lErro <> SUCESSO Then Error 12944

            CGC.Format = "000\.000\.000-00; ; ; "
            CGC.Text = CGC.Text

        Case STRING_CGC  'CGC

            lErro = Cgc_Critica(CGC.Text)
            If lErro <> SUCESSO Then Error 12945

            CGC.Format = "00\.000\.000\/0000-00; ; ; "
            CGC.Text = CGC.Text

        Case Else

            Error 12946

    End Select

    Exit Sub

Erro_CGC_Validate:

    Cancel = True


    Select Case Err

        Case 12944, 12945

        Case 12946
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC_CPF", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160298)

    End Select

    Exit Sub

End Sub

Private Sub InscricaoEstadual_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub InscricaoMunicipal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaContabil_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaFornConsig_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaContabil_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaMascarada As String
Dim sContaFormatada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ContaContabil_Validate

    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica_Modulo", ContaContabil.Text, ContaContabil.ClipText, objPlanoConta, MODULO_CONTASAPAGAR)
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then Error 39807

    If lErro = SUCESSO Then

        sContaFormatada = objPlanoConta.sConta

        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 39808

        ContaContabil.PromptInclude = False
        ContaContabil.Text = sContaMascarada
        ContaContabil.PromptInclude = True

    'se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then

        lErro = CF("Conta_Critica", ContaContabil.Text, sContaFormatada, objPlanoConta, MODULO_CONTASAPAGAR)
        If lErro <> SUCESSO And lErro <> 5700 Then Error 12949

        If lErro = 5700 Then Error 12950

    End If

    Exit Sub

Erro_ContaContabil_Validate:

    Cancel = True


    Select Case Err

        Case 12949, 39807

        Case 12950
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", Err, ContaContabil.Text)

        Case 39808
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160299)

    End Select

    Exit Sub

End Sub

Private Sub ContaFornConsig_Validate(Cancel As Boolean)
'Função criada em 27/06/2001 por Luiz Gustavo de Freitas Nogueira
'Função criada em função da inclusão do campo ContaFornConsig na tela

Dim lErro As Long
Dim sContaMascarada As String
Dim sContaFormatada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ContaFornConsig_Validate

    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica_Modulo", ContaFornConsig.Text, ContaFornConsig.ClipText, objPlanoConta, MODULO_CONTASAPAGAR)
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 90551

    If lErro = SUCESSO Then

        sContaFormatada = objPlanoConta.sConta

        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then gError 90552

        ContaFornConsig.PromptInclude = False
        ContaFornConsig.Text = sContaMascarada
        ContaFornConsig.PromptInclude = True

    'se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then

        lErro = CF("Conta_Critica", ContaFornConsig.Text, sContaFormatada, objPlanoConta, MODULO_CONTASAPAGAR)
        If lErro <> SUCESSO And lErro <> 5700 Then gError 90553

        If lErro = 5700 Then gError 90554

    End If

    Exit Sub

Erro_ContaFornConsig_Validate:

    Cancel = True

    Select Case gErr

        Case 90553, 90551

        Case 90554
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", gErr, ContaFornConsig.Text)

        Case 90552
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160300)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodFilial_Click()

Dim colSelecao As Collection
Dim objFilialFornecedor As New ClassFilialFornecedor

    'Preenche com o fornecedor da tela
    objFilialFornecedor.lCodFornecedor = StrParaLong(CodFornecedor.Text)
    objFilialFornecedor.iCodFilial = StrParaInt(CodFilial.Text)
    
    'Chama Tela FiliaisFornecedoresLista
    Call Chama_Tela("FiliaisFornecedoresLista", colSelecao, objFilialFornecedor, objEventoFilialFornecedor)

End Sub

Private Sub LabelCodFornecedor_Click()

Dim colSelecao As Collection
Dim objFornecedor As New ClassFornecedor

    'Preenche com o Fornecedor da tela
    objFornecedor.lCodigo = StrParaLong(CodFornecedor.Text)
    
    'Chama Tela FornecedorLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

End Sub

Private Sub Nome_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Nome_Validate(Cancel As Boolean)

    FilialLabel(0).Caption = Nome.Text
    FilialLabel(1).Caption = Nome.Text
    FilialLabel(2).Caption = Nome.Text

End Sub

Private Sub objEventoFilialFornecedor_evSelecao(obj1 As Object)

Dim objFilialFornecedor As ClassFilialFornecedor
Dim objFilialFornecedorEstatistica As New ClassFilialFornecedorEst
Dim lErro As Long

On Error GoTo Erro_objEventoFilialFornecedor_evSelecao

    Set objFilialFornecedor = obj1

    lErro = CF("FilialFornecedor_Le_Estendida", objFilialFornecedor, objFilialFornecedorEstatistica)
    If lErro <> SUCESSO And lErro <> 52702 Then gError 82704

    'Verifica se Filial existe
    If lErro <> SUCESSO Then gError 82705

    lErro = Exibe_Dados_FilialFornecedor(objFilialFornecedor, objFilialFornecedorEstatistica)
    If lErro <> SUCESSO Then gError 82706

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoFilialFornecedor_evSelecao:

    Select Case gErr

        Case 82704, 82706

        Case 82705
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objFilialFornecedor.iCodFilial, objFilialFornecedor.lCodFornecedor)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160302)

    End Select

    Exit Sub

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor
Dim bCancel As Boolean

    Set objFornecedor = obj1

    'Preenche campo
    CodFornecedor.Text = CStr(objFornecedor.lCodigo)

    Call CodFornecedor_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Private Sub IdEstrangeiro_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_Load()

'*** FUNÇÃO ALTERADA EM 27/06/2001 ***
'Incluído código para incialização do campo ContaFornConsig
'*************************************

Dim iIndice As Integer
Dim lErro As Long
Dim colCodigo As New Collection
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome
Dim vCodigo As Variant
Dim sMascaraConta As String
Dim objFilialEmpresa As New AdmFiliais
Dim objTela As Object

On Error GoTo Erro_Form_Load
    
    iFrameAtual = 1
    iAlterado = 0
    
    Set objGridCategoria = New AdmGrid

    'Inicializa o Grid de Categoria
    lErro = Inicializa_Grid_Categoria(objGridCategoria)
    If lErro <> SUCESSO Then gError 91242

    'Carrega a combobox de Categoria Cliente
    lErro = Carrega_ComboCategoriaFornecedor()
    If lErro <> SUCESSO Then gError 91243
    
    'Carrega a List da Combo de Tipos de Cobrança
    lErro = Carrega_TipoCobranca()
    If lErro <> SUCESSO Then gError 106006

    Set objEventoFilialFornecedor = New AdmEvento
    Set objEventoFornecedor = New AdmEvento

    Set objTela = Me
    lErro = gobjTabEnd.Inicializa(objTela, TabEnd(0))
    If lErro <> SUCESSO Then gError 12915

    'Verifica se o modulo de contabilidade esta ativo antes das inicializacoes
    If (gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO) Then
        
        'Conta de fornecedores
        Set objEventoContaContabil = New AdmEvento
        
        'Inicializa propriedade Mask de ContaContabil
        lErro = MascaraConta(sMascaraConta)
        If lErro <> SUCESSO Then gError 12917
    
        ContaContabil.Mask = sMascaraConta
        
        'Alteracao Daniel
        'Conta de Estoque/Despesa
        Set objEventoContaDespesa = New AdmEvento
        
        'Inicializa propriedade Mask de ContaDespesa
        lErro = MascaraConta(sMascaraConta)
        If lErro <> SUCESSO Then gError 106003
    
        ContaDespesa.Mask = sMascaraConta
        'Fim da Alteracao Daniel
        
        '*** Incluído por Luiz Gustavo em 27/06/2001 ***
        'Conta de fornecedores em consignação
        Set objEventoContaFornConsig = New AdmEvento
        
        'Inicializa propriedade Mask de ContaFornConsig
        lErro = MascaraConta(sMascaraConta)
        If lErro <> SUCESSO Then gError 90556
    
        ContaFornConsig.Mask = sMascaraConta
        '***********************************************
    
    Else
       
        'Incluido a inicialização da máscara para não dar erro na gravação de clientes com conta mas que o módulo de contabilidade foi desabilitado
        lErro = MascaraConta(sMascaraConta)
        If lErro <> SUCESSO Then gError 12917
       
        ContaContabil.Mask = sMascaraConta
        ContaDespesa.Mask = sMascaraConta
        ContaFornConsig.Mask = sMascaraConta
       
        'Conta de fornecedores
        ContaContabil.Enabled = False
        ContaContabilLabel.Enabled = False
        
        'Alteracao Daniel
        ContaDespesa.Enabled = False
        LabelContaDespesa.Enabled = False
        'Fim da Alteracao Daniel
        
        '*** Incluído por Luiz Gustavo em 27/06/2001 ***
        'Conta de fornecedores em consignação
        ContaFornConsig.Enabled = False
        ContaFornConsig.Enabled = False
        '***********************************************
        
    End If
    
    'Carrega a Combo TipoFrete
    lErro = Carrega_TipoFrete()
    If lErro <> SUCESSO Then gError 207476
    
    'Prepara as Combos  do Tab de Inscricao
    lErro = Inicializa_Tab_Inscricao()
    If lErro <> SUCESSO Then gError 207397
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 12915, 12916, 90556, 90557, 91242, 91243, 106003, 106006, 207397, 207476

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160304)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Saida_Celula_Categoria(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Categoria do grid que está deixando de ser a corrente

'******** Janaina ********

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
    
    If GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL) <> ComboCategoriaFornecedor.Text Then

        If Len(Trim(ComboCategoriaFornecedor)) > 0 Then

            'Preenche o objeto com a Categoria
            objCategoriaFornecedor.sCategoria = ComboCategoriaFornecedor.Text

            'Lê Categoria De Cliente no BD
            lErro = CategoriaFornecedor_Le_Todos(colCategorias)
            If lErro <> SUCESSO And lErro <> 28847 Then gError 91244

            'Categoria não está cadastrada
            If lErro = 28847 Then gError 91245

            'Verifica se já existe a categoria no Grid
            If GridCategoria.Row > 0 Then

                iIndice1 = GridCategoria.Row

                For iIndice = 1 To objGridCategoria.iLinhasExistentes

                    If iIndice <> iIndice1 Then

                        'Se já existe a categoria, erro
                        If GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL) = ComboCategoriaFornecedor.Text Then gError 91246

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
    If lErro <> SUCESSO Then Error 91247

    If GridCategoria.Row - GridCategoria.FixedRows = objGridInt.iLinhasExistentes Then

        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1

    End If

    Saida_Celula_Categoria = SUCESSO

    Exit Function

Erro_Saida_Celula_Categoria:

    Saida_Celula_Categoria = gErr

    Select Case gErr

        Case 91244
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 91245 'Categoria não está cadastrada

            'Perguntar se deseja criar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DESEJA_CRIAR_CATEGORIAFORNECEDOR")

            If vbMsgRes = vbYes Then

                'Chama a Tela "CategoriaFornecedor"
                Call Chama_Tela("CategoriaFornec", objCategoriaFornecedor)
                
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            Else

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 91246
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_JA_SELECIONADA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 91247
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160305)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Valor(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Item do grid que está deixando de ser a corrente

'******** Janaina ********

Dim lErro As Long
Dim objCategoriaFornecedor As New ClassCategoriaFornecedor
Dim objCategoriaFornecedorItem As New ClassCategoriaFornItem
Dim colItens As New Collection
Dim ColItensCtegoria As New Collection

Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridCategoria.objControle = ComboCategoriaFornecedorItem

    If Len(Trim(ComboCategoriaFornecedorItem)) > 0 Then

        'se o campo de categoria estiver vazio ==> erro
        If Len(GridCategoria.TextMatrix(GridCategoria.Row, GRID_CATEGORIA_COL)) = 0 Then gError 91248

        'Tenta selecionar na combo
        lErro = Combo_Item_Igual(ComboCategoriaFornecedorItem)
        If lErro <> SUCESSO Then Error 91249

            'Preenche o objeto com a Categoria
            objCategoriaFornecedorItem.sCategoria = GridCategoria.TextMatrix(GridCategoria.Row, GRID_CATEGORIA_COL)
            objCategoriaFornecedorItem.sItem = ComboCategoriaFornecedorItem.Text

            'Lê Item De Categoria De Fornecedor no BD
            lErro = CategoriaFornecedor_Le_Itens(ColItensCtegoria)
            If lErro <> SUCESSO And lErro <> 91180 Then Error 91250

            'Item da Categoria não está cadastrado
            If lErro = 91180 Then Error 91251

        End If

        If GridCategoria.Row - GridCategoria.FixedRows = objGridInt.iLinhasExistentes Then

            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1

        End If

    'End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 91252

    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = gErr

    Select Case gErr

        Case 91248
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_CATEGORIA_NAO_PREENCHIDA", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 91249

        Case 91250
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 91251 'Item da Categoria não está cadastrado

            'Perguntar se deseja criar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DESEJA_CRIAR_CATEGORIAFORNECEDORITEM")

            If vbMsgRes = vbYes Then

                'Preenche o objeto com a Categoria
                objCategoriaFornecedor.sCategoria = GridCategoria.TextMatrix(GridCategoria.Row, GRID_CATEGORIA_COL)

                'Chama a Tela "CategoriaFornecedor"
                Call Chama_Tela("CategoriaFornecedor", objCategoriaFornecedor, objCategoriaFornecedorItem.sItem)
                
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            Else

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 91252
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160306)

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
                If lErro <> SUCESSO Then gError 91253

            Case GRID_VALOR_COL

                'Chama SaidaCelula de Valor
                lErro = Saida_Celula_Valor(objGridInt)
                If lErro <> SUCESSO Then gError 91254

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 91255

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 91253, 91254

        Case 91255
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160307)

    End Select

    Exit Function

End Function

Private Function Carrega_ComboCategoriaFornecedor() As Long
'Carrega as Categorias na Combobox

'******** Janaina ********

Dim lErro As Long
Dim colCategorias As New Collection
Dim objCategoriaFornecedor As New ClassCategoriaFornecedor

On Error GoTo Erro_Carrega_ComboCategoriaFornecedor
  
    'Lê o código e a descrição de todas as categorias
    lErro = CategoriaFornecedor_Le_Todos(colCategorias)
    If lErro <> SUCESSO And lErro <> 68486 Then gError 91256

    For Each objCategoriaFornecedor In colCategorias

        'Insere na combo CategoriaCliente
        ComboCategoriaFornecedor.AddItem objCategoriaFornecedor.sCategoria

    Next

    Carrega_ComboCategoriaFornecedor = SUCESSO

    Exit Function

Erro_Carrega_ComboCategoriaFornecedor:

    Carrega_ComboCategoriaFornecedor = gErr

    Select Case gErr

        Case 91256

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160308)

    End Select

    Exit Function

End Function

Function Trata_Parametros(Optional objFilial As ClassFilialFornecedor) As Long

Dim lErro As Long
Dim objFilialFornecedorEstatistica As New ClassFilialFornecedorEst

On Error GoTo Erro_Trata_Parametros

    If Not (objFilial Is Nothing) Then

        If objFilial.lCodFornecedor <> 0 Then

            'foi passado a chave da Filial
            If objFilial.iCodFilial <> 0 Then

                lErro = CF("FilialFornecedor_Le_Estendida", objFilial, objFilialFornecedorEstatistica)
                If lErro <> SUCESSO And lErro <> 52702 Then Error 12924

                'Se a Filial não existir
                If lErro <> SUCESSO Then

                    Call Limpa_Tela_FiliaisFornecedores
                    CodFornecedor.Text = CStr(objFilial.lCodFornecedor)
                    CodFilial.Text = CStr(objFilial.iCodFilial)

                Else

                    lErro = Exibe_Dados_FilialFornecedor(objFilial, objFilialFornecedorEstatistica)
                    If lErro <> SUCESSO Then Error 12925

                End If

            Else 'apenas o código do Fornecedor foi passado

                Call Limpa_Tela_FiliaisFornecedores
                CodFornecedor.Text = CStr(objFilial.lCodFornecedor)

            End If

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 12924, 12925

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160309)

    End Select

    iAlterado = 0

    Exit Function

End Function

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
                Parent.HelpContextID = IDH_FILIAL_FORN_IDENT
                
            Case TAB_Complemento
                Parent.HelpContextID = IDH_FILIALFORN_DADOS_FIN '??? Criar IDH para esse tab
                
            Case TAB_DadosFinanceiros
                Parent.HelpContextID = IDH_FILIALFORN_DADOS_FIN
            
            Case TAB_Inscricoes
                Parent.HelpContextID = IDH_FILIALFORN_INSCRICOES
            
            Case TAB_Endereco
                Parent.HelpContextID = IDH_FILIALFORN_ENDERECO
            
            Case TAB_Estatisticas
                Parent.HelpContextID = IDH_FILIALFORN_ESTATISTICAS
        
        End Select
    
    End If

End Sub

Private Function Le_Dados_FilialFornecedor(objFilialFornecedor As ClassFilialFornecedor) As Long
'Le os dados que estao na tela de FiliaisFornecedores e coloca em objFilialFornecedor

'*** FUNÇÃO ALTERADA EM 27/06/2001 ***
'Incluído código para incialização do campo ContaFornConsig
'*************************************

Dim lErro As Long
Dim iContaPreenchida As Integer
Dim sConta As String
Dim sContaFornConsig As String
Dim iIndice As Integer
Dim objFilialFornCategoria  As ClassFilialFornCategoria
Dim objFornecedor As ClassFornecedor

On Error GoTo Erro_Le_Dados_FilialFornecedor

    'IDENTIFICACAO :

    If Len(Trim(CodFornecedor.Text)) > 0 Then objFilialFornecedor.lCodFornecedor = CLng(CodFornecedor.Text)

    If Len(Trim(CodFilial.Text)) > 0 Then objFilialFornecedor.iCodFilial = CodFilial.Text

    objFilialFornecedor.sNome = Nome.Text

    objFilialFornecedor.sCgc = CGC.Text

    objFilialFornecedor.sInscricaoEstadual = InscricaoEstadual.Text
    objFilialFornecedor.sInscricaoSuframa = Trim(InscricaoSuframa.ClipText)

    objFilialFornecedor.sInscricaoMunicipal = InscricaoMunicipal.Text

    objFilialFornecedor.sObservacao = Observacao.Text
    objFilialFornecedor.iRegimeTributario = RegimeTributario.ItemData(RegimeTributario.ListIndex)
    objFilialFornecedor.sIdEstrangeiro = Trim(IdEstrangeiro.Text)

    If IEIsento.Value = vbChecked Then
        objFilialFornecedor.iIEIsento = MARCADO
    Else
        objFilialFornecedor.iIEIsento = DESMARCADO
    End If
    If IENaoContrib.Value = vbChecked Then
        objFilialFornecedor.iIENaoContrib = MARCADO
    Else
        objFilialFornecedor.iIENaoContrib = DESMARCADO
    End If
    
    'Alteracao Daniel
    objFilialFornecedor.sInscricaoINSS = InscricaoINSS.Text
    objFilialFornecedor.iGeraCredICMS = GeraICMS.Value
    If TipoFrete.ListIndex <> -1 Then
        objFilialFornecedor.iTipoFrete = TipoFrete.ItemData(TipoFrete.ListIndex)
    Else
        objFilialFornecedor.iTipoFrete = -1
    End If
        
    objFilialFornecedor.iTipoCobranca = Codigo_Extrai(TipoCobranca.Text)

    'Preenche uma coleção com todas as linhas "existentes" do grid de categorias
    For iIndice = 1 To objGridCategoria.iLinhasExistentes

        'Verifica se a Categoria foi preenchida
        If Len(Trim(GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL))) <> 0 Then

            Set objFilialFornCategoria = New ClassFilialFornCategoria
           
            objFilialFornCategoria.lFornecedor = objFilialFornecedor.lCodFornecedor
            objFilialFornCategoria.sCategoria = GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL)
            objFilialFornCategoria.sItem = GridCategoria.TextMatrix(iIndice, GRID_VALOR_COL)

            objFilialFornecedor.colCategoriaItem.Add objFilialFornCategoria

        End If

    Next
    
'****************************
   
   'DADOS FINANCEIROS:

    objFilialFornecedor.sContaCorrente = ContaCorrente.Text

    objFilialFornecedor.sAgencia = Agencia.Text

    If Len(Trim(Banco.Text)) > 0 Then objFilialFornecedor.iBanco = CInt(Banco.Text)

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaContabil.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 12966

    If iContaPreenchida = CONTA_VAZIA Then
        objFilialFornecedor.sContaContabil = ""
    Else
        objFilialFornecedor.sContaContabil = sConta
    End If

    'Alteracao Daniel
    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaDespesa.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 12966

    If iContaPreenchida = CONTA_VAZIA Then
        objFilialFornecedor.sContaDespesa = ""
    Else
        objFilialFornecedor.sContaDespesa = sConta
    End If
    'Fim da Alteracao Daniel
    
    '*** Incluído por Luiz Gustavo em 27/06/2001 ***
    sContaFornConsig = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaFornConsig.Text, sContaFornConsig, iContaPreenchida)
    If lErro <> SUCESSO Then gError 90555

    If iContaPreenchida = CONTA_VAZIA Then
        objFilialFornecedor.sContaFornConsig = ""
    Else
        objFilialFornecedor.sContaFornConsig = sContaFornConsig
    End If
    '*************************************************

    Le_Dados_FilialFornecedor = SUCESSO

    Exit Function

Erro_Le_Dados_FilialFornecedor:

    Le_Dados_FilialFornecedor = gErr

    Select Case gErr

        Case 90555

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160310)

    End Select

    Exit Function

End Function

Function Exibe_Dados_FilialFornecedor(objFilialFornecedor As ClassFilialFornecedor, objFilialFornecedorEstatistica As ClassFilialFornecedorEst) As Long

' *** ALTERADA EM 27/06/2001 POR LUIZ GUSTAVO DE FREITAS NOGUEIRA ***
' Alteração feita em função da inclusão do campo ContaFornConsig
' Portanto, foi incluído código para exibir esse campo
' *******************************************************************

Dim objEndereco As New ClassEndereco
Dim objFornecedor As New AdmlCodigoNome
Dim colCategoriaItem As New Collection
Dim iIndice As Integer
Dim lErro As Long
Dim sContaEnxuta As String
Dim colEnderecos As New Collection

On Error GoTo Erro_Exibe_Dados_FilialFornecedor

    'pesquisa a primeira compra
    lErro = CF("Pesquisa_Primeira_Compra_FornFil", objFilialFornecedorEstatistica)
    If lErro <> SUCESSO Then gError 89404

    'pesquisa a ultima compra
    lErro = CF("Pesquisa_Ultima_Compra_FornFil", objFilialFornecedorEstatistica)
    If lErro <> SUCESSO Then gError 89405

    'pesquisa a ultima nota fiscal de Devolução de Compra da FilialEmpresa (data de emissao mais recente) para o Fornecedor/FilialForn em questão.
    lErro = CF("Pesquisa_Ultima_Devolucao_FornFil", objFilialFornecedorEstatistica)
    If lErro <> SUCESSO Then gError 89416

    objFornecedor.lCodigo = objFilialFornecedor.lCodFornecedor

    'Le o Nome Reduzido do Fornecedor relacionado a Filial
    lErro = CF("Fornecedor_Le_NomeRed", objFornecedor)
    If lErro <> SUCESSO Then gError 12930

    For iIndice = 0 To 4
        FornecedorLabel(iIndice).Caption = objFornecedor.sNome
    Next

    'IDENTIFICACAO :

    CodFornecedor.Text = CStr(objFilialFornecedor.lCodFornecedor)

    CodFilial.Text = CStr(objFilialFornecedor.iCodFilial)

    Nome.Text = objFilialFornecedor.sNome


    'DADOS FINANCEIROS :

    FilialLabel(0).Caption = objFilialFornecedor.sNome

    ContaCorrente.Text = objFilialFornecedor.sContaCorrente

    Agencia.Text = objFilialFornecedor.sAgencia

    If objFilialFornecedor.iBanco = 0 Then
        Banco.Text = ""
    Else
        Banco.Text = CStr(objFilialFornecedor.iBanco)
    End If

    'Conta de Fornecedores
    ContaContabil.PromptInclude = False

    If Len(Trim(objFilialFornecedor.sContaContabil)) = 0 Then
        ContaContabil.Text = objFilialFornecedor.sContaContabil
    Else
        lErro = Mascara_RetornaContaEnxuta(objFilialFornecedor.sContaContabil, sContaEnxuta)
        If lErro <> SUCESSO Then gError 12931

        ContaContabil.Text = sContaEnxuta
    End If

    ContaContabil.PromptInclude = True
    
    'Alteracao Daniel
    ContaDespesa.PromptInclude = False

    If Len(Trim(objFilialFornecedor.sContaDespesa)) = 0 Then
        ContaDespesa.Text = objFilialFornecedor.sContaDespesa
    Else
        lErro = Mascara_RetornaContaEnxuta(objFilialFornecedor.sContaDespesa, sContaEnxuta)
        If lErro <> SUCESSO Then gError 106004

        ContaDespesa.Text = sContaEnxuta
    End If

    ContaDespesa.PromptInclude = True
    'Fim da Alteracao
    
    '*** Incluído por Luiz Gustavo em 27/06/2001 ***
    'Conta de Fornecedores em consignação
    ContaFornConsig.PromptInclude = False

    If Len(Trim(objFilialFornecedor.sContaFornConsig)) = 0 Then
        ContaFornConsig.Text = objFilialFornecedor.sContaFornConsig
    Else
        lErro = Mascara_RetornaContaEnxuta(objFilialFornecedor.sContaFornConsig, sContaEnxuta)
        If lErro <> SUCESSO Then gError 90548

        ContaFornConsig.Text = sContaEnxuta
    End If
    
    ContaFornConsig.PromptInclude = True
    '***********************************************

    'INSCRICOES:

    FilialLabel(1).Caption = objFilialFornecedor.sNome

    CGC.Text = objFilialFornecedor.sCgc
    Call CGC_Validate(bSGECancelDummy)

    InscricaoEstadual.Text = objFilialFornecedor.sInscricaoEstadual
    Call Trata_IE
    If objFilialFornecedor.iIEIsento = MARCADO Then
        IEIsento.Value = vbChecked
    Else
        IEIsento.Value = vbUnchecked
    End If
    If objFilialFornecedor.iIENaoContrib = MARCADO Then
        IENaoContrib.Value = vbChecked
    Else
        IENaoContrib.Value = vbUnchecked
    End If
    
    InscricaoMunicipal.Text = objFilialFornecedor.sInscricaoMunicipal
    InscricaoSuframa.Text = objFilialFornecedor.sInscricaoSuframa

    Observacao.Text = objFilialFornecedor.sObservacao
    IdEstrangeiro.Text = objFilialFornecedor.sIdEstrangeiro
    
    'Alteracao Daniel
    InscricaoINSS.Text = objFilialFornecedor.sInscricaoINSS

    'Busca no ItemData da Combo de Tipo de Cobranca o código
    For iIndice = 0 To RegimeTributario.ListCount - 1
        If RegimeTributario.ItemData(iIndice) = objFilialFornecedor.iRegimeTributario Then
            RegimeTributario.ListIndex = iIndice
            Exit For
        End If
    Next

    'ENDERECOS :

    FilialLabel(2).Caption = objFilialFornecedor.sNome

    objEndereco.lCodigo = objFilialFornecedor.lEndereco

    'Le os dados do Endereco da Filial Fornecedor
    lErro = CF("Endereco_Le", objEndereco)
    If lErro <> SUCESSO Then gError 12932
    
    colEnderecos.Add objEndereco
    
    lErro = gobjTabEnd.Traz_Endereco_Tela(colEnderecos)
    If lErro <> SUCESSO Then gError 12932
    
    'ESTATÍSTICAS :
    
    FilialLabel(3).Caption = objFilialFornecedor.sNome
    
    SaldoTitulos.Caption = Format(objFilialFornecedorEstatistica.dSaldoTitulos, "Standard")
    NumeroCompras.Caption = Format(objFilialFornecedorEstatistica.lNumeroCompras, FORMATO_INTEIRO)
    
    If objFilialFornecedorEstatistica.lNumeroCompras > 0 Then
        MediaCompra.Caption = Format(objFilialFornecedorEstatistica.dValorAcumuladoCompras / objFilialFornecedorEstatistica.lNumeroCompras, "Standard")
    Else
        MediaCompra.Caption = CStr(0)
    End If
    
    ValorAcumuladoCompras.Caption = Format(objFilialFornecedorEstatistica.dValorAcumuladoCompras, "Standard")
    
    If objFilialFornecedorEstatistica.dtDataPrimeiraCompra = DATA_NULA Then
        DataPrimeiraCompra.Caption = ""
    Else
        DataPrimeiraCompra.Caption = Format(objFilialFornecedorEstatistica.dtDataPrimeiraCompra, "dd/mm/yyyy")
        
    End If
    
    If objFilialFornecedorEstatistica.dtDataUltimaCompra = DATA_NULA Then
        DataUltimaCompra.Caption = ""
    Else
        DataUltimaCompra.Caption = Format(objFilialFornecedorEstatistica.dtDataUltimaCompra, "dd/mm/yyyy")
    End If

    If objFilialFornecedorEstatistica.lItensPedidosRecebidos > 0 Then
        MediaAtraso.Caption = Format(objFilialFornecedorEstatistica.lAtrasoAcumulado / objFilialFornecedorEstatistica.lItensPedidosRecebidos, FORMATO_MEDIA_ATRASO)
    Else
        MediaAtraso.Caption = CStr(0)
    End If
    
'    MaiorAtraso.Caption = CStr(objFilialFornecedorEstatistica.lMaiorAtraso)

    ValorAcumuladoDevolucoes.Caption = Format(objFilialFornecedorEstatistica.dValorAcumuladoDevolucoes, "Standard")
    NumTotalDevolucoes.Caption = Format(objFilialFornecedorEstatistica.lNumTotalDevolucoes, FORMATO_INTEIRO)

    If objFilialFornecedorEstatistica.dtDataUltDevolucao = DATA_NULA Then
        DataUltDevolucao.Caption = ""
    Else
        DataUltDevolucao.Caption = Format(objFilialFornecedorEstatistica.dtDataUltDevolucao, "dd/mm/yyyy")
    End If

    '********** Janaina ***********
    
   'Lê as categoria deste fornecedor
    objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
    'objFilialFornecedor.iCodFilial = FILIAL_MATRIZ
    
    lErro = FilialFornecedor_Le_Categorias(objFilialFornecedor, colCategoriaItem)
    If lErro <> SUCESSO Then gError 91262

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
    GeraICMS.Value = objFilialFornecedor.iGeraCredICMS
    
    If objFilialFornecedor.iTipoFrete <> -1 Then
    
        For iIndice = 0 To TipoFrete.ListCount - 1
            If TipoFrete.ItemData(iIndice) = objFilialFornecedor.iTipoFrete Then
                TipoFrete.ListIndex = iIndice
                Exit For
            End If
        Next
    
    Else
    
        TipoFrete.ListIndex = -1
    End If
    
    'Busca no ItemData da Combo de Tipo de Cobranca o código
    For iIndice = 0 To TipoCobranca.ListCount - 1
        If TipoCobranca.ItemData(iIndice) = objFilialFornecedor.iTipoCobranca Then
            TipoCobranca.ListIndex = iIndice
            Exit For
        End If
    Next
    'Fim da Alteracao Daniel
    
    iAlterado = 0

    Exibe_Dados_FilialFornecedor = SUCESSO

    Exit Function

Erro_Exibe_Dados_FilialFornecedor:

    Exibe_Dados_FilialFornecedor = gErr

    Select Case gErr

        Case 12930, 12932, 12933, 89404, 89405, 89416, 91262

        Case 12931
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objFilialFornecedor.sContaContabil)

        Case 90548
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objFilialFornecedor.sContaFornConsig)
            
        Case 106004
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objFilialFornecedor.sContaDespesa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160311)

    End Select

    Exit Function

End Function

'RETORNO PLANO_CONTA :

Private Sub objEventoContaContabil_evSelecao(obj1 As Object)

'   Retirar o iAlterado = 0 de objEventoContaContabil; ***

Dim objPlanoConta As ClassPlanoConta
Dim lErro As Long
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaContabil_evSelecao

    Set objPlanoConta = obj1

    If objPlanoConta.sConta = "" Then

        ContaContabil.Text = ""

    Else

        ContaContabil.PromptInclude = False

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then Error 12608

        ContaContabil.Text = sContaEnxuta

        ContaContabil.PromptInclude = True

    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case Err

        Case 12608
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160312)

    End Select

    Exit Sub

End Sub

Private Sub objEventoContaFornConsig_evSelecao(obj1 As Object)
'Função criada em 27/06/2001 por Luiz Gustavo de Freitas Nogueira
'Função criada em função da inclusão do campo ContaFornConsig na tela

Dim objPlanoConta As ClassPlanoConta
Dim lErro As Long
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaFornConsig_evSelecao

    Set objPlanoConta = obj1

    If objPlanoConta.sConta = "" Then

        ContaFornConsig.Text = ""

    Else

        ContaFornConsig.PromptInclude = False

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then gError 90549

        ContaFornConsig.Text = sContaEnxuta

        ContaFornConsig.PromptInclude = True

    End If

    Me.Show

    Exit Sub

Erro_objEventoContaFornConsig_evSelecao:

    Select Case gErr

        Case 90549
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160313)

    End Select

    Exit Sub

End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA
'""""""""""""""""""""""""""""""""""""""""""""""

'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objFilialFornecedor As New ClassFilialFornecedor

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "FiliaisFornecedores"

    'Le os dados da Tela FilialFornecedor
    lErro = Le_Dados_FilialFornecedor(objFilialFornecedor)
    If lErro <> SUCESSO Then Error 12761

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo

    colCampoValor.Add "CodFornecedor", objFilialFornecedor.lCodFornecedor, 0, "CodFornecedor"
    colCampoValor.Add "CodFilial", objFilialFornecedor.iCodFilial, 0, "CodFilial"
    colCampoValor.Add "Nome", objFilialFornecedor.sNome, STRING_FILIAL_FORNECEDOR_NOME, "Nome"
    colCampoValor.Add "Endereco", objFilialFornecedor.lEndereco, 0, "Endereco"
    colCampoValor.Add "CGC", objFilialFornecedor.sCgc, STRING_CGC, "CGC"
    colCampoValor.Add "InscricaoEstadual", objFilialFornecedor.sInscricaoEstadual, STRING_INSCR_EST, "InscricaoEstadual"
    colCampoValor.Add "InscricaoMunicipal", objFilialFornecedor.sInscricaoMunicipal, STRING_INSCR_MUN, "InscricaoMunicipal"
    colCampoValor.Add "ContaContabil", objFilialFornecedor.sContaContabil, STRING_CONTA, "ContaContabil"
    colCampoValor.Add "ContaFornConsig", objFilialFornecedor.sContaFornConsig, STRING_CONTA, "ContaFornConsig"
    colCampoValor.Add "Banco", objFilialFornecedor.iBanco, 0, "Banco"
    colCampoValor.Add "Agencia", objFilialFornecedor.sAgencia, STRING_AGENCIA, "Agencia"
    colCampoValor.Add "ContaCorrente", objFilialFornecedor.sContaCorrente, STRING_CONTA, "ContaCorrente"
    colCampoValor.Add "Observacao", objFilialFornecedor.sObservacao, STRING_FILIAL_FORNECEDOR_OBS, "Observacao"
    
    'Alteracao Daniel (incluir no browser)
    colCampoValor.Add "InscricaoINSS", objFilialFornecedor.sInscricaoINSS, STRING_INSCR_INSS, "InscricaoINSS"
    colCampoValor.Add "ContaDespesa", objFilialFornecedor.sContaDespesa, STRING_CONTA, "ContaDespesa"
    colCampoValor.Add "TipoCobranca", objFilialFornecedor.iTipoCobranca, 0, "TipoCobranca"
    colCampoValor.Add "TipoFrete", objFilialFornecedor.iTipoFrete, 0, "TipoFrete"
    'Fim da Alteracao Daniel

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 12761

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160314)

    End Select

    Exit Sub

End Sub

'Preenche os campos da tela com os correspondentes do BD
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objFilialFornecedorEstatistica As New ClassFilialFornecedorEst
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Tela_Preenche

    objFilialFornecedor.lCodFornecedor = colCampoValor.Item("CodFornecedor").vValor

    If objFilialFornecedor.lCodFornecedor <> 0 Then

        objFilialFornecedor.iCodFilial = colCampoValor.Item("CodFilial").vValor
        
        lErro = CF("FilialFornecedor_Le_Estendida", objFilialFornecedor, objFilialFornecedorEstatistica)
        If lErro <> SUCESSO And lErro <> 52702 Then gError 52703
        
        lErro = Exibe_Dados_FilialFornecedor(objFilialFornecedor, objFilialFornecedorEstatistica)
        If lErro <> SUCESSO Then Error 12960

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 12960, 52703 'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160315)

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
    
    Set objEventoContaContabil = Nothing
    Set objEventoContaFornConsig = Nothing
    Set objEventoFilialFornecedor = Nothing
    Set objEventoFornecedor = Nothing
    
    Call gobjTabEnd.Finaliza
    Set gobjTabEnd = Nothing
    
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_FILIAL_FORN_IDENT
    Set Form_Load_Ocx = Me
    Caption = "Filiais de Fornecedores"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "FiliaisFornecedores"
    
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
        If Me.ActiveControl Is ContaContabil Then
            Call ContaContabilLabel_Click
        ElseIf Me.ActiveControl Is ContaDespesa Then
            Call LabelContaDespesa_Click
        ElseIf Me.ActiveControl Is CodFornecedor Then
            Call LabelCodFornecedor_Click
        ElseIf Me.ActiveControl Is CodFilial Then
            Call LabelCodFilial_Click
        ElseIf Me.ActiveControl Is ContaFornConsig Then
            Call ContaFornConsigLabel_Click
        End If
    End If

End Sub

Private Sub FornecedorLabel_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(FornecedorLabel(Index), Source, X, Y)
End Sub

Private Sub FornecedorLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel(Index), Button, Shift, X, Y)
End Sub

Private Sub FilialLabel_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(FilialLabel(Index), Source, X, Y)
End Sub

Private Sub FilialLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialLabel(Index), Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
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

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
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

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
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

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
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

Private Sub Label28_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label28, Source, X, Y)
End Sub

Private Sub Label28_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label28, Button, Shift, X, Y)
End Sub

Private Sub Label25_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label25, Source, X, Y)
End Sub

Private Sub Label25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label25, Button, Shift, X, Y)
End Sub

Private Sub Label22_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label22, Source, X, Y)
End Sub

Private Sub Label22_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label22, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
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

'*** Fernando, favor transportar as funções abaixo para o Rotinas ***
Function FilialFornecedor_Le_Categorias(objFilialFornecedor As ClassFilialFornecedor, colCategoriaItem As Collection) As Long
'Lê na tabela de FilialFornecedorCategorias, todas as categorias da Filial de fornecedores (com o respectivo valor em cada categoria), guardando na coleção colCategoriaItem

'****** Janaina ******

Dim lErro As Long
Dim iIndice As Integer
Dim lComando As Long
Dim sCategoria As String
Dim sItem As String
Dim objFilialFornCategoria As ClassFilialFornCategoria

On Error GoTo Erro_FilialFornecedor_Le_Categorias

    'Inicializar comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 91262

    sCategoria = String(STRING_CATEGORIAFORNECEDOR_CATEGORIA, 0)
    sItem = String(STRING_CATEGORIAFORNECEDORITEM_ITEM, 0)

    'Pesquisa no BD as categoria da filial do fornecedor
    lErro = Comando_Executar(lComando, "SELECT Categoria, Item FROM FilialFornecedorCategorias WHERE Fornecedor = ? AND Filial = ?", sCategoria, sItem, objFilialFornecedor.lCodFornecedor, objFilialFornecedor.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then gError 91264

    'Verifica se alguma categoria existe
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 91265

    'Passa para a collection
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objFilialFornCategoria = New ClassFilialFornCategoria

        objFilialFornCategoria.lFornecedor = objFilialFornecedor.lCodFornecedor
        objFilialFornCategoria.iFilial = objFilialFornecedor.iCodFilial
        objFilialFornCategoria.sCategoria = sCategoria
        objFilialFornCategoria.sItem = sItem

        colCategoriaItem.Add objFilialFornCategoria

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 91266

    Loop

    'Fechamento comando
    Call Comando_Fechar(lComando)

    FilialFornecedor_Le_Categorias = SUCESSO

    Exit Function

Erro_FilialFornecedor_Le_Categorias:

    FilialFornecedor_Le_Categorias = gErr

    Select Case gErr

        Case 91263
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 91264, 91265, 91266
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALCLIENTECATEGORIAS2", gErr, objFilialFornCategoria.lFornecedor)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160316)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CategoriaFornecedor_Le_Itens(colItensCategoria As Collection) As Long
'Busca no BD Categoria de FornecedorItem

'******** Janaina ********

Dim lErro As Long
Dim lComando As Long
Dim sCategoria As String
Dim sDescricao As String
Dim sItem As String
Dim objCategoriaFornecedorItem As New ClassCategoriaFornItem

On Error GoTo Erro_CategoriaFornecedor_Le_Itens

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 91257

    sCategoria = String(STRING_CATEGORIAFORNECEDOR_CATEGORIA, 0)
    sDescricao = String(STRING_CATEGORIAFORNECEDOR_DESCRICAO, 0)
    sItem = String(STRING_CATEGORIAFORNECEDOR_ITEM, 0)
    
    objCategoriaFornecedorItem.sCategoria = GridCategoria.TextMatrix(GridCategoria.Row, GRID_CATEGORIA_COL)

    lErro = Comando_Executar(lComando, "SELECT descricao, item FROM CategoriaFornecedorItem where categoria = ? ", sDescricao, sItem, objCategoriaFornecedorItem.sCategoria)
    If lErro <> AD_SQL_SUCESSO Then gError 91258

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 91259

    'Se nao encontrar => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 91260
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objCategoriaFornecedorItem = New ClassCategoriaFornItem

        'Preenche objCategoriaFornecedorItem com o que foi lido do banco de dados
        objCategoriaFornecedorItem.sDescricao = sDescricao
        objCategoriaFornecedorItem.sItem = sItem
        
        colItensCategoria.Add objCategoriaFornecedorItem
        
        'Lê a próximo Categoria, Descicao e Item
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 91261

    Loop
    
    'Fecha o comando
    lErro = Comando_Fechar(lComando)

    CategoriaFornecedor_Le_Itens = SUCESSO

    Exit Function

Erro_CategoriaFornecedor_Le_Itens:

    CategoriaFornecedor_Le_Itens = gErr

    Select Case gErr

        Case 91257
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 91258, 91259
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATGORIAFORNECEDOR", gErr)

        Case 91260, 91261
            'Erro tratado na rotina chamada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160317)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function CategoriaFornecedor_Le_Todos(colCategorias As Collection) As Long
'Busca no BD Categoria de Fornecedor

Dim lErro As Long
Dim lComando As Long
Dim objCategoriaFornecedor As New ClassCategoriaFornecedor
Dim sCategoria As String
Dim sDescricao As String

On Error GoTo Erro_CategoriaFornecedor_Le_Todos

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 68483

    sCategoria = String(STRING_CATEGORIAFORNECEDOR_CATEGORIA, 0)
    sDescricao = String(STRING_CATEGORIAFORNECEDOR_DESCRICAO, 0)

    lErro = Comando_Executar(lComando, "SELECT categoria, descricao FROM CategoriaFornecedor", sCategoria, sDescricao)
    If lErro <> AD_SQL_SUCESSO Then gError 68484

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 68485

    'Se nao encontrar => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 68486

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objCategoriaFornecedor = New ClassCategoriaFornecedor
        
        'Preenche objCategoriaFornecedor com o que foi lido do banco de dados
        objCategoriaFornecedor.sCategoria = sCategoria
        objCategoriaFornecedor.sDescricao = sDescricao

        colCategorias.Add objCategoriaFornecedor
        
        'Lê a próximo Categoria e Descicao
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 75396

    Loop
    
    'Fecha o comando
    lErro = Comando_Fechar(lComando)

    CategoriaFornecedor_Le_Todos = SUCESSO

    Exit Function

Erro_CategoriaFornecedor_Le_Todos:

    CategoriaFornecedor_Le_Todos = gErr

    Select Case gErr

        Case 68483
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 68484, 68485
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATGORIAFORNECEDOR", gErr)

        Case 68486
            'Erro tratado na rotina chamada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160318)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function
'*** Fim do trecho que deve ser transportado para o Rotinas ***

'Alteracao Daniel
Private Sub GeraICMS_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub InscricaoINSS_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ContaDespesa_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub TipoCobranca_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub TipoFrete_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub LabelContaDespesa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelContaDespesa, Source, X, Y)
End Sub

Private Sub LabelContaDespesa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelContaDespesa, Button, Shift, X, Y)
End Sub

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160319)

    End Select

End Sub

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160320)

    End Select

End Sub

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160321)

    End Select

End Function

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
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160322)

    End Select

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
    If lErro <> SUCESSO Then gError 207474

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

        Case 207474

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207475)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207399)
        
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
