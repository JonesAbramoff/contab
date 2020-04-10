VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl FiliaisClientesOcx 
   ClientHeight    =   5370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9285
   KeyPreview      =   -1  'True
   ScaleHeight     =   5370
   ScaleWidth      =   9285
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4545
      Index           =   2
      Left            =   270
      TabIndex        =   9
      Top             =   690
      Visible         =   0   'False
      Width           =   8820
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
         Left            =   5475
         TabIndex        =   131
         Top             =   1455
         Value           =   1  'Checked
         Width           =   2820
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
         Left            =   4425
         TabIndex        =   108
         Top             =   1500
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.ComboBox RegimeTributario 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2205
         TabIndex        =   129
         Top             =   4170
         Width           =   4395
      End
      Begin VB.ComboBox Mensagem 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2220
         TabIndex        =   121
         Top             =   3150
         Width           =   4395
      End
      Begin VB.TextBox Observacao 
         Height          =   315
         Left            =   2220
         MaxLength       =   150
         MultiLine       =   -1  'True
         TabIndex        =   111
         Top             =   3690
         Width           =   4185
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
         Height          =   795
         Left            =   135
         TabIndex        =   40
         Top             =   30
         Width           =   8595
         Begin VB.Label Label7 
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
            Left            =   180
            TabIndex        =   96
            Top             =   210
            Width           =   630
         End
         Begin VB.Label ClienteLabel 
            Height          =   210
            Index           =   3
            Left            =   960
            TabIndex        =   97
            Top             =   210
            Width           =   3255
         End
         Begin VB.Label FilialLabel3 
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
            Left            =   330
            TabIndex        =   98
            Top             =   510
            Width           =   480
         End
         Begin VB.Label FilialLabel 
            Height          =   210
            Index           =   2
            Left            =   960
            TabIndex        =   99
            Top             =   510
            Width           =   3045
         End
      End
      Begin MSMask.MaskEdBox InscricaoEstadual 
         Height          =   315
         Left            =   2220
         TabIndex        =   106
         Top             =   1440
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   18
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox InscricaoMunicipal 
         Height          =   315
         Left            =   2220
         TabIndex        =   107
         Top             =   2010
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
         Left            =   2220
         TabIndex        =   109
         Top             =   2580
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   12
         Mask            =   "##.####-##-#"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CGC 
         Height          =   315
         Left            =   2220
         TabIndex        =   104
         Top             =   900
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
         Left            =   5910
         TabIndex        =   112
         Top             =   900
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         PromptChar      =   " "
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
         Left            =   510
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   130
         Top             =   4200
         Width           =   1560
      End
      Begin VB.Label MensagemNFLabel 
         Caption         =   "Mensagem p/ N. Fiscal:"
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
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   122
         Top             =   3180
         Width           =   2040
      End
      Begin VB.Label Label1 
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
         Index           =   0
         Left            =   5475
         TabIndex        =   113
         Top             =   960
         Width           =   345
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "CGC/CPF:"
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
         Left            =   1245
         TabIndex        =   110
         Top             =   960
         Width           =   885
      End
      Begin VB.Label Label100 
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
         Left            =   525
         TabIndex        =   105
         Top             =   2640
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
         Height          =   195
         Left            =   405
         TabIndex        =   100
         Top             =   2070
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
         Left            =   465
         TabIndex        =   101
         Top             =   1500
         Width           =   1650
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
         Left            =   990
         TabIndex        =   102
         Top             =   3720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4500
      Index           =   4
      Left            =   255
      TabIndex        =   14
      Top             =   750
      Visible         =   0   'False
      Width           =   8820
      Begin VB.Frame Frame3 
         Caption         =   "Redespacho"
         Height          =   1155
         Left            =   4380
         TabIndex        =   117
         Top             =   2685
         Width           =   4260
         Begin VB.ComboBox TranspRedespacho 
            Height          =   315
            Left            =   1665
            TabIndex        =   119
            Top             =   285
            Width           =   2475
         End
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
            TabIndex        =   118
            Top             =   750
            Width           =   2100
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
            TabIndex        =   120
            Top             =   345
            Width           =   1365
         End
      End
      Begin VB.TextBox Guia 
         Height          =   300
         Left            =   6420
         MaxLength       =   10
         TabIndex        =   115
         Top             =   2295
         Width           =   1290
      End
      Begin VB.ComboBox TipoFrete 
         Height          =   315
         ItemData        =   "FiliaisClientesOcx.ctx":0000
         Left            =   1800
         List            =   "FiliaisClientesOcx.ctx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   2790
         Width           =   1125
      End
      Begin VB.ComboBox Regiao 
         Height          =   315
         Left            =   1815
         TabIndex        =   23
         Top             =   2295
         Width           =   2145
      End
      Begin VB.ComboBox Cobrador 
         Height          =   315
         Left            =   6420
         TabIndex        =   18
         Top             =   1395
         Width           =   1965
      End
      Begin VB.ComboBox Transportadora 
         Height          =   315
         Left            =   1800
         TabIndex        =   22
         Top             =   3270
         Width           =   1965
      End
      Begin MSMask.MaskEdBox ContaContabil 
         Height          =   315
         Left            =   6405
         TabIndex        =   16
         Top             =   930
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ComissaoVendas 
         Height          =   300
         Left            =   1800
         TabIndex        =   17
         Top             =   1395
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox FreqVisitas 
         Height          =   300
         Left            =   6420
         TabIndex        =   20
         Top             =   1860
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataUltVisita 
         Height          =   300
         Left            =   1815
         TabIndex        =   19
         Top             =   1845
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
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
         Height          =   795
         Left            =   135
         TabIndex        =   36
         Top             =   30
         Width           =   8595
         Begin VB.Label FilialLabel1 
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
            Left            =   330
            TabIndex        =   79
            Top             =   510
            Width           =   480
         End
         Begin VB.Label Label10 
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
            Index           =   0
            Left            =   180
            TabIndex        =   80
            Top             =   210
            Width           =   660
         End
         Begin VB.Label FilialLabel 
            Height          =   210
            Index           =   1
            Left            =   960
            TabIndex        =   81
            Top             =   510
            Width           =   3545
         End
         Begin VB.Label ClienteLabel 
            Height          =   210
            Index           =   2
            Left            =   960
            TabIndex        =   82
            Top             =   210
            Width           =   3255
         End
      End
      Begin MSMask.MaskEdBox Vendedor 
         Height          =   300
         Left            =   1800
         TabIndex        =   15
         Top             =   930
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
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
         Left            =   5880
         TabIndex        =   116
         Top             =   2325
         Width           =   555
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
         Left            =   525
         TabIndex        =   103
         Top             =   2820
         Width           =   1215
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
         Left            =   5010
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   83
         Top             =   990
         Width           =   1335
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
         Left            =   600
         TabIndex        =   84
         Top             =   1905
         Width           =   1125
      End
      Begin VB.Label Label48 
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
         Height          =   210
         Left            =   6990
         TabIndex        =   85
         Top             =   1890
         Width           =   345
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
         Left            =   4455
         TabIndex        =   86
         Top             =   1905
         Width           =   1905
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
         Left            =   1065
         TabIndex        =   87
         Top             =   2355
         Width           =   675
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
         Left            =   840
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   88
         Top             =   990
         Width           =   885
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
         Left            =   375
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   89
         Top             =   3330
         Width           =   1365
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
         Left            =   855
         TabIndex        =   90
         Top             =   1455
         Width           =   870
      End
      Begin VB.Label LabelCobrador 
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
         Left            =   5520
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   91
         Top             =   1455
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4635
      Index           =   3
      Left            =   255
      TabIndex        =   10
      Top             =   675
      Visible         =   0   'False
      Width           =   8820
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3375
         Index           =   2
         Left            =   105
         TabIndex        =   127
         Top             =   1365
         Visible         =   0   'False
         Width           =   8595
         Begin TelasFAT.TabEndereco TabEnd 
            Height          =   3435
            Index           =   2
            Left            =   105
            TabIndex        =   128
            Top             =   0
            Width           =   8355
            _ExtentX        =   14737
            _ExtentY        =   6059
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3375
         Index           =   0
         Left            =   105
         TabIndex        =   125
         Top             =   1365
         Width           =   8595
         Begin TelasFAT.TabEndereco TabEnd 
            Height          =   3435
            Index           =   0
            Left            =   105
            TabIndex        =   126
            Top             =   0
            Width           =   8355
            _ExtentX        =   14737
            _ExtentY        =   6059
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3375
         Index           =   1
         Left            =   105
         TabIndex        =   123
         Top             =   1365
         Visible         =   0   'False
         Width           =   8595
         Begin TelasFAT.TabEndereco TabEnd 
            Height          =   3435
            Index           =   1
            Left            =   105
            TabIndex        =   124
            Top             =   -15
            Width           =   8355
            _ExtentX        =   14737
            _ExtentY        =   6059
         End
      End
      Begin VB.Frame SSFrame5 
         Caption         =   "Endereços"
         Height          =   510
         Left            =   135
         TabIndex        =   37
         Top             =   780
         Width           =   8595
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
            Left            =   1110
            TabIndex        =   11
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
            Left            =   3720
            TabIndex        =   12
            Top             =   180
            Width           =   1200
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
            Left            =   6345
            TabIndex        =   13
            Top             =   180
            Width           =   1290
         End
      End
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
         Height          =   795
         Left            =   135
         TabIndex        =   38
         Top             =   -30
         Width           =   8595
         Begin VB.Label Label2 
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
            Left            =   180
            TabIndex        =   92
            Top             =   210
            Width           =   630
         End
         Begin VB.Label ClienteLabel 
            Height          =   210
            Index           =   1
            Left            =   960
            TabIndex        =   93
            Top             =   210
            Width           =   3255
         End
         Begin VB.Label FilialLabel2 
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
            Left            =   330
            TabIndex        =   94
            Top             =   510
            Width           =   480
         End
         Begin VB.Label FilialLabel 
            Height          =   210
            Index           =   0
            Left            =   960
            TabIndex        =   95
            Top             =   510
            Width           =   3545
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4170
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   945
      Width           =   8820
      Begin VB.CheckBox Ativo 
         Caption         =   "Ativo"
         Height          =   255
         Left            =   3480
         TabIndex        =   114
         Top             =   1140
         Width           =   855
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   2520
         Picture         =   "FiliaisClientesOcx.ctx":0018
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Numeração Automática"
         Top             =   1125
         Width           =   300
      End
      Begin VB.Frame Frame7 
         Caption         =   "Categorias"
         Height          =   1980
         Left            =   555
         TabIndex        =   32
         Top             =   2025
         Width           =   4890
         Begin VB.ComboBox ComboCategoriaClienteItem 
            Height          =   315
            Left            =   2445
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   495
            Width           =   1632
         End
         Begin VB.ComboBox ComboCategoriaCliente 
            Height          =   315
            Left            =   930
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   495
            Width           =   1548
         End
         Begin MSFlexGridLib.MSFlexGrid GridCategoria 
            Height          =   1545
            Left            =   555
            TabIndex        =   7
            Top             =   315
            Width           =   3780
            _ExtentX        =   6668
            _ExtentY        =   2725
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
      Begin MSComctlLib.TreeView Filiais 
         Height          =   3600
         Left            =   5760
         TabIndex        =   8
         Top             =   450
         Visible         =   0   'False
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   6350
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         PathSeparator   =   "/"
         Style           =   7
         Appearance      =   1
      End
      Begin MSMask.MaskEdBox CodFilial 
         Height          =   315
         Left            =   1980
         TabIndex        =   2
         Top             =   1110
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CodCliente 
         Height          =   315
         Left            =   1980
         TabIndex        =   1
         Top             =   600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Nome 
         Height          =   315
         Left            =   1980
         TabIndex        =   4
         Top             =   1575
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   40
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
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
         Height          =   210
         Left            =   675
         TabIndex        =   41
         Top             =   1620
         Width           =   1275
      End
      Begin VB.Label LabelCodFilial 
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
         Height          =   255
         Left            =   810
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   42
         Top             =   1170
         Width           =   1125
      End
      Begin VB.Label Label13 
         Caption         =   "Clientes - Filiais"
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
         Left            =   5745
         TabIndex        =   43
         Top             =   180
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label LabelCodCliente 
         Caption         =   "Código Cliente:"
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
         Left            =   630
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   44
         Top             =   645
         Width           =   1320
      End
      Begin VB.Label Label6 
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
         Left            =   1245
         TabIndex        =   45
         Top             =   187
         Width           =   690
      End
      Begin VB.Label ClienteLabel 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   1980
         TabIndex        =   46
         Top             =   135
         Width           =   2385
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4170
      Index           =   5
      Left            =   240
      TabIndex        =   24
      Top             =   945
      Visible         =   0   'False
      Width           =   8820
      Begin VB.Frame SSFrame4 
         Caption         =   "Compras"
         Height          =   1560
         Left            =   4005
         TabIndex        =   31
         Top             =   1020
         Width           =   4620
         Begin VB.Label DataUltimaCompra 
            Caption         =   "  /  /    "
            Height          =   210
            Left            =   3255
            TabIndex        =   47
            Top             =   712
            Width           =   1170
         End
         Begin VB.Label DataPrimeiraCompra 
            Caption         =   "  /  /    "
            Height          =   210
            Left            =   3255
            TabIndex        =   48
            Top             =   322
            Width           =   1170
         End
         Begin VB.Label ValorAcumuladoCompras 
            Caption         =   "0,00"
            Height          =   210
            Left            =   1830
            TabIndex        =   49
            Top             =   1102
            Width           =   1575
         End
         Begin VB.Label MediaCompra 
            Caption         =   "0,00"
            Height          =   210
            Left            =   900
            TabIndex        =   50
            Top             =   712
            Width           =   1410
         End
         Begin VB.Label NumeroCompras 
            Caption         =   "0"
            Height          =   210
            Left            =   1050
            TabIndex        =   51
            Top             =   322
            Width           =   585
         End
         Begin VB.Label Label31 
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
            TabIndex        =   52
            Top             =   330
            Width           =   720
         End
         Begin VB.Label Label37 
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
            Left            =   225
            TabIndex        =   53
            Top             =   720
            Width           =   585
         End
         Begin VB.Label Label38 
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
            TabIndex        =   54
            Top             =   712
            Width           =   615
         End
         Begin VB.Label Label40 
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
            Left            =   2430
            TabIndex        =   55
            Top             =   322
            Width           =   765
         End
         Begin VB.Label Label41 
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
            TabIndex        =   56
            Top             =   1110
            Width           =   1500
         End
      End
      Begin VB.Frame SSFrame6 
         Caption         =   "Cheques Devolvidos"
         Height          =   1230
         Left            =   6225
         TabIndex        =   39
         Top             =   2760
         Width           =   2400
         Begin VB.Label DataUltChequeDevolvido 
            Caption         =   "  /  /    "
            Height          =   210
            Left            =   990
            TabIndex        =   57
            Top             =   855
            Width           =   1170
         End
         Begin VB.Label NumChequesDevolvidos 
            Caption         =   "0"
            Height          =   210
            Left            =   990
            TabIndex        =   58
            Top             =   420
            Width           =   405
         End
         Begin VB.Label Label42 
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
            Left            =   180
            TabIndex        =   59
            Top             =   420
            Width           =   750
         End
         Begin VB.Label Label43 
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
            Left            =   330
            TabIndex        =   60
            Top             =   855
            Width           =   600
         End
      End
      Begin VB.Frame SSFrame7 
         Caption         =   "Atraso"
         Height          =   1230
         Left            =   180
         TabIndex        =   33
         Top             =   2760
         Width           =   5925
         Begin VB.Label MaiorAtraso 
            Caption         =   "0"
            Height          =   210
            Left            =   765
            TabIndex        =   61
            Top             =   810
            Width           =   750
         End
         Begin VB.Label MediaAtraso 
            Caption         =   "0"
            Height          =   210
            Left            =   765
            TabIndex        =   62
            Top             =   360
            Width           =   750
         End
         Begin VB.Label ValorPagtosAtraso 
            Caption         =   "0,00"
            Height          =   210
            Left            =   4440
            TabIndex        =   63
            Top             =   810
            Width           =   1395
         End
         Begin VB.Label SaldoAtrasados 
            Caption         =   "0,00"
            Height          =   210
            Left            =   4440
            TabIndex        =   64
            Top             =   360
            Width           =   1395
         End
         Begin VB.Label Label45 
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
            Left            =   165
            TabIndex        =   65
            Top             =   810
            Width           =   570
         End
         Begin VB.Label Label46 
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
            Left            =   135
            TabIndex        =   66
            Top             =   360
            Width           =   585
         End
         Begin VB.Label Label56 
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
            Left            =   1770
            TabIndex        =   67
            Top             =   810
            Width           =   2610
         End
         Begin VB.Label Label57 
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
            Left            =   2655
            TabIndex        =   68
            Top             =   360
            Width           =   1725
         End
      End
      Begin VB.Frame SSFrame8 
         Caption         =   "Saldos"
         Height          =   1560
         Left            =   180
         TabIndex        =   34
         Top             =   1020
         Width           =   3675
         Begin VB.Label SaldoPedidosLiberados 
            Caption         =   "0,00"
            Height          =   210
            Left            =   1800
            TabIndex        =   69
            Top             =   1140
            Width           =   1575
         End
         Begin VB.Label SaldoDuplicatas 
            Caption         =   "0,00"
            Height          =   210
            Left            =   1800
            TabIndex        =   70
            Top             =   735
            Width           =   1575
         End
         Begin VB.Label SaldoTitulos 
            Caption         =   "0,00"
            Height          =   210
            Left            =   1800
            TabIndex        =   71
            Top             =   345
            Width           =   1575
         End
         Begin VB.Label Label63 
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
            Left            =   105
            TabIndex        =   72
            Top             =   1140
            Width           =   1650
         End
         Begin VB.Label Label64 
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
            Left            =   750
            TabIndex        =   73
            Top             =   345
            Width           =   1020
         End
         Begin VB.Label Label65 
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
            Left            =   435
            TabIndex        =   74
            Top             =   735
            Width           =   1320
         End
      End
      Begin VB.Frame SSFrame9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   135
         TabIndex        =   35
         Top             =   30
         Width           =   8595
         Begin VB.Label FilialLabel 
            Height          =   210
            Index           =   3
            Left            =   960
            TabIndex        =   75
            Top             =   510
            Width           =   3045
         End
         Begin VB.Label Label66 
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
            Left            =   330
            TabIndex        =   76
            Top             =   510
            Width           =   480
         End
         Begin VB.Label ClienteLabel 
            Height          =   210
            Index           =   4
            Left            =   960
            TabIndex        =   77
            Top             =   210
            Width           =   3255
         End
         Begin VB.Label Label67 
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
            Left            =   180
            TabIndex        =   78
            Top             =   210
            Width           =   630
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6975
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   45
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "FiliaisClientesOcx.ctx":0102
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "FiliaisClientesOcx.ctx":025C
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "FiliaisClientesOcx.ctx":03E6
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "FiliaisClientesOcx.ctx":0918
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5010
      Left            =   150
      TabIndex        =   30
      Top             =   330
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   8837
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
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
Attribute VB_Name = "FiliaisClientesOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTFiliaisClientes
Attribute objCT.VB_VarHelpID = -1

Private Sub Ativo_Click()
     Call objCT.Ativo_Click
End Sub

Private Sub BotaoProxNum_Click()
     Call objCT.BotaoProxNum_Click
End Sub

Private Sub BotaoExcluir_Click()
     Call objCT.BotaoExcluir_Click
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
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

Private Sub CGC_GotFocus()
     Call objCT.CGC_GotFocus
End Sub

Private Sub CGC_Validate(Cancel As Boolean)
     Call objCT.CGC_Validate(Cancel)
End Sub

Private Sub RG_Change()
     Call objCT.RG_Change
End Sub

Private Sub RG_GotFocus()
     Call objCT.RG_GotFocus
End Sub

Private Sub Cobrador_Change()
     Call objCT.Cobrador_Change
End Sub

Private Sub Cobrador_Click()
     Call objCT.Cobrador_Click
End Sub

Private Sub Cobrador_Validate(Cancel As Boolean)
     Call objCT.Cobrador_Validate(Cancel)
End Sub

Private Sub CodCliente_Change()
     Call objCT.CodCliente_Change
End Sub

Private Sub CodCliente_GotFocus()
     Call objCT.CodCliente_GotFocus
End Sub

Private Sub CodCliente_Validate(Cancel As Boolean)
     Call objCT.CodCliente_Validate(Cancel)
End Sub

Private Sub LabelCodCliente_Click()
     Call objCT.LabelCodCliente_Click
End Sub

Private Sub CodFilial_Change()
     Call objCT.CodFilial_Change
End Sub

Private Sub CodFilial_GotFocus()
     Call objCT.CodFilial_GotFocus
End Sub

Private Sub CodFilial_Validate(Cancel As Boolean)
     Call objCT.CodFilial_Validate(Cancel)
End Sub

Private Sub LabelCodFilial_Click()
     Call objCT.LabelCodFilial_Click
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

Private Sub ContaContabil_Change()
     Call objCT.ContaContabil_Change
End Sub

Private Sub ContaContabil_Validate(Cancel As Boolean)
     Call objCT.ContaContabil_Validate(Cancel)
End Sub

Private Sub ComissaoVendas_Change()
     Call objCT.ComissaoVendas_Change
End Sub

Private Sub ComissaoVendas_Validate(Cancel As Boolean)
     Call objCT.ComissaoVendas_Validate(Cancel)
End Sub

Private Sub DataUltVisita_Change()
     Call objCT.DataUltVisita_Change
End Sub

Private Sub DataUltVisita_GotFocus()
     Call objCT.DataUltVisita_GotFocus
End Sub

Private Sub DataUltVisita_Validate(Cancel As Boolean)
     Call objCT.DataUltVisita_Validate(Cancel)
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
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

Private Sub TransportadoraLabel_Click()
     Call objCT.TransportadoraLabel_Click
End Sub

Private Sub LabelCobrador_Click()
     Call objCT.LabelCobrador_Click
End Sub

Private Sub Observacao_Change()
     Call objCT.Observacao_Change
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

Private Sub Transportadora_Validate(Cancel As Boolean)
     Call objCT.Transportadora_Validate(Cancel)
End Sub

Private Sub UserControl_Initialize()
    Set objCT = New CTFiliaisClientes
    Set objCT.objUserControl = Me
End Sub

Private Sub Vendedor_Change()
     Call objCT.Vendedor_Change
End Sub

Private Sub Vendedor_Validate(Cancel As Boolean)
     Call objCT.Vendedor_Validate(Cancel)
End Sub

Private Sub Transportadora_Change()
     Call objCT.Transportadora_Change
End Sub

Private Sub Transportadora_Click()
     Call objCT.Transportadora_Click
End Sub

Private Sub OpcaoEndereco_Click(Index As Integer)
     Call objCT.OpcaoEndereco_Click(Index)
End Sub

Private Sub Filiais_NodeClick(ByVal Node As MSComctlLib.Node)
     Call objCT.Filiais_NodeClick(Node)
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Function Trata_Parametros(Optional objFilialCliente As ClassFilialCliente) As Long
     Trata_Parametros = objCT.Trata_Parametros(objFilialCliente)
End Function

Private Sub InscricaoEstadual_Change()
     Call objCT.InscricaoEstadual_Change
End Sub

Private Sub Inscricaosuframa_Change()
     Call objCT.Inscricaosuframa_Change
End Sub

Private Sub InscricaoMunicipal_Change()
     Call objCT.InscricaoMunicipal_Change
End Sub

Private Sub Nome_Change()
     Call objCT.Nome_Change
End Sub

Private Sub Nome_Validate(Cancel As Boolean)
     Call objCT.Nome_Validate(Cancel)
End Sub

Private Sub Opcao_Click()
     Call objCT.Opcao_Click
End Sub

Private Sub VendedorLabel_Click()
     Call objCT.VendedorLabel_Click
End Sub

Private Sub ContaContabilLabel_Click()
     Call objCT.ContaContabilLabel_Click
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

Private Sub LabelCodFilial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodFilial, Source, X, Y)
End Sub

Private Sub LabelCodFilial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodFilial, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub LabelCodCliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodCliente, Source, X, Y)
End Sub

Private Sub LabelCodCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodCliente, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
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

Private Sub Label31_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label31, Source, X, Y)
End Sub

Private Sub Label31_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label31, Button, Shift, X, Y)
End Sub

Private Sub Label37_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label37, Source, X, Y)
End Sub

Private Sub Label37_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label37, Button, Shift, X, Y)
End Sub

Private Sub Label38_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label38, Source, X, Y)
End Sub

Private Sub Label38_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label38, Button, Shift, X, Y)
End Sub

Private Sub Label40_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label40, Source, X, Y)
End Sub

Private Sub Label40_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label40, Button, Shift, X, Y)
End Sub

Private Sub Label41_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label41, Source, X, Y)
End Sub

Private Sub Label41_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label41, Button, Shift, X, Y)
End Sub

Private Sub DataUltChequeDevolvido_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataUltChequeDevolvido, Source, X, Y)
End Sub

Private Sub DataUltChequeDevolvido_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataUltChequeDevolvido, Button, Shift, X, Y)
End Sub

Private Sub NumChequesDevolvidos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumChequesDevolvidos, Source, X, Y)
End Sub

Private Sub NumChequesDevolvidos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumChequesDevolvidos, Button, Shift, X, Y)
End Sub

Private Sub Label42_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label42, Source, X, Y)
End Sub

Private Sub Label42_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label42, Button, Shift, X, Y)
End Sub

Private Sub Label43_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label43, Source, X, Y)
End Sub

Private Sub Label43_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label43, Button, Shift, X, Y)
End Sub

Private Sub MaiorAtraso_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(MaiorAtraso, Source, X, Y)
End Sub

Private Sub MaiorAtraso_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(MaiorAtraso, Button, Shift, X, Y)
End Sub

Private Sub MediaAtraso_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(MediaAtraso, Source, X, Y)
End Sub

Private Sub MediaAtraso_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(MediaAtraso, Button, Shift, X, Y)
End Sub

Private Sub ValorPagtosAtraso_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorPagtosAtraso, Source, X, Y)
End Sub

Private Sub ValorPagtosAtraso_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorPagtosAtraso, Button, Shift, X, Y)
End Sub

Private Sub SaldoAtrasados_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SaldoAtrasados, Source, X, Y)
End Sub

Private Sub SaldoAtrasados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SaldoAtrasados, Button, Shift, X, Y)
End Sub

Private Sub Label45_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label45, Source, X, Y)
End Sub

Private Sub Label45_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label45, Button, Shift, X, Y)
End Sub

Private Sub Label46_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label46, Source, X, Y)
End Sub

Private Sub Label46_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label46, Button, Shift, X, Y)
End Sub

Private Sub Label56_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label56, Source, X, Y)
End Sub

Private Sub Label56_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label56, Button, Shift, X, Y)
End Sub

Private Sub Label57_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label57, Source, X, Y)
End Sub

Private Sub Label57_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label57, Button, Shift, X, Y)
End Sub

Private Sub SaldoPedidosLiberados_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SaldoPedidosLiberados, Source, X, Y)
End Sub

Private Sub SaldoPedidosLiberados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SaldoPedidosLiberados, Button, Shift, X, Y)
End Sub

Private Sub SaldoDuplicatas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SaldoDuplicatas, Source, X, Y)
End Sub

Private Sub SaldoDuplicatas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SaldoDuplicatas, Button, Shift, X, Y)
End Sub

Private Sub SaldoTitulos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SaldoTitulos, Source, X, Y)
End Sub

Private Sub SaldoTitulos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SaldoTitulos, Button, Shift, X, Y)
End Sub

Private Sub Label63_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label63, Source, X, Y)
End Sub

Private Sub Label63_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label63, Button, Shift, X, Y)
End Sub

Private Sub Label64_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label64, Source, X, Y)
End Sub

Private Sub Label64_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label64, Button, Shift, X, Y)
End Sub

Private Sub Label65_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label65, Source, X, Y)
End Sub

Private Sub Label65_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label65, Button, Shift, X, Y)
End Sub

Private Sub Label66_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label66, Source, X, Y)
End Sub

Private Sub Label66_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label66, Button, Shift, X, Y)
End Sub

Private Sub Label67_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label67, Source, X, Y)
End Sub

Private Sub Label67_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label67, Button, Shift, X, Y)
End Sub

Private Sub FilialLabel1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialLabel1, Source, X, Y)
End Sub

Private Sub FilialLabel1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialLabel1, Button, Shift, X, Y)
End Sub

Private Sub ContaContabilLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContaContabilLabel, Source, X, Y)
End Sub

Private Sub ContaContabilLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContaContabilLabel, Button, Shift, X, Y)
End Sub

Private Sub Label49_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label49, Source, X, Y)
End Sub

Private Sub Label49_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label49, Button, Shift, X, Y)
End Sub

Private Sub Label48_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label48, Source, X, Y)
End Sub

Private Sub Label48_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label48, Button, Shift, X, Y)
End Sub

Private Sub Label47_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label47, Source, X, Y)
End Sub

Private Sub Label47_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label47, Button, Shift, X, Y)
End Sub

Private Sub Label33_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label33, Source, X, Y)
End Sub

Private Sub Label33_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label33, Button, Shift, X, Y)
End Sub

Private Sub VendedorLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(VendedorLabel, Source, X, Y)
End Sub

Private Sub VendedorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(VendedorLabel, Button, Shift, X, Y)
End Sub

Private Sub TransportadoraLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TransportadoraLabel, Source, X, Y)
End Sub

Private Sub TransportadoraLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TransportadoraLabel, Button, Shift, X, Y)
End Sub

Private Sub Label44_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label44, Source, X, Y)
End Sub

Private Sub Label44_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label44, Button, Shift, X, Y)
End Sub

Private Sub LabelCobrador_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCobrador, Source, X, Y)
End Sub

Private Sub LabelCobrador_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCobrador, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub FilialLabel2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialLabel2, Source, X, Y)
End Sub

Private Sub FilialLabel2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialLabel2, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub FilialLabel3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialLabel3, Source, X, Y)
End Sub

Private Sub FilialLabel3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialLabel3, Button, Shift, X, Y)
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

'#########################################################
'Inserido por Wagner 26/09/05
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

Private Sub MensagemNFLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(MensagemNFLabel, Source, X, Y)
End Sub

Private Sub MensagemNFLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(MensagemNFLabel, Button, Shift, X, Y)
End Sub
'#########################################################

Private Sub IEIsento_Click()
    Call objCT.IEIsento_Click
End Sub

Private Sub IENaoContrib_Click()
    Call objCT.IENaoContrib_Click
End Sub

Private Sub InscricaoEstadual_Validate(Cancel As Boolean)
    Call objCT.InscricaoEstadual_Validate(Cancel)
End Sub
