VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ContatosOcx 
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9300
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5685
   ScaleWidth      =   9300
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4515
      Index           =   4
      Left            =   165
      TabIndex        =   23
      Top             =   1080
      Visible         =   0   'False
      Width           =   8850
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3375
         Index           =   2
         Left            =   75
         TabIndex        =   99
         Top             =   1080
         Visible         =   0   'False
         Width           =   8595
         Begin TelasFAT.TabEndereco TabEnd 
            Height          =   3435
            Index           =   2
            Left            =   195
            TabIndex        =   100
            Top             =   105
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
         Left            =   75
         TabIndex        =   97
         Top             =   1080
         Width           =   8595
         Begin TelasFAT.TabEndereco TabEnd 
            Height          =   3435
            Index           =   0
            Left            =   195
            TabIndex        =   98
            Top             =   105
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
         Left            =   75
         TabIndex        =   50
         Top             =   1080
         Visible         =   0   'False
         Width           =   8595
         Begin TelasFAT.TabEndereco TabEnd 
            Height          =   3435
            Index           =   1
            Left            =   195
            TabIndex        =   96
            Top             =   105
            Width           =   8355
            _ExtentX        =   14737
            _ExtentY        =   6059
         End
      End
      Begin VB.Frame SSFrame5 
         Caption         =   "Endereços"
         Height          =   570
         Left            =   240
         TabIndex        =   51
         Top             =   510
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
            TabIndex        =   24
            Top             =   210
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
            TabIndex        =   25
            Top             =   210
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
            TabIndex        =   26
            Top             =   210
            Width           =   1350
         End
      End
      Begin VB.Frame SSFrame7 
         Height          =   555
         Index           =   2
         Left            =   240
         TabIndex        =   52
         Top             =   -45
         Width           =   8445
         Begin VB.Label Label30 
            Caption         =   "Cliente Futuro:"
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
            TabIndex        =   72
            Top             =   210
            Width           =   1380
         End
         Begin VB.Label ContatoLabel 
            Height          =   210
            Index           =   2
            Left            =   1980
            TabIndex        =   73
            Top             =   210
            Width           =   6150
         End
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
      Height          =   520
      Left            =   3555
      Picture         =   "ContatosOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4140
      Index           =   5
      Left            =   150
      TabIndex        =   27
      Top             =   1260
      Visible         =   0   'False
      Width           =   8850
      Begin VB.TextBox Guia 
         Height          =   300
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   38
         Top             =   3015
         Width           =   1290
      End
      Begin VB.Frame Frame3 
         Caption         =   "Redespacho"
         Height          =   1155
         Left            =   4470
         TabIndex        =   91
         Top             =   2925
         Width           =   4260
         Begin VB.CheckBox RedespachoCli 
            Caption         =   "por conta do Contato"
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
            TabIndex        =   40
            Top             =   750
            Width           =   2475
         End
         Begin VB.ComboBox TranspRedespacho 
            Height          =   315
            Left            =   1665
            TabIndex        =   39
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
            TabIndex        =   92
            Top             =   345
            Width           =   1365
         End
      End
      Begin VB.ComboBox TipoFrete 
         Height          =   315
         ItemData        =   "ContatosOcx.ctx":20B6
         Left            =   6255
         List            =   "ContatosOcx.ctx":20C0
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   2055
         Width           =   1125
      End
      Begin VB.ComboBox PadraoCobranca 
         Height          =   315
         ItemData        =   "ContatosOcx.ctx":20CE
         Left            =   2040
         List            =   "ContatosOcx.ctx":20D0
         TabIndex        =   32
         Top             =   1590
         Width           =   1965
      End
      Begin VB.ComboBox Transportadora 
         Height          =   315
         Left            =   6255
         TabIndex        =   37
         Top             =   2520
         Width           =   2475
      End
      Begin VB.ComboBox Cobrador 
         Height          =   315
         Left            =   6255
         TabIndex        =   31
         Top             =   1125
         Width           =   2475
      End
      Begin VB.ComboBox Regiao 
         Height          =   315
         Left            =   2040
         TabIndex        =   36
         Top             =   2520
         Width           =   2385
      End
      Begin MSMask.MaskEdBox ContaContabil 
         Height          =   315
         Left            =   6255
         TabIndex        =   29
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
         TabIndex        =   30
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
         TabIndex        =   33
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
         TabIndex        =   34
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
         TabIndex        =   28
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
         TabIndex        =   49
         Top             =   30
         Width           =   8445
         Begin VB.Label Label30 
            Caption         =   "Cliente Futuro:"
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
            TabIndex        =   60
            Top             =   210
            Width           =   1395
         End
         Begin VB.Label ContatoLabel 
            Height          =   210
            Index           =   0
            Left            =   1635
            TabIndex        =   61
            Top             =   210
            Width           =   6540
         End
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
         TabIndex        =   93
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
         TabIndex        =   87
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
         TabIndex        =   62
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
         TabIndex        =   63
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
         TabIndex        =   64
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
         TabIndex        =   65
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
         TabIndex        =   66
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
         TabIndex        =   67
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
         TabIndex        =   68
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
         TabIndex        =   69
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
         TabIndex        =   70
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
         TabIndex        =   71
         Top             =   1200
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4140
      Index           =   3
      Left            =   165
      TabIndex        =   16
      Top             =   1320
      Visible         =   0   'False
      Width           =   8850
      Begin VB.TextBox Observacao2 
         Height          =   315
         Left            =   2265
         MaxLength       =   150
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   3210
         Width           =   4185
      End
      Begin MSMask.MaskEdBox CGC 
         Height          =   315
         Left            =   2265
         TabIndex        =   17
         Top             =   990
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   14
         Mask            =   "##############"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox InscricaoEstadual 
         Height          =   315
         Left            =   2265
         TabIndex        =   19
         Top             =   1545
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
         Left            =   2265
         TabIndex        =   20
         Top             =   2100
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   18
         PromptChar      =   " "
      End
      Begin VB.Frame SSFrame7 
         Height          =   555
         Index           =   3
         Left            =   240
         TabIndex        =   53
         Top             =   30
         Width           =   8445
         Begin VB.Label Label30 
            Caption         =   "Cliente Futuro:"
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
            TabIndex        =   74
            Top             =   210
            Width           =   1395
         End
         Begin VB.Label ContatoLabel 
            Height          =   210
            Index           =   3
            Left            =   1620
            TabIndex        =   75
            Top             =   210
            Width           =   6345
         End
      End
      Begin MSMask.MaskEdBox InscricaoSuframa 
         Height          =   315
         Left            =   2265
         TabIndex        =   21
         Top             =   2655
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
         Left            =   5415
         TabIndex        =   18
         Top             =   990
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         PromptChar      =   " "
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
         Left            =   4980
         TabIndex        =   89
         Top             =   1050
         Width           =   345
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
         Left            =   540
         TabIndex        =   88
         Top             =   2715
         Width           =   1605
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
         Left            =   1050
         TabIndex        =   76
         Top             =   3270
         Width           =   1095
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
         Height          =   195
         Left            =   1170
         TabIndex        =   77
         Top             =   1050
         Width           =   990
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
         Left            =   495
         TabIndex        =   78
         Top             =   1605
         Width           =   1650
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
         Left            =   420
         TabIndex        =   79
         Top             =   2160
         Width           =   1725
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4140
      Index           =   1
      Left            =   165
      TabIndex        =   0
      Top             =   1275
      Width           =   8850
      Begin VB.TextBox RazaoSocial 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   570
         Width           =   3720
      End
      Begin VB.CheckBox Ativo 
         Caption         =   "Ativo"
         Height          =   255
         Left            =   3480
         TabIndex        =   90
         Top             =   195
         Width           =   1215
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   2535
         Picture         =   "ContatosOcx.ctx":20D2
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Numeração Automática"
         Top             =   180
         Width           =   300
      End
      Begin VB.Frame Frame7 
         Caption         =   "Categorias"
         Height          =   1896
         Left            =   225
         TabIndex        =   48
         Top             =   2160
         Width           =   5145
         Begin VB.ComboBox ComboCategoriaContato 
            Height          =   315
            Left            =   975
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   405
            Width           =   1545
         End
         Begin VB.ComboBox ComboCategoriaContatoItem 
            Height          =   315
            Left            =   2490
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   405
            Width           =   1635
         End
         Begin MSFlexGridLib.MSFlexGrid GridCategoria 
            Height          =   1560
            Left            =   630
            TabIndex        =   9
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
      Begin VB.ComboBox Tipo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   1365
         Width           =   2790
      End
      Begin VB.TextBox Observacao 
         Height          =   315
         Left            =   1680
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1755
         Width           =   3675
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   1695
         TabIndex        =   1
         Top             =   165
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
         Left            =   1680
         TabIndex        =   4
         Top             =   960
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
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
         Left            =   960
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   55
         Top             =   240
         Width           =   660
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
         Left            =   1050
         TabIndex        =   56
         Top             =   630
         Width           =   555
      End
      Begin VB.Label Label3 
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
         Left            =   180
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   57
         Top             =   1005
         Width           =   1410
      End
      Begin VB.Label TipoContatoLabel 
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
         Left            =   1140
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   58
         Top             =   1425
         Width           =   450
      End
      Begin VB.Label Label5 
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
         Left            =   510
         TabIndex        =   59
         Top             =   1815
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4140
      Index           =   2
      Left            =   165
      TabIndex        =   10
      Top             =   1260
      Visible         =   0   'False
      Width           =   8850
      Begin VB.CheckBox Bloqueado 
         Caption         =   "Contato com crédito bloqueado"
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
         Left            =   4680
         TabIndex        =   94
         Top             =   1530
         Width           =   3585
      End
      Begin VB.ComboBox TabelaPreco 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2880
         TabIndex        =   13
         Top             =   2070
         Width           =   2220
      End
      Begin VB.ComboBox Mensagem 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2895
         TabIndex        =   15
         Top             =   3210
         Width           =   4395
      End
      Begin VB.ComboBox CondicaoPagto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2910
         TabIndex        =   14
         Top             =   2610
         Width           =   2670
      End
      Begin MSMask.MaskEdBox LimiteCredito 
         Height          =   315
         Left            =   2880
         TabIndex        =   12
         Top             =   1500
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
         Left            =   2880
         TabIndex        =   11
         Top             =   953
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin VB.Frame SSFrame7 
         Height          =   555
         Index           =   4
         Left            =   240
         TabIndex        =   54
         Top             =   30
         Width           =   8445
         Begin VB.Label Label30 
            Caption         =   "Contato:"
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
            TabIndex        =   80
            Top             =   210
            Width           =   780
         End
         Begin VB.Label ContatoLabel 
            Height          =   210
            Index           =   1
            Left            =   1080
            TabIndex        =   81
            Top             =   210
            Width           =   7080
         End
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
         Left            =   1230
         TabIndex        =   82
         Top             =   2100
         Width           =   1590
      End
      Begin VB.Label MensagemNFLabel 
         Caption         =   "Mensagem para Nota Fiscal:"
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
         TabIndex        =   83
         Top             =   3210
         Width           =   2460
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
         Left            =   1935
         TabIndex        =   84
         Top             =   1005
         Width           =   885
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
         Left            =   660
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   85
         Top             =   2655
         Width           =   2160
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
         Left            =   1290
         TabIndex        =   86
         Top             =   1545
         Width           =   1530
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
      Height          =   520
      Left            =   5220
      Picture         =   "ContatosOcx.ctx":21BC
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6975
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "ContatosOcx.ctx":2F5E
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "ContatosOcx.ctx":30B8
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1110
         Picture         =   "ContatosOcx.ctx":3242
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1620
         Picture         =   "ContatosOcx.ctx":3774
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   4950
      Left            =   120
      TabIndex        =   47
      Top             =   705
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   8731
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Financeiros"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Inscrições"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Endereços"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Vendas"
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
Attribute VB_Name = "ContatosOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Pendencias: diminuir tamanho do Form_Load

Option Explicit

Event Unload()

Private WithEvents objCT As CTContatos
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

Private Sub ComboCategoriaContato_Change()
     Call objCT.ComboCategoriaContato_Change
End Sub

Private Sub ComboCategoriaContato_Click()
     Call objCT.ComboCategoriaContato_Click
End Sub

Private Sub ComboCategoriaContato_GotFocus()
     Call objCT.ComboCategoriaContato_GotFocus
End Sub

Private Sub ComboCategoriaContato_KeyPress(KeyAscii As Integer)
     Call objCT.ComboCategoriaContato_KeyPress(KeyAscii)
End Sub

Private Sub ComboCategoriaContato_Validate(Cancel As Boolean)
     Call objCT.ComboCategoriaContato_Validate(Cancel)
End Sub

Private Sub ComboCategoriaContatoItem_Change()
     Call objCT.ComboCategoriaContatoItem_Change
End Sub

Private Sub ComboCategoriaContatoItem_Click()
     Call objCT.ComboCategoriaContatoItem_Click
End Sub

Private Sub ComboCategoriaContatoItem_GotFocus()
     Call objCT.ComboCategoriaContatoItem_GotFocus
End Sub

Private Sub ComboCategoriaContatoItem_KeyPress(KeyAscii As Integer)
     Call objCT.ComboCategoriaContatoItem_KeyPress(KeyAscii)
End Sub

Private Sub ComboCategoriaContatoItem_Validate(Cancel As Boolean)
     Call objCT.ComboCategoriaContatoItem_Validate(Cancel)
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

Private Sub TipoContatoLabel_Click()
     Call objCT.TipoContatoLabel_Click
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
    Set objCT = New CTContatos
    Set objCT.objUserControl = Me
End Sub

Private Sub Vendedor_Change()
     Call objCT.Vendedor_Change
End Sub

Private Sub Vendedor_Validate(Cancel As Boolean)
     Call objCT.Vendedor_Validate(Cancel)
End Sub

Function Trata_Parametros(Optional objContato As ClassContatos) As Long
     Trata_Parametros = objCT.Trata_Parametros(objContato)
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

Private Sub ContatoLabel_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(ContatoLabel(Index), Source, X, Y)
End Sub

Private Sub ContatoLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContatoLabel(Index), Button, Shift, X, Y)
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

Private Sub TipoContatoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoContatoLabel, Source, X, Y)
End Sub

Private Sub TipoContatoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoContatoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
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

Private Sub Bloqueado_Click()
     Call objCT.Bloqueado_Click
End Sub

Private Sub BotaoContatos_Click()
     Call objCT.BotaoContatos_Click
End Sub
