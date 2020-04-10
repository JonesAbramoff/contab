VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl FiliaisContatosOcx 
   ClientHeight    =   5370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9285
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5370
   ScaleWidth      =   9285
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4635
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   630
      Visible         =   0   'False
      Width           =   8820
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3375
         Index           =   1
         Left            =   105
         TabIndex        =   88
         Top             =   1365
         Visible         =   0   'False
         Width           =   8595
         Begin TelasFAT.TabEndereco TabEnd 
            Height          =   3435
            Index           =   1
            Left            =   105
            TabIndex        =   89
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
         TabIndex        =   86
         Top             =   1365
         Width           =   8595
         Begin TelasFAT.TabEndereco TabEnd 
            Height          =   3435
            Index           =   0
            Left            =   105
            TabIndex        =   87
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
         Index           =   2
         Left            =   105
         TabIndex        =   84
         Top             =   1365
         Visible         =   0   'False
         Width           =   8595
         Begin TelasFAT.TabEndereco TabEnd 
            Height          =   3435
            Index           =   2
            Left            =   105
            TabIndex        =   85
            Top             =   0
            Width           =   8355
            _ExtentX        =   14737
            _ExtentY        =   6059
         End
      End
      Begin VB.Frame SSFrame5 
         Caption         =   "Endereços"
         Height          =   540
         Left            =   135
         TabIndex        =   32
         Top             =   840
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
            Left            =   3720
            TabIndex        =   12
            Top             =   195
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
            Top             =   195
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
         TabIndex        =   33
         Top             =   30
         Width           =   8595
         Begin VB.Label Label2 
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
            Left            =   180
            TabIndex        =   54
            Top             =   210
            Width           =   1350
         End
         Begin VB.Label ContatoLabel 
            Height          =   210
            Index           =   1
            Left            =   1695
            TabIndex        =   55
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
            Left            =   975
            TabIndex        =   56
            Top             =   510
            Width           =   480
         End
         Begin VB.Label FilialLabel 
            Height          =   210
            Index           =   0
            Left            =   1695
            TabIndex        =   57
            Top             =   510
            Width           =   3540
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4170
      Index           =   4
      Left            =   255
      TabIndex        =   14
      Top             =   945
      Visible         =   0   'False
      Width           =   8820
      Begin VB.Frame Frame3 
         Caption         =   "Redespacho"
         Height          =   1155
         Left            =   4380
         TabIndex        =   78
         Top             =   2685
         Width           =   4260
         Begin VB.ComboBox TranspRedespacho 
            Height          =   315
            Left            =   1665
            TabIndex        =   80
            Top             =   285
            Width           =   2475
         End
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
            TabIndex        =   79
            Top             =   750
            Width           =   2385
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
            TabIndex        =   81
            Top             =   345
            Width           =   1365
         End
      End
      Begin VB.TextBox Guia 
         Height          =   300
         Left            =   6420
         MaxLength       =   10
         TabIndex        =   76
         Top             =   2295
         Width           =   1290
      End
      Begin VB.ComboBox TipoFrete 
         Height          =   315
         ItemData        =   "FiliaisContatoOcx.ctx":0000
         Left            =   1800
         List            =   "FiliaisContatoOcx.ctx":000A
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
         TabIndex        =   31
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
            Left            =   960
            TabIndex        =   41
            Top             =   510
            Width           =   480
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   42
            Top             =   210
            Width           =   1260
         End
         Begin VB.Label FilialLabel 
            Height          =   210
            Index           =   1
            Left            =   1635
            TabIndex        =   43
            Top             =   510
            Width           =   3540
         End
         Begin VB.Label ContatoLabel 
            Height          =   210
            Index           =   2
            Left            =   1635
            TabIndex        =   44
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
         TabIndex        =   77
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
         TabIndex        =   65
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
         TabIndex        =   45
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
         TabIndex        =   46
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
         TabIndex        =   47
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
         TabIndex        =   48
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
         TabIndex        =   49
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
         TabIndex        =   50
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
         TabIndex        =   51
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
         TabIndex        =   52
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
         TabIndex        =   53
         Top             =   1455
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4170
      Index           =   2
      Left            =   255
      TabIndex        =   9
      Top             =   915
      Visible         =   0   'False
      Width           =   8820
      Begin VB.ComboBox Mensagem 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2220
         TabIndex        =   82
         Top             =   3150
         Width           =   5745
      End
      Begin VB.TextBox Observacao 
         Height          =   315
         Left            =   2220
         MaxLength       =   150
         MultiLine       =   -1  'True
         TabIndex        =   72
         Top             =   3690
         Width           =   5715
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
         TabIndex        =   34
         Top             =   30
         Width           =   8595
         Begin VB.Label Label7 
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
            Left            =   180
            TabIndex        =   58
            Top             =   210
            Width           =   1365
         End
         Begin VB.Label ContatoLabel 
            Height          =   210
            Index           =   3
            Left            =   1620
            TabIndex        =   59
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
            Left            =   945
            TabIndex        =   60
            Top             =   510
            Width           =   480
         End
         Begin VB.Label FilialLabel 
            Height          =   210
            Index           =   2
            Left            =   1620
            TabIndex        =   61
            Top             =   510
            Width           =   3045
         End
      End
      Begin MSMask.MaskEdBox InscricaoEstadual 
         Height          =   315
         Left            =   2220
         TabIndex        =   68
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
         TabIndex        =   69
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
         TabIndex        =   70
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
         TabIndex        =   66
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
         TabIndex        =   73
         Top             =   900
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         PromptChar      =   " "
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
         TabIndex        =   83
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
         Left            =   5475
         TabIndex        =   74
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
         TabIndex        =   71
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
         TabIndex        =   67
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
         TabIndex        =   62
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
         TabIndex        =   63
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
         TabIndex        =   64
         Top             =   3720
         Width           =   1095
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
         TabIndex        =   75
         Top             =   1140
         Width           =   855
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   2520
         Picture         =   "FiliaisContatoOcx.ctx":0018
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
         TabIndex        =   30
         Top             =   2025
         Width           =   4890
         Begin VB.ComboBox ComboCategoriaContatoItem 
            Height          =   315
            Left            =   2445
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   495
            Width           =   1632
         End
         Begin VB.ComboBox ComboCategoriaContato 
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
      Begin MSMask.MaskEdBox CodContato 
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
         TabIndex        =   35
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
         TabIndex        =   36
         Top             =   1170
         Width           =   1125
      End
      Begin VB.Label Label13 
         Caption         =   "Clientes Futuros - Filiais"
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
         TabIndex        =   37
         Top             =   180
         Visible         =   0   'False
         Width           =   2190
      End
      Begin VB.Label LabelCodContato 
         Caption         =   "Código Cliente Futuro:"
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
         Left            =   0
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   38
         Top             =   660
         Width           =   2010
      End
      Begin VB.Label Label6 
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
         Left            =   645
         TabIndex        =   39
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label ContatoLabel 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   1980
         TabIndex        =   40
         Top             =   135
         Width           =   2385
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6930
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   45
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "FiliaisContatoOcx.ctx":0102
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "FiliaisContatoOcx.ctx":025C
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "FiliaisContatoOcx.ctx":03E6
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "FiliaisContatoOcx.ctx":0918
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5010
      Left            =   150
      TabIndex        =   29
      Top             =   315
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   8837
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
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
Attribute VB_Name = "FiliaisContatosOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTFiliaisContatos
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

Private Sub CodContato_Change()
     Call objCT.CodContato_Change
End Sub

Private Sub CodContato_GotFocus()
     Call objCT.CodContato_GotFocus
End Sub

Private Sub CodContato_Validate(Cancel As Boolean)
     Call objCT.CodContato_Validate(Cancel)
End Sub

Private Sub LabelCodContato_Click()
     Call objCT.LabelCodContato_Click
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
    Set objCT = New CTFiliaisContatos
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

Function Trata_Parametros(Optional objFilialContato As ClassFilialContato) As Long
     Trata_Parametros = objCT.Trata_Parametros(objFilialContato)
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

Private Sub ContatoLabel_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(ContatoLabel(Index), Source, X, Y)
End Sub

Private Sub ContatoLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContatoLabel(Index), Button, Shift, X, Y)
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

Private Sub LabelCodContato_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodContato, Source, X, Y)
End Sub

Private Sub LabelCodContato_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodContato, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
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
