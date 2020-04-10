VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl PedComprasGeradoOcx 
   ClientHeight    =   9195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   KeyPreview      =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   16995
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   8220
      Index           =   5
      Left            =   204
      TabIndex        =   90
      Top             =   744
      Visible         =   0   'False
      Width           =   16545
      Begin VB.CommandButton BotaoLiberaBloqueio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         Picture         =   "PedComprasGeradoOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   112
         Top             =   7590
         Width           =   1575
      End
      Begin VB.Frame SSFrame1 
         Caption         =   "Bloqueios"
         Height          =   7560
         Left            =   15
         TabIndex        =   104
         Top             =   -15
         Width           =   16365
         Begin VB.ComboBox TipoBloqueio 
            Height          =   315
            ItemData        =   "PedComprasGeradoOcx.ctx":25FA
            Left            =   180
            List            =   "PedComprasGeradoOcx.ctx":25FC
            TabIndex        =   106
            Top             =   570
            Width           =   2000
         End
         Begin MSMask.MaskEdBox ResponsavelLib 
            Height          =   270
            Left            =   6990
            TabIndex        =   111
            Top             =   480
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataLiberacao 
            Height          =   270
            Left            =   5880
            TabIndex        =   110
            Top             =   780
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodUsuario 
            Height          =   270
            Left            =   3450
            TabIndex        =   108
            Top             =   585
            Width           =   2000
            _ExtentX        =   3519
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ResponsavelBL 
            Height          =   270
            Left            =   4830
            TabIndex        =   109
            Top             =   570
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataBloqueio 
            Height          =   270
            Left            =   2205
            TabIndex        =   107
            Top             =   585
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridBloqueios 
            Height          =   2805
            Left            =   120
            TabIndex        =   105
            Top             =   240
            Width           =   16020
            _ExtentX        =   28258
            _ExtentY        =   4948
            _Version        =   393216
            Rows            =   7
            Cols            =   5
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   8220
      Index           =   6
      Left            =   150
      TabIndex        =   119
      Top             =   765
      Visible         =   0   'False
      Width           =   16650
      Begin VB.Frame Frame10 
         Caption         =   "Notas"
         Height          =   5730
         Left            =   105
         TabIndex        =   121
         Top             =   2430
         Width           =   16470
         Begin VB.TextBox NotaPC 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1260
            MaxLength       =   150
            TabIndex        =   123
            Top             =   570
            Width           =   14685
         End
         Begin MSFlexGridLib.MSFlexGrid GridNotas 
            Height          =   1455
            Left            =   210
            TabIndex        =   124
            Top             =   255
            Width           =   16110
            _ExtentX        =   28416
            _ExtentY        =   2566
            _Version        =   393216
            Rows            =   5
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.TextBox Nota 
         Height          =   1290
         Left            =   930
         MaxLength       =   150
         MultiLine       =   -1  'True
         TabIndex        =   120
         Top             =   375
         Width           =   10170
      End
      Begin VB.CommandButton BotaoIncluirNota 
         Caption         =   "Incluir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   915
         TabIndex        =   122
         Top             =   1725
         Width           =   1650
      End
      Begin VB.Label LabelNota 
         Caption         =   "Nota :"
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
         Left            =   345
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   125
         Top             =   390
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   8205
      Index           =   4
      Left            =   264
      TabIndex        =   91
      Top             =   792
      Visible         =   0   'False
      Width           =   16500
      Begin VB.Frame Frame5 
         Caption         =   "Distribuição dos Produtos"
         Height          =   7590
         Left            =   45
         TabIndex        =   92
         Top             =   45
         Width           =   16245
         Begin VB.TextBox DescProd 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   95
            Top             =   300
            Width           =   4000
         End
         Begin VB.ComboBox UnidadeMed 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   5400
            Style           =   2  'Dropdown List
            TabIndex        =   98
            Top             =   300
            Width           =   1395
         End
         Begin VB.ComboBox Prod 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   195
            Style           =   2  'Dropdown List
            TabIndex        =   94
            Top             =   375
            Width           =   1600
         End
         Begin MSMask.MaskEdBox ContaContabil 
            Height          =   225
            Left            =   7065
            TabIndex        =   100
            Top             =   345
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Quant 
            Height          =   225
            Left            =   6090
            TabIndex        =   99
            Top             =   330
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
         Begin MSMask.MaskEdBox CentroCusto 
            Height          =   225
            Left            =   3030
            TabIndex        =   96
            Top             =   330
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Almoxarifado 
            Height          =   225
            Left            =   4155
            TabIndex        =   97
            Top             =   330
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridDistribuicao 
            Height          =   2670
            Left            =   210
            TabIndex        =   93
            Top             =   495
            Width           =   15840
            _ExtentX        =   27940
            _ExtentY        =   4710
            _Version        =   393216
            Rows            =   7
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.CommandButton BotaoContasContabeis 
         Caption         =   "Plano de Contas"
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
         Left            =   5115
         TabIndex        =   103
         Top             =   7725
         Width           =   2145
      End
      Begin VB.CommandButton BotaoAlmoxarifado 
         Caption         =   "Almoxarifado"
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
         Left            =   75
         TabIndex        =   101
         Top             =   7710
         Width           =   1875
      End
      Begin VB.CommandButton BotaoCcl 
         Caption         =   "Centro de Custo/Lucro"
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
         Left            =   2250
         TabIndex        =   102
         Top             =   7710
         Width           =   2580
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8220
      Index           =   3
      Left            =   288
      TabIndex        =   60
      Top             =   768
      Visible         =   0   'False
      Width           =   16410
      Begin VB.Frame Frame2 
         Caption         =   "Local de Entrega"
         Height          =   2910
         Left            =   135
         TabIndex        =   61
         Top             =   105
         Width           =   8505
         Begin VB.Frame FrameTipo 
            BorderStyle     =   0  'None
            Height          =   675
            Index           =   1
            Left            =   4770
            TabIndex        =   68
            Top             =   450
            Visible         =   0   'False
            Width           =   3645
            Begin VB.ComboBox FilialFornec 
               Height          =   315
               Left            =   1230
               TabIndex        =   72
               Top             =   360
               Width           =   2160
            End
            Begin MSMask.MaskEdBox Fornec 
               Height          =   300
               Left            =   1230
               TabIndex        =   70
               Top             =   0
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin VB.Label Label40 
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
               TabIndex        =   71
               Top             =   405
               Width           =   465
            End
            Begin VB.Label FornLabel 
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
               Left            =   150
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   69
               Top             =   60
               Width           =   1035
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Tipo"
            Height          =   555
            Left            =   270
            TabIndex        =   62
            Top             =   510
            Width           =   3990
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
               Left            =   480
               TabIndex        =   63
               Top             =   225
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
               Left            =   2280
               TabIndex        =   64
               Top             =   225
               Width           =   1335
            End
         End
         Begin VB.Frame FrameTipo 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   675
            Index           =   0
            Left            =   4740
            TabIndex        =   65
            Top             =   450
            Width           =   3645
            Begin VB.ComboBox FilialEmpresa 
               Height          =   315
               Left            =   1230
               TabIndex        =   67
               Text            =   "FilialEmpresa"
               Top             =   195
               Width           =   2160
            End
            Begin VB.Label Label6 
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
               Index           =   1
               Left            =   690
               TabIndex        =   66
               Top             =   225
               Width           =   465
            End
         End
         Begin VB.Label Label73 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
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
            TabIndex        =   73
            Top             =   1560
            Width           =   915
         End
         Begin VB.Label Label72 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
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
            Left            =   3285
            TabIndex        =   77
            Top             =   1980
            Width           =   675
         End
         Begin VB.Label Label71 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
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
            TabIndex        =   81
            Top             =   2415
            Width           =   675
         End
         Begin VB.Label Label70 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
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
            TabIndex        =   75
            Top             =   1980
            Width           =   585
         End
         Begin VB.Label Label65 
            AutoSize        =   -1  'True
            Caption         =   "CEP:"
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
            Left            =   6150
            TabIndex        =   79
            Top             =   1980
            Width           =   465
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "País:"
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
            Left            =   3465
            TabIndex        =   83
            Top             =   2415
            Width           =   495
         End
         Begin VB.Label Endereco 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1260
            TabIndex        =   74
            Top             =   1500
            Width           =   6345
         End
         Begin VB.Label Bairro 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1260
            TabIndex        =   76
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label Cidade 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4020
            TabIndex        =   78
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label CEP 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6675
            TabIndex        =   80
            Top             =   1920
            Width           =   930
         End
         Begin VB.Label Estado 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1260
            TabIndex        =   82
            Top             =   2355
            Width           =   495
         End
         Begin VB.Label Pais 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4035
            TabIndex        =   84
            Top             =   2355
            Width           =   1995
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frete"
         Height          =   915
         Left            =   150
         TabIndex        =   85
         Top             =   3090
         Width           =   8505
         Begin VB.ComboBox TipoFrete 
            Height          =   315
            ItemData        =   "PedComprasGeradoOcx.ctx":25FE
            Left            =   1560
            List            =   "PedComprasGeradoOcx.ctx":2608
            Style           =   2  'Dropdown List
            TabIndex        =   87
            Top             =   390
            Width           =   1125
         End
         Begin VB.ComboBox Transportadora 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5400
            TabIndex        =   89
            Top             =   390
            Width           =   2190
         End
         Begin VB.Label Label31 
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
            Left            =   570
            TabIndex        =   86
            Top             =   450
            Width           =   945
         End
         Begin VB.Label TransportadoraLabel 
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
            Height          =   210
            Left            =   3930
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   88
            Top             =   442
            Width           =   1410
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   8145
      Index           =   2
      Left            =   195
      TabIndex        =   28
      Top             =   840
      Visible         =   0   'False
      Width           =   16560
      Begin VB.Frame Frame9 
         Caption         =   "Valores"
         Height          =   870
         Index           =   1
         Left            =   135
         TabIndex        =   44
         Top             =   6570
         Width           =   8775
         Begin MSMask.MaskEdBox ValorFrete 
            Height          =   285
            Left            =   1425
            TabIndex        =   53
            Top             =   420
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            _Version        =   393216
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorSeguro 
            Height          =   285
            Left            =   2640
            TabIndex        =   54
            Top             =   420
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            _Version        =   393216
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox OutrasDespesas 
            Height          =   285
            Left            =   3840
            TabIndex        =   55
            Top             =   420
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            _Version        =   393216
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDesc 
            Height          =   285
            Left            =   5040
            TabIndex        =   56
            Top             =   420
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            _Version        =   393216
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorIPI 
            Height          =   285
            Left            =   6270
            TabIndex        =   57
            Top             =   420
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            _Version        =   393216
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Total"
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
            Left            =   7815
            TabIndex        =   51
            Top             =   225
            Width           =   450
         End
         Begin VB.Label ValorTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7455
            TabIndex        =   58
            Top             =   420
            Width           =   1125
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "IPI"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   165
            Index           =   1
            Left            =   6675
            TabIndex        =   50
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Seguro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2865
            TabIndex        =   47
            Top             =   225
            Width           =   615
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Desconto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   5220
            TabIndex        =   49
            Top             =   225
            Width           =   825
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Frete"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1695
            TabIndex        =   46
            Top             =   225
            Width           =   450
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Despesas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3960
            TabIndex        =   48
            Top             =   225
            Width           =   840
         End
         Begin VB.Label ValorProdutos 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   225
            TabIndex        =   52
            Top             =   420
            Width           =   1125
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   405
            TabIndex        =   45
            Top             =   225
            Width           =   765
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Itens"
         Height          =   6465
         Left            =   135
         TabIndex        =   29
         Top             =   60
         Width           =   16305
         Begin VB.TextBox DescCompleta 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   975
            MaxLength       =   50
            TabIndex        =   137
            Top             =   645
            Width           =   5460
         End
         Begin MSMask.MaskEdBox TotalMoedaReal 
            Height          =   228
            Left            =   5976
            TabIndex        =   133
            Top             =   1476
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   423
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
         Begin MSMask.MaskEdBox PrecoUnitarioMoedaReal 
            Height          =   228
            Left            =   4752
            TabIndex        =   134
            Top             =   1476
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   423
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
         Begin VB.TextBox Observ 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   6225
            MaxLength       =   255
            TabIndex        =   43
            Top             =   2355
            Width           =   2565
         End
         Begin VB.TextBox DescProduto 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   1875
            MaxLength       =   50
            TabIndex        =   32
            Top             =   240
            Width           =   4000
         End
         Begin VB.ComboBox RecebForaFaixa 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2790
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   2295
            Width           =   2235
         End
         Begin MSMask.MaskEdBox PercentMaisReceb 
            Height          =   225
            Left            =   90
            TabIndex        =   37
            Top             =   2355
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
         Begin MSMask.MaskEdBox UM 
            Height          =   225
            Left            =   3405
            TabIndex        =   33
            Top             =   255
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PrecoTotal 
            Height          =   225
            Left            =   6660
            TabIndex        =   36
            Top             =   270
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
         Begin MSMask.MaskEdBox DataLimite 
            Height          =   225
            Left            =   1590
            TabIndex        =   38
            Top             =   2070
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox AliquotaICM 
            Height          =   225
            Left            =   7185
            TabIndex        =   42
            Top             =   1995
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
         Begin MSMask.MaskEdBox ValorIPIItem 
            Height          =   225
            Left            =   5970
            TabIndex        =   41
            Top             =   1995
            Width           =   960
            _ExtentX        =   1693
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PrecoUnitario 
            Height          =   225
            Left            =   5535
            TabIndex        =   35
            Top             =   270
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
         Begin MSMask.MaskEdBox AliquotaIPI 
            Height          =   225
            Left            =   3900
            TabIndex        =   39
            Top             =   2010
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   225
            Left            =   4515
            TabIndex        =   34
            Top             =   210
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
         Begin MSMask.MaskEdBox Produto 
            Height          =   225
            Left            =   570
            TabIndex        =   31
            Top             =   240
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridItens 
            Height          =   2190
            Left            =   150
            TabIndex        =   30
            Top             =   195
            Width           =   15960
            _ExtentX        =   28152
            _ExtentY        =   3863
            _Version        =   393216
            Rows            =   6
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.CommandButton BotaoGerador 
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
         Left            =   120
         Picture         =   "PedComprasGeradoOcx.ctx":2616
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   7605
         Width           =   2205
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   8145
      Index           =   1
      Left            =   225
      TabIndex        =   1
      Top             =   840
      Width           =   16560
      Begin VB.Frame Frame8 
         Caption         =   "Datas"
         Height          =   1200
         Left            =   105
         TabIndex        =   17
         Top             =   4950
         Width           =   9690
         Begin MSComCtl2.UpDown UpDownDataEnvio 
            Height          =   300
            Left            =   5310
            TabIndex        =   26
            Top             =   810
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEnvio 
            Height          =   300
            Left            =   4200
            TabIndex        =   25
            Top             =   810
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataRefFluxo 
            Height          =   300
            Left            =   7980
            TabIndex        =   140
            TabStop         =   0   'False
            Top             =   225
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataRefFluxo 
            Height          =   300
            Left            =   6885
            TabIndex        =   141
            ToolTipText     =   "Data de Referência usada em Fluxo de Caixa"
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Fluxo:"
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
            Left            =   6255
            TabIndex        =   142
            ToolTipText     =   "Data de Referência usada em Fluxo de Caixa"
            Top             =   300
            Width           =   525
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Alterado:"
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
            TabIndex        =   22
            Top             =   840
            Width           =   780
         End
         Begin VB.Label DataAlteracao 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1605
            TabIndex        =   23
            Top             =   810
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Gerado:"
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
            Left            =   870
            TabIndex        =   18
            Top             =   300
            Width           =   690
         End
         Begin VB.Label Data 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1605
            TabIndex        =   19
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Emitido:"
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
            Left            =   3465
            TabIndex        =   20
            Top             =   300
            Width           =   690
         End
         Begin VB.Label DataEmissao 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4200
            TabIndex        =   21
            Top             =   240
            Width           =   1110
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Enviado:"
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
            Left            =   3390
            TabIndex        =   24
            Top             =   870
            Width           =   765
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Cabeçalho"
         Height          =   4350
         Left            =   105
         TabIndex        =   2
         Top             =   195
         Width           =   9690
         Begin VB.TextBox ObsEmbalagem 
            Height          =   285
            Left            =   1470
            MaxLength       =   150
            TabIndex        =   135
            Top             =   2085
            Width           =   2190
         End
         Begin VB.CommandButton BotaoTrazCotacao 
            Caption         =   "$"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   8070
            Style           =   1  'Graphical
            TabIndex        =   132
            ToolTipText     =   "Numeração Automática"
            Top             =   1642
            Width           =   345
         End
         Begin VB.TextBox Contato 
            Height          =   300
            Left            =   6240
            MaxLength       =   255
            TabIndex        =   14
            Top             =   1230
            Width           =   2175
         End
         Begin VB.TextBox Observacao 
            Height          =   1470
            Index           =   0
            Left            =   1470
            MaxLength       =   255
            TabIndex        =   16
            Top             =   2535
            Width           =   6585
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   300
            Left            =   1470
            TabIndex        =   4
            Top             =   375
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Embalagem 
            Height          =   315
            Left            =   -10000
            TabIndex        =   126
            Top             =   2085
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Taxa 
            Height          =   312
            Left            =   6240
            TabIndex        =   128
            Top             =   1656
            Width           =   1824
            _ExtentX        =   3228
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            Format          =   "#,##0.00##"
            PromptChar      =   " "
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Embalagem:"
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
            TabIndex        =   136
            Top             =   2130
            Width           =   1035
         End
         Begin VB.Label Moeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1470
            TabIndex        =   131
            Top             =   1650
            Width           =   1965
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Taxa:"
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
            Left            =   5670
            TabIndex        =   130
            Top             =   1710
            Width           =   495
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Moeda:"
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
            Left            =   765
            TabIndex        =   129
            Top             =   1680
            Width           =   615
         End
         Begin VB.Label LabelEmbalagem 
            AutoSize        =   -1  'True
            Caption         =   "Embalagem:"
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
            Height          =   195
            Left            =   -10000
            TabIndex        =   127
            Top             =   2145
            Width           =   1035
         End
         Begin VB.Label ObsLabel 
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
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   15
            Top             =   2595
            Width           =   1095
         End
         Begin VB.Label CondPagtoLabel 
            AutoSize        =   -1  'True
            Caption         =   "Cond Pagto:"
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
            Left            =   345
            TabIndex        =   11
            Top             =   1275
            Width           =   1065
         End
         Begin VB.Label CondPagto 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1470
            TabIndex        =   12
            Top             =   1230
            Width           =   1965
         End
         Begin VB.Label Comprador 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6240
            TabIndex        =   6
            Top             =   375
            Width           =   2175
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
            Left            =   5205
            TabIndex        =   5
            Top             =   405
            Width           =   975
         End
         Begin VB.Label CodigoLabel 
            AutoSize        =   -1  'True
            Caption         =   "Nº Pedido:"
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
            Left            =   480
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   3
            Top             =   375
            Width           =   930
         End
         Begin VB.Label Filial 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6240
            TabIndex        =   10
            Top             =   795
            Width           =   2175
         End
         Begin VB.Label Fornecedor 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1470
            TabIndex        =   8
            Top             =   795
            Width           =   2145
         End
         Begin VB.Label Label6 
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
            Index           =   2
            Left            =   5715
            TabIndex        =   9
            Top             =   855
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
            Left            =   375
            TabIndex        =   7
            Top             =   795
            Width           =   1035
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   5445
            TabIndex        =   13
            Top             =   1290
            Width           =   735
         End
      End
      Begin VB.CommandButton BotaoPedidosGerados 
         Caption         =   "Pedidos Gerados"
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
         Left            =   105
         TabIndex        =   27
         Top             =   6330
         Width           =   2205
      End
   End
   Begin VB.CheckBox ImprimePedido 
      Caption         =   "Imprimir o pedido ao gravar"
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
      Left            =   105
      TabIndex        =   139
      Top             =   105
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   13755
      ScaleHeight     =   495
      ScaleWidth      =   3075
      TabIndex        =   113
      TabStop         =   0   'False
      Top             =   75
      Width           =   3135
      Begin VB.CommandButton BotaoEmail 
         Height          =   360
         Left            =   60
         Picture         =   "PedComprasGeradoOcx.ctx":4CF0
         Style           =   1  'Graphical
         TabIndex        =   138
         ToolTipText     =   "Enviar email"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   1575
         Picture         =   "PedComprasGeradoOcx.ctx":5692
         Style           =   1  'Graphical
         TabIndex        =   116
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoBaixar 
         Height          =   360
         Left            =   1065
         Picture         =   "PedComprasGeradoOcx.ctx":57EC
         Style           =   1  'Graphical
         TabIndex        =   115
         ToolTipText     =   "Baixar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   2070
         Picture         =   "PedComprasGeradoOcx.ctx":5952
         Style           =   1  'Graphical
         TabIndex        =   117
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   2580
         Picture         =   "PedComprasGeradoOcx.ctx":5E84
         Style           =   1  'Graphical
         TabIndex        =   118
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   570
         Picture         =   "PedComprasGeradoOcx.ctx":6002
         Style           =   1  'Graphical
         TabIndex        =   114
         ToolTipText     =   "Imprimir"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   8670
      Left            =   75
      TabIndex        =   0
      Top             =   420
      Width           =   16800
      _ExtentX        =   29633
      _ExtentY        =   15293
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pedido"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Itens"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Entrega"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Distribuição"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Bloqueios"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Notas"
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
Attribute VB_Name = "PedComprasGeradoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iGrid_Produto_Col As Integer
Dim iGrid_DescricaoProduto_Col As Integer
Dim iGrid_UnidadeMed_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_QuantRecebida_Col As Integer
Dim iGrid_PrecoUnitario_Col As Integer
Dim iGrid_PercentDesc_Col As Integer
Dim iGrid_PrecoTotal_Col As Integer
'Alteracao Daniel
Dim iGrid_PrecoUnitarioMoedaReal_Col As Integer
Dim iGrid_TotalMoedaReal_Col As Integer
'Fim da Alteracao Daniel
Dim iGrid_DataLimite_Col As Integer
Dim iGrid_AliquotaIPI_Col As Integer
Dim iGrid_ValorIPIItem_Col As Integer
Dim iGrid_AliquotaICMS_Col As Integer
Dim iGrid_PercentMaisReceb_Col As Integer
Dim iGrid_RecebForaFaixa_Col As Integer
Dim iGrid_Observacao_Col As Integer
Dim iGrid_DescCompleta_Col As Integer 'leo

Private WithEvents objEventoNotaPC As AdmEvento
Attribute objEventoNotaPC.VB_VarHelpID = -1

Dim iGrid_Prod_Col As Integer
Dim iGrid_DescProduto_Col As Integer
Dim iGrid_CentroCusto_Col As Integer
Dim iGrid_Almoxarifado_Col As Integer
Dim iGrid_UM_Col As Integer
Dim iGrid_Quant_Col As Integer
Dim iGrid_ContaContabil_Col As Integer

Dim iGrid_TipoBloqueio_Col As Integer
Dim iGrid_DataBloqueio_Col As Integer
Dim iGrid_CodUsuario_Col As Integer
Dim iGrid_ResponsavelBL_Col As Integer
Dim iGrid_DataLiberacao_Col As Integer
Dim iGrid_ResponsavelLib_Col As Integer

Dim iFrameAtual As Integer
Dim iAlterado As Integer
Dim iFornecAlterado As Integer
Dim iFornecedorAlterado As Integer
Dim iClienteAlterado As Integer
Dim iFrameTipoDestinoAtual As Integer
Dim iGridBloqueiosLinhasBloqueadas As Integer
Dim objGridItens As AdmGrid
Dim objGridDistribuicao As AdmGrid
Dim objGridBloqueio As AdmGrid
Dim gcolItemPedido As Collection
Dim giVerificacaoBloqueioAlcada As Integer
Dim bExibirColReal As Boolean

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoFornec As AdmEvento
Attribute objEventoFornec.VB_VarHelpID = -1
Private WithEvents objEventoBotaoCcl As AdmEvento
Attribute objEventoBotaoCcl.VB_VarHelpID = -1
Private WithEvents objEventoObservacao As AdmEvento
Attribute objEventoObservacao.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1
Private WithEvents objEventoAlmoxarifados As AdmEvento
Attribute objEventoAlmoxarifados.VB_VarHelpID = -1
Private WithEvents objEventoTransportadora As AdmEvento
Attribute objEventoTransportadora.VB_VarHelpID = -1
Private WithEvents objEventoBotaoPedidosGerados As AdmEvento
Attribute objEventoBotaoPedidosGerados.VB_VarHelpID = -1

Dim dTaxaInformada As Double 'leo
Private WithEvents objEventoEmbalagens As AdmEvento 'leo
Attribute objEventoEmbalagens.VB_VarHelpID = -1

'Grid de Notas '??? POR LEO
Dim objGridNotas As AdmGrid
Dim iGrid_NotaPC_Col As Integer


Private Sub AliquotaICM_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub AliquotaICM_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub AliquotaIPI_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub AliquotaIPI_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub BotaoEmail_Click()

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras
Dim objRelatorio As New AdmRelatorio
Dim sMailTo As String, sFiltro As String
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objEndereco As New ClassEndereco
Dim objBloqueioPC As ClassBloqueioPC, sInfoEmail As String

On Error GoTo Erro_BotaoEmail_Click

    lErro = Move_Tela_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then gError 53305

    If objPedidoCompra.lCodigo = 0 Then gError 76054

    'Verifica se o Pedido de Compra informado existe
    lErro = CF("PedidoCompra_Le_Numero", objPedidoCompra)
    If lErro <> SUCESSO And lErro <> 56142 Then gError 76027
    
    'Se o Pedido não existe ==> erro
    If lErro = 56142 Then gError 76028
    
    If objPedidoCompra.dtDataRegAprov = DATA_NULA Then
        If gobjCOM.iPedCompraBloqEnvioSemAprov = MARCADO Then gError 213170
    End If
        
    lErro = CF("BloqueiosPC_Le", objPedidoCompra)
    If lErro <> SUCESSO Then gError 76056
    
    'If objPedidoCompra.colBloqueiosPC.Count > 0 Then gError 76050
    
    For Each objBloqueioPC In objPedidoCompra.colBloqueiosPC
            
        If objBloqueioPC.dtDataLib = DATA_NULA Then gError 76050
    
    Next
    
    If objPedidoCompra.lFornecedor <> 0 And objPedidoCompra.iFilial <> 0 Then

        objFilialFornecedor.lCodFornecedor = objPedidoCompra.lFornecedor
        objFilialFornecedor.iCodFilial = objPedidoCompra.iFilial

        lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 12929 Then gError 129314
         
        If lErro = SUCESSO Then
        
            objEndereco.lCodigo = objFilialFornecedor.lEndereco
            
            lErro = CF("Endereco_Le", objEndereco)
            If lErro <> SUCESSO Then gError 129315
        
            sMailTo = objEndereco.sEmail
            
        End If
        
        sInfoEmail = "Fornecedor: " & CStr(objFilialFornecedor.lCodFornecedor) & " - " & Fornecedor.Caption & " . Filial: " & Filial.Caption
        
    End If
    
    If Len(Trim(sMailTo)) = 0 Then gError 129316
    
    'Preenche a Data de Entrada com a Data Atual
    DataEmissao.Caption = Format(gdtDataHoje, "dd/mm/yy")

    objPedidoCompra.dtDataEmissao = gdtDataHoje

    'Atualiza data de emissao no BD para a data atual
    lErro = CF("PedidoCompra_Atualiza_DataEmissao", objPedidoCompra, True)
    If lErro <> SUCESSO And lErro <> 56348 Then gError 53306

    'se nao encontrar ---> erro
    If lErro = 56348 Then gError 53307

    sFiltro = "REL_PCOM.PC_NumIntDoc = @NPEDCOM"
    lErro = CF("Relatorio_ObterFiltro", "Pedido de Compra Gerado", sFiltro)
    If lErro <> SUCESSO Then gError 76029
    
    'Executa o relatório
    lErro = objRelatorio.ExecutarDiretoEmail("Pedido de Compra Gerado", sFiltro, 0, "PEDCOM", "NPEDCOM", objPedidoCompra.lNumIntDoc, "TTO_EMAIL", sMailTo, "TSUBJECT", "Pedido de Compra " & CStr(objPedidoCompra.lCodigo), "TALIASATTACH", "PedCompra" & CStr(objPedidoCompra.lCodigo), "TINFO_EMAIL", sInfoEmail)
    If lErro <> SUCESSO Then gError 76029
    
    Exit Sub

Erro_BotaoEmail_Click:

    Select Case gErr
    
        Case 53305, 53306, 76056, 129314, 129315

        Case 53307
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", gErr, objPedidoCompra.lCodigo)

        Case 76021
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
            
        Case 76027, 76029
        
        Case 76028
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", gErr, objPedidoCompra.lCodigo)
        
        Case 76050
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_BLOQUEADO", gErr, objPedidoCompra.lCodigo)
            
        Case 76054
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDCOMPRA_IMPRESSAO", gErr)
            
        Case 129316
            Call Rotina_Erro(vbOKOnly, "ERRO_EMAIL_NAO_ENCONTRADO", gErr, objPedidoCompra.lCodigo)
            
        Case 213170
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_APROVADO", gErr, objPedidoCompra.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164407)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGerador_Click()
'Chama a tela de ConcorrenciaCons ou PedCotacaoCons, de acordo com
'a forma de geracao do Pedido de Compra (por Concorrencia ou PedidoCotacao)

Dim lErro As Long
Dim objItemPC As New ClassItemPedCompra
Dim objPedidoCotacao As New ClassPedidoCotacao
Dim objConcorrencia As New ClassConcorrencia
Dim objPedidoCompra As New ClassPedidoCompras
Dim lNumInt As Long
Dim objItemPedCotacao As New ClassItemPedCotacao
Dim objCotacaoItemConc As New ClassCotacaoItemConc


On Error GoTo Erro_BotaoGerador_Click

    If objGridItens.iLinhasExistentes = 0 Then Exit Sub

    Set objItemPC = New ClassItemPedCompra
    lNumInt = gcolItemPedido.Item(1)
    
    objItemPC.lNumIntDoc = lNumInt
    
    objPedidoCompra.lCodigo = StrParaLong(Codigo.Text)
    objPedidoCompra.iFilialEmpresa = giFilialEmpresa
    
    'Lê o Pedido de Compra
    lErro = CF("PedidoCompra_Le_Todos", objPedidoCompra)
    If lErro <> SUCESSO Then gError 68478
    
    'Lê os Itens do Pedido de Compra
    lErro = CF("ItensPC_LeTodos", objPedidoCompra)
    If lErro <> SUCESSO Then gError 68479
    
    Set objItemPC = objPedidoCompra.colItens.Item(1)
    'Verifica se NumIntDocOrigem está preenchido
    If objItemPC.lNumIntOrigem = 0 Then gError 68442

    'Verifica se o ItemPC não tem origem (TipoOrigem=0)
    If objItemPC.iTipoOrigem = 0 Then gError 68443

    If objItemPC.iTipoOrigem = TIPO_ITEMPEDCOTACAO Then
        
        objItemPedCotacao.lNumIntDoc = objItemPC.lNumIntOrigem
        objPedidoCotacao.colItens.Add objItemPedCotacao
        
        'Lê o PedidoCotacao cujo NumIntDoc do ItemPedCotacao foi fornecido
        lErro = CF("ItemPedCotacao_Le_PedidoCotacao", objItemPedCotacao, objPedidoCotacao)
        If lErro <> SUCESSO Then gError 68496
        
        Call Chama_Tela("PedidoCotacaoCons", objPedidoCotacao)

    ElseIf objItemPC.iTipoOrigem = TIPO_COTACAOITEMCONCORRENCIA Then

        objCotacaoItemConc.lNumIntDoc = objItemPC.lNumIntOrigem
            
        'Lê a Concorrencia a partir do NumIntDoc de CotacaoItemConcorrencia
        lErro = CF("ItensPedCompra_Le_CotacaoItemConcorrencia", objCotacaoItemConc, objConcorrencia)
        If lErro <> SUCESSO Then gError 74861
        
        Call Chama_Tela("ConcorrenciaCons", objConcorrencia)

    End If
 
    Exit Sub

Erro_BotaoGerador_Click:

    Select Case gErr

        Case 68442, 68443
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDCOMPRA_NAO_GERADO", gErr, objPedidoCompra.lCodigo)

        Case 68478, 68479, 74861
            'Erros tratados nas rotinas chamadas
        
        Case 68496
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOTACAO_NAO_ENCONTRADO1", gErr, objPedidoCompra.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164408)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLiberaBloqueio_Click()

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras

On Error GoTo Erro_BotaoLiberaBloqueio_Click

    'Verifica se o número do Pedido de Compra está preenchido
    If Len(Trim(Codigo.Text)) > 0 Then

        'Recolhe os dados da tela
        lErro = Move_Tela_Memoria(objPedidoCompra)
        If lErro <> SUCESSO Then gError 74866

        'Chama tela LiberaBloqueioPC
        Call Chama_Tela("LiberaBloqueioPC", objPedidoCompra)

    End If

    Exit Sub

Erro_BotaoLiberaBloqueio_Click:

    Select Case gErr

        Case 74866

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164409)

    End Select

    Exit Sub

End Sub

Private Sub DataEnvio_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataEnvio, iAlterado)
    
End Sub

Private Sub DataEnvio_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dtDataEnvio As Date
Dim dtData As Date

On Error GoTo Erro_DataEnvio_Validate

    'Verifica se a data foi preenchida
    If Len(DataEnvio.ClipText) > 0 Then

        'Critica a data informada
        lErro = Data_Critica(DataEnvio.Text)
        If lErro <> SUCESSO Then Error 53245

        If Len(Trim(Data.Caption)) = 0 Then
            If (CDate(DataEnvio.Text) < DATA_NULA) Then Error 53246
        Else
            If (CDate(DataEnvio.Text) < CDate(Data.Caption)) Then Error 53246
        End If

    End If

    Exit Sub

Erro_DataEnvio_Validate:

    Cancel = True

    Select Case Err

        Case 53245

        Case 53246
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAENVIO_INFERIOR_DATAPEDIDO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164410)

    End Select

    Exit Sub

End Sub

Private Sub DescProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub DescProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub DescProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DescProduto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub AliquotaIPI_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub AliquotaIPI_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = AliquotaIPI
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub AliquotaICM_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub AliquotaICM_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = AliquotaICM
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Embalagem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Embalagem_Validate(Cancel As Boolean)
'leo
Dim objEmbalagem As New ClassEmbalagem
Dim lErro As Long

On Error GoTo Erro_Embalagem_Validate
        
    If Len(Trim(Embalagem.Text)) = 0 Then Exit Sub
    
    'se o texto é numérico ou se está no formato Código + Separador + Sigla
    If IsNumeric(Embalagem.Text) = True Or Codigo_Extrai(Embalagem.Text) <> 0 Then
        
        If IsNumeric(Embalagem.Text) Then
        
            objEmbalagem.iCodigo = StrParaInt(Embalagem.Text)
        
        Else
            
            objEmbalagem.iCodigo = Codigo_Extrai(Embalagem.Text)
            
        End If
            
        'Lê a embalagem através do Código
        lErro = CF("Embalagem_Le", objEmbalagem)
        If lErro <> SUCESSO And lErro <> 82763 Then gError 103394
        
        If lErro <> SUCESSO Then gError 103395
    
    'Se é uma String
    Else
        
        objEmbalagem.sSigla = Embalagem.Text
    
        'Lê a Embalagem através da sigla
        lErro = CF("Embalagem_Le_Sigla", objEmbalagem)
        If lErro <> SUCESSO And lErro <> 95088 Then gError 103396
        
        If lErro <> SUCESSO Then gError 103397
        
    End If
    
    Embalagem.Text = objEmbalagem.sSigla
    
    Exit Sub
    
Erro_Embalagem_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 103394, 103396
            
        Case 103395, 103397
            Call Rotina_Erro(vbOKOnly, "ERRO_EMBALAGEM_INEXISTENTE", gErr, Embalagem.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164411)
    
    End Select
    
    Exit Sub

End Sub

Private Sub LabelEmbalagem_Click()

'leo
Dim objEmbalagem As New ClassEmbalagem
Dim colSelecao As New Collection
Dim lErro As Long

On Error GoTo Erro_LabelEmbalagem_Click
        
    objEmbalagem.sSigla = Embalagem.Text
    
    lErro = CF("Embalagem_Le_Sigla", objEmbalagem)
    If lErro <> SUCESSO And lErro <> 95088 Then gError 103399
    
    'chama a tela de browser
    Call Chama_Tela("EmbalagensLista", colSelecao, objEmbalagem, objEventoEmbalagens)
    
    Exit Sub
    
Erro_LabelEmbalagem_Click:
    
    Select Case gErr
    
        Case 103399
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMBALAGEM", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164412)
    
    End Select
    
    Exit Sub

End Sub

Private Sub LabelNota_Click()

Dim objNotasPC As New ClassNotaPC
Dim colSelecao As New Collection

    'Preenche objFornecedor com NomeReduzido da tela
    objNotasPC.sNota = Nota.Text

    Call Chama_Tela("NotasPCLista", colSelecao, objNotasPC, objEventoNotaPC)

End Sub

Private Sub objEventoEmbalagens_evSelecao(obj1 As Object)
'leo
Dim objEmbalagem As ClassEmbalagem

On Error GoTo Erro_objEventoEmbalagens_evSelecao
                       
    'Define o tipo de obj recebido (Tipo Embalagem)
    Set objEmbalagem = obj1
        
    Embalagem.Text = objEmbalagem.sSigla
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoEmbalagens_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164413)
              
    End Select
    
    Exit Sub

End Sub

Private Sub Observ_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Observ_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub Observ_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub Observ_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Observ
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Prod_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Prod_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Prod_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDistribuicao)

End Sub

Private Sub Prod_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDistribuicao)

End Sub

Private Sub Prod_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDistribuicao.objControle = Prod
    lErro = Grid_Campo_Libera_Foco(objGridDistribuicao)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub DescProd_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescProd_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescProd_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDistribuicao)

End Sub

Private Sub DescProd_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDistribuicao)

End Sub

Private Sub DescProd_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDistribuicao.objControle = DescProd
    lErro = Grid_Campo_Libera_Foco(objGridDistribuicao)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub CentroCusto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CentroCusto_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CentroCusto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDistribuicao)

End Sub

Private Sub CentroCusto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDistribuicao)

End Sub

Private Sub CentroCusto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDistribuicao.objControle = CentroCusto
    lErro = Grid_Campo_Libera_Foco(objGridDistribuicao)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub ContaContabil_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaContabil_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDistribuicao)

End Sub

Private Sub ContaContabil_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDistribuicao)

End Sub

Private Sub ContaContabil_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDistribuicao.objControle = ContaContabil
    lErro = Grid_Campo_Libera_Foco(objGridDistribuicao)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Almoxarifado_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Almoxarifado_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDistribuicao)

End Sub

Private Sub Almoxarifado_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDistribuicao)

End Sub

Private Sub Almoxarifado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDistribuicao.objControle = Almoxarifado
    lErro = Grid_Campo_Libera_Foco(objGridDistribuicao)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Taxa_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Taxa_Validate(Cancel As Boolean)

Dim lErro As Long
    
On Error GoTo Erro_Taxa_Validate
    
    'se a taxa já foi informada e está sendo alterada, proibe a alteração
    If dTaxaInformada <> 0 And dTaxaInformada <> StrParaDbl(Taxa.Text) Then gError 103404
        
    'Verifica se algum valor foi digitado
    If Len(Trim(Taxa.Text)) > 0 Then

        'Critica se é valor Positivo
        lErro = Valor_Positivo_Critica_Double(Taxa.Text)
        If lErro <> SUCESSO Then gError 103398
    
        'Põe o valor formatado na tela
        Taxa.Text = Format(Taxa.Text, FORMATO_TAXA_CONVERSAO_MOEDA)
        
        'Calcula o comparativo em real para o grid de itens
        Call ComparativoMoedaReal_Calcula(StrParaDbl(Taxa.Text))
    
    End If
            
    Exit Sub
    
Erro_Taxa_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 103398
        
        Case 103404
            Call Rotina_Erro(vbOKOnly, "ERRO_TAXA_NAO_ALTERADA", gErr)
            Taxa.Text = Format(dTaxaInformada, FORMATO_TAXA_CONVERSAO_MOEDA)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164414)

    End Select

    Exit Sub

End Sub

Private Sub UnidadeMed_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UnidadeMed_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UnidadeMed_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDistribuicao)

End Sub

Private Sub UnidadeMed_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDistribuicao)

End Sub

Private Sub UnidadeMed_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDistribuicao.objControle = UnidadeMed
    lErro = Grid_Campo_Libera_Foco(objGridDistribuicao)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Quant_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub
Private Sub Quant_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDistribuicao)

End Sub

Private Sub Quant_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDistribuicao)

End Sub

Private Sub Quant_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDistribuicao.objControle = Quant
    lErro = Grid_Campo_Libera_Foco(objGridDistribuicao)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub TipoBloqueio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoBloqueio_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoBloqueio_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridBloqueio)

End Sub

Private Sub TipoBloqueio_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBloqueio)

End Sub

Private Sub TipoBloqueio_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = TipoBloqueio
    lErro = Grid_Campo_Libera_Foco(objGridBloqueio)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub ResponsavelBL_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ResponsavelBL_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridBloqueio)

End Sub

Private Sub ResponsavelBL_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBloqueio)

End Sub

Private Sub ResponsavelBL_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBloqueio.objControle = ResponsavelBL
    lErro = Grid_Campo_Libera_Foco(objGridBloqueio)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub RecebForaFaixa_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub RecebForaFaixa_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub RecebForaFaixa_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub RecebForaFaixa_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub RecebForaFaixa_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = RecebForaFaixa
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub PercentMaisReceb_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PercentMaisReceb_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub PercentMaisReceb_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub PercentMaisReceb_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PercentMaisReceb
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Almoxarifado Then
            Call BotaoAlmoxarifado_Click
        ElseIf Me.ActiveControl Is ContaContabil Then
            Call BotaoContasContabeis_Click
        ElseIf Me.ActiveControl Is CentroCusto Then
            Call BotaoCcl_Click
        ElseIf Me.ActiveControl Is Transportadora Then
            Call TransportadoraLabel_Click
        ElseIf Me.ActiveControl Is Fornec Then
            Call FornLabel_Click
        ElseIf Me.ActiveControl Is Codigo Then
            Call CodigoLabel_Click
        End If
    End If

End Sub

Private Sub ValorIPIItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorIPIItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub ValorIPIItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub ValorIPIItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = ValorIPIItem
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub DataLimite_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataLimite_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub DataLimite_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub DataLimite_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DataLimite
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub FilialEmpresa_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_FilialEmpresa_Validate

    'Verifica se FilialEmpresa foi preenchida
    If Len(Trim(FilialEmpresa.Text)) > 0 Then

        If FilialEmpresa.ListIndex <> -1 Then Exit Sub

        lErro = Combo_Seleciona(FilialEmpresa, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 53262

        'Se nao encontra o ítem com o código informado
        If lErro = 6730 Then

            'preeenche objFilialEmpresa
            objFilialEmpresa.iCodFilial = iCodigo

            'Le a FilialEmpresa
            lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
            If lErro <> SUCESSO And lErro <> 27378 Then Error 53263

            'Se nao encontrou => erro
            If lErro = 27378 Then Error 53264

            If lErro = SUCESSO Then

                'Coloca na tela o codigo e o nome da FilialEmpresa
                FilialEmpresa.Text = objFilialEmpresa.lCodEmpresa & SEPARADOR & objFilialEmpresa.sNome

                If objFilialEmpresa.objEnderecoEntrega.lCodigo <> 0 Then
                    Call Preenche_Endereco(objFilialEmpresa.objEnderecoEntrega)
                Else
                    Call Preenche_Endereco(objFilialEmpresa.objEndereco)
                End If

                lErro = Distribuicao_Grid_Atualiza()
                If lErro <> SUCESSO Then Error 53265

            End If

        End If

        'Se nao encontrou e nao era codigo
        If lErro = 6731 Then Error 53266

    Else

        'Limpa frame de Endereco
        Call Limpa_Frame_Endereco

    End If

    Exit Sub

Erro_FilialEmpresa_Validate:

    Cancel = True

    Select Case Err

        Case 53262, 53263, 53265

        Case 53264
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err, iCodigo)

        Case 53266
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA1", Err, objFilialEmpresa.sNome)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164415)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim objComprador As New ClassComprador
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    
    bExibirColReal = True

    objComprador.sCodUsuario = gsUsuario

    iGridBloqueiosLinhasBloqueadas = 0

    '#####################################
    'Inserido por Wagner
    Call Formata_Controles
    '#####################################

    'Verifica se gsUsuario e comprador
    lErro = CF("Comprador_Le_Usuario", objComprador)
    If lErro <> SUCESSO And lErro <> 50059 Then gError 53173
    If lErro = 50059 Then gError 53174

    objUsuarios.sCodUsuario = objComprador.sCodUsuario

    'le  o usuário contido na tabela de Usuarios
    lErro = CF("Usuarios_Le", objUsuarios)
    If lErro <> SUCESSO And lErro <> 40832 Then gError 53175
    If lErro = 40832 Then gError 53176

    'Coloca nome reduzido do Comprador na tela
    Comprador.Caption = objUsuarios.sNomeReduzido

    'Carrega a combo de Transportadora
    lErro = Carrega_Transportadora()
    If lErro <> SUCESSO Then gError 53177

    'Carrega a combo de Tipos de Bloqueio
    lErro = Carrega_TipoBloqueio()
    If lErro <> SUCESSO Then gError 53178

    'Carrega a combo de filial
    lErro = Carrega_FilialEmpresa()
    If lErro <> SUCESSO Then gError 53393

    'Carrega a combo de RecebForaFaixa
    lErro = Carrega_RecebForaFaixa()
    If lErro <> SUCESSO Then gError 53179

    'Inicializa mascara do produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 53182

    'Inicializa mascara do Ccl
    lErro = Inicializa_MascaraCcl()
    If lErro <> SUCESSO Then gError 53183

    'Inicializa mascara da conta contabil
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabil)
    If lErro <> SUCESSO Then gError 53184

    Quant.Format = FORMATO_ESTOQUE

    Set gcolItemPedido = New Collection

    Set objGridItens = New AdmGrid
    Set objGridDistribuicao = New AdmGrid
    Set objGridBloqueio = New AdmGrid
    Set objEventoEmbalagens = New AdmEvento 'leo
    Set objEventoTransportadora = New AdmEvento
    Set objEventoAlmoxarifados = New AdmEvento
    Set objEventoBotaoCcl = New AdmEvento
    Set objEventoContaContabil = New AdmEvento
    Set objEventoFornec = New AdmEvento
    Set objEventoCodigo = New AdmEvento
    Set objEventoObservacao = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoBotaoPedidosGerados = New AdmEvento
    Set objEventoNotaPC = New AdmEvento
    
    Set objGridNotas = New AdmGrid
    
    'Faz a inicializacao do grid itens
    lErro = Inicializa_Grid_Itens(objGridItens)
    If lErro <> SUCESSO Then gError 53185

    'Faz a inicializacao do grid distribuicao
    lErro = Inicializa_Grid_Distribuicao(objGridDistribuicao)
    If lErro <> SUCESSO Then gError 53186

    'Faz a inicializacao do grid bloqueio
    lErro = Inicializa_Grid_Bloqueios(objGridBloqueio)
    If lErro <> SUCESSO Then gError 53187
    
    '??? por leo
    lErro = Inicializa_GridNotas(objGridNotas)
    If lErro <> SUCESSO Then gError 103353

    
    'seleciona tipo de destino empresa
    TipoDestino(TIPO_DESTINO_EMPRESA).Value = True
    Call CF("Filial_Seleciona", FilialEmpresa, giFilialEmpresa)

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 53173, 53175, 53177, 53178, 53179, 53182, 53183, 53184, 53185, 53186, 53187, 53393, 103353

        Case 53174
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_COMPRADOR", gErr, objComprador.sCodUsuario)

        Case 53176
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", gErr, objUsuarios.sCodUsuario)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164416)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Private Sub Fornec_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Fornec_Validate

    If iFornecAlterado <> 0 Then

        'Verifica se Fornec esta preenchido
        If Len(Trim(Fornec.Text)) <> 0 Then

            'Le Fornec
            lErro = TP_Fornecedor_Le(Fornec, objFornecedor, iCodFilial)
            If lErro <> SUCESSO Then Error 53267

            'Le as filiais do Fornec
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO And lErro <> 6698 Then Error 53268

            'Preenche a combo de filiais
            Call CF("Filial_Preenche", FilialFornec, colCodigoNome)

            'Seleciona a filial na combo de filiais
            Call CF("Filial_Seleciona", FilialFornec, iCodFilial)

        Else
            'Limpa a combobox
            FilialFornec.Clear
            FilialFornec.Text = ""

        End If

    End If

    iFornecAlterado = 0

    Exit Sub

Erro_Fornec_Validate:

    Cancel = True

    Select Case Err

        Case 53267, 53268

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164417)

    End Select

    Exit Sub

End Sub

Private Sub GridBloqueios_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDelete Then
        If GridBloqueios.Row <= iGridBloqueiosLinhasBloqueadas Then Exit Sub
    End If

    Call Grid_Trata_Tecla1(KeyCode, objGridBloqueio)

    Exit Sub

End Sub

Private Sub Observacao_Change(Index As Integer)

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub OutrasDespesas_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dOutrasDespesas As Double

On Error GoTo Erro_OutrasDespesas_Validate

    'Verifica se o OutrasDespesas esta preenchido
    If Len(Trim(OutrasDespesas.Text)) > 0 Then

        'Critica o valor informado
        lErro = Valor_NaoNegativo_Critica(OutrasDespesas.Text)
        If lErro <> SUCESSO Then Error 53253

        dOutrasDespesas = StrParaDbl(OutrasDespesas.Text)

        'Coloca o valor no formato standard da tela
        OutrasDespesas.Text = Format(dOutrasDespesas, "Standard")

    End If

    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then Error 53254

    Exit Sub

Erro_OutrasDespesas_Validate:

    Cancel = True

    Select Case Err

        Case 53253, 53254

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164418)

    End Select

    Exit Sub

End Sub

Private Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Produto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub
Private Sub Quantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub PrecoUnitario_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PrecoUnitario_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub PrecoUnitario_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub PrecoUnitario_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PrecoUnitario
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub
Private Sub TabStrip1_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index = iFrameAtual Then Exit Sub

    If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

    'Torna Frame correspondente ao Tab selecionado visivel
    Frame1(TabStrip1.SelectedItem.Index).Visible = True
    'Torna Frame atual visivel
    Frame1(iFrameAtual).Visible = False
    'Armazena novo valor de iFrameAtual
    iFrameAtual = TabStrip1.SelectedItem.Index

End Sub

Private Function Carrega_Transportadora() As Long

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_Transportadora

    'Lê o Código e o NOme de Todas as Transportadoras do BD
    lErro = CF("Cod_Nomes_Le", "Transportadoras", "Codigo", "NomeReduzido", STRING_NOME_TABELA, colCodigoNome)
    If lErro <> SUCESSO Then Error 53180

    'Carrega a combo de Transportadoras
    For Each objCodigoNome In colCodigoNome
        Transportadora.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        Transportadora.ItemData(Transportadora.NewIndex) = objCodigoNome.iCodigo
    Next

    Carrega_Transportadora = SUCESSO

    Exit Function

Erro_Carrega_Transportadora:

    Carrega_Transportadora = Err

    Select Case Err

        Case 53180

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164419)

    End Select

    Exit Function

End Function

Private Function Carrega_TipoBloqueio() As Long

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_TipoBloqueio

    'Lê o Código e o NOme de Todas os Tipos de Bloqueio do BD
    lErro = CF("Cod_Nomes_Le", "TiposDeBloqueioPC", "Codigo", "NomeReduzido", STRING_NOME_TABELA, colCodigoNome)
    If lErro <> SUCESSO Then Error 53181

    'Carrega a combo de Tipo de Bloqueio
    For Each objCodigoNome In colCodigoNome
        If objCodigoNome.iCodigo <> BLOQUEIO_ALCADA Then

            TipoBloqueio.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
            TipoBloqueio.ItemData(TipoBloqueio.NewIndex) = objCodigoNome.iCodigo

        End If
    Next

    Carrega_TipoBloqueio = SUCESSO

    Exit Function

Erro_Carrega_TipoBloqueio:

    Carrega_TipoBloqueio = Err

    Select Case Err

        Case 53181

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164420)

    End Select

    Exit Function

End Function

Private Function Carrega_FilialEmpresa() As Long

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_FilialEmpresa

    'Lê o Código e o NOme de Toda FilialEmpresa do BD
    lErro = CF("Cod_Nomes_Le_FilEmp", colCodigoNome)
    If lErro <> SUCESSO Then Error 53394

    'Carrega a combo de Filial Empresa
    For Each objCodigoNome In colCodigoNome
        FilialEmpresa.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        FilialEmpresa.ItemData(FilialEmpresa.NewIndex) = objCodigoNome.iCodigo
    Next

    Carrega_FilialEmpresa = SUCESSO

    Exit Function

Erro_Carrega_FilialEmpresa:

    Carrega_FilialEmpresa = Err

    Select Case Err

        Case 53394

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164421)

    End Select

    Exit Function

End Function

Private Function Carrega_RecebForaFaixa() As Long

    'Limpa a combo
    RecebForaFaixa.Clear

    RecebForaFaixa.AddItem MENSAGEM_NAO_AVISA_ACEITA_RECEBIMENTO
    RecebForaFaixa.ItemData(RecebForaFaixa.NewIndex) = NAO_AVISA_E_ACEITA_RECEBIMENTO

    RecebForaFaixa.AddItem MENSAGEM_REJEITA_RECEBIMENTO
    RecebForaFaixa.ItemData(RecebForaFaixa.NewIndex) = ERRO_E_REJEITA_RECEBIMENTO

    RecebForaFaixa.AddItem MENSAGEM_ACEITA_RECEBIMENTO
    RecebForaFaixa.ItemData(RecebForaFaixa.NewIndex) = AVISA_E_ACEITA_RECEBIMENTO

    Exit Function

End Function

Private Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Itens

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("Preço Total")
    'Se a moeda for Diferente de Real => Exibe as Colunas de Comparacao
    If bExibirColReal = True Then
        objGridInt.colColuna.Add ("Preço (R$)")
        objGridInt.colColuna.Add ("Total (R$)")
    End If
    objGridInt.colColuna.Add ("Data de Entrega")
    objGridInt.colColuna.Add ("Alíquota IPI")
    objGridInt.colColuna.Add ("Valor IPI ")
    objGridInt.colColuna.Add ("Alíquota ICMS")
    objGridInt.colColuna.Add ("% a Mais Receb")
    objGridInt.colColuna.Add ("Ação Receb Fora Faixa")
    objGridInt.colColuna.Add ("Observação")
    objGridInt.colColuna.Add ("Desc. Completa") 'leo

    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescProduto.Name)
    objGridInt.colCampo.Add (UM.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (PrecoUnitario.Name)
    objGridInt.colCampo.Add (PrecoTotal.Name)
    'Se a moeda for Diferente de Real => Exibe as Colunas de Comparacao
    If bExibirColReal = True Then
        objGridInt.colCampo.Add (PrecoUnitarioMoedaReal.Name)
        objGridInt.colCampo.Add (TotalMoedaReal.Name)
    End If
    objGridInt.colCampo.Add (DataLimite.Name)
    objGridInt.colCampo.Add (AliquotaIPI.Name)
    objGridInt.colCampo.Add (ValorIPIItem.Name)
    objGridInt.colCampo.Add (AliquotaICM.Name)
    objGridInt.colCampo.Add (PercentMaisReceb.Name)
    objGridInt.colCampo.Add (RecebForaFaixa.Name)
    objGridInt.colCampo.Add (Observ.Name)
    objGridInt.colCampo.Add (DescCompleta.Name) 'leo

   'indica onde estao situadas as colunas do grid
    iGrid_Produto_Col = 1
    iGrid_DescProduto_Col = 2
    iGrid_UnidadeMed_Col = 3
    iGrid_Quantidade_Col = 4
    iGrid_PrecoUnitario_Col = 5
    iGrid_PrecoTotal_Col = 6
    
    If bExibirColReal = True Then
        
        iGrid_PrecoUnitarioMoedaReal_Col = 7
        iGrid_TotalMoedaReal_Col = 8
        iGrid_DataLimite_Col = 9
        iGrid_AliquotaIPI_Col = 10
        iGrid_ValorIPIItem_Col = 11
        iGrid_AliquotaICMS_Col = 12
        iGrid_PercentMaisReceb_Col = 13
        iGrid_RecebForaFaixa_Col = 14
        iGrid_Observacao_Col = 15
        iGrid_DescCompleta_Col = 16
        
    Else
    
        iGrid_DataLimite_Col = 7
        iGrid_AliquotaIPI_Col = 8
        iGrid_ValorIPIItem_Col = 9
        iGrid_AliquotaICMS_Col = 10
        iGrid_PercentMaisReceb_Col = 11
        iGrid_RecebForaFaixa_Col = 12
        iGrid_Observacao_Col = 13
        iGrid_DescCompleta_Col = 14
    
    End If

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridItens

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_PEDIDO_COMPRAS + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 16

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'proibido incluir e excluir linhas
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function
Private Function Inicializa_Grid_Distribuicao(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Distribuicao

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Centro de Custo")
    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("Unidade Medida")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Conta Contábil")

    objGridInt.colCampo.Add (Prod.Name)
    objGridInt.colCampo.Add (DescProd.Name)
    objGridInt.colCampo.Add (CentroCusto.Name)
    objGridInt.colCampo.Add (Almoxarifado.Name)
    objGridInt.colCampo.Add (UnidadeMed.Name)
    objGridInt.colCampo.Add (Quant.Name)
    objGridInt.colCampo.Add (ContaContabil.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_Prod_Col = 1
    iGrid_DescProduto_Col = 2
    iGrid_CentroCusto_Col = 3
    iGrid_Almoxarifado_Col = 4
    iGrid_UM_Col = 5
    iGrid_Quant_Col = 6
    iGrid_ContaContabil_Col = 7

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridDistribuicao

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_DISTRIBUICAO + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 20

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Distribuicao = SUCESSO

End Function

Private Function Inicializa_Grid_Bloqueios(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Distribuicao

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Tipo Bloqueio")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Usuário")
    objGridInt.colColuna.Add ("Responsável")
    objGridInt.colColuna.Add ("Data Liberação")
    objGridInt.colColuna.Add ("Resp. Liberação")

    ' campos de edição do grid
    objGridInt.colCampo.Add (TipoBloqueio.Name)
    objGridInt.colCampo.Add (DataBloqueio.Name)
    objGridInt.colCampo.Add (CodUsuario.Name)
    objGridInt.colCampo.Add (ResponsavelBL.Name)
    objGridInt.colCampo.Add (DataLiberacao.Name)
    objGridInt.colCampo.Add (ResponsavelLib.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_TipoBloqueio_Col = 1
    iGrid_DataBloqueio_Col = 2
    iGrid_CodUsuario_Col = 3
    iGrid_ResponsavelBL_Col = 4
    iGrid_DataLiberacao_Col = 5
    iGrid_ResponsavelLib_Col = 6

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridBloqueios

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_BLOQUEIOS + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 20

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Bloqueios = SUCESSO

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSiglas As New Collection
Dim objClasseUM As New ClassClasseUM
Dim objUM As ClassUnidadeDeMedida
Dim sUM As String, sProd As String
Dim iIndice As Integer

On Error GoTo Erro_Rotina_Grid_Enable
    
    'Pesquisa controle da coluna em questão
    Select Case objControl.Name

        Case PercentMaisReceb.Name, RecebForaFaixa.Name, Observ.Name

            'Verifica se o produto está preenchido
            lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 53188

            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        'se for Prod
        Case Prod.Name

            If Len(Trim(GridDistribuicao.TextMatrix(iLinha, iGrid_Prod_Col))) <> 0 Then
                sProd = GridDistribuicao.TextMatrix(iLinha, iGrid_Prod_Col)
            Else
                sProd = Prod.Text
            End If
           
            Prod.Clear
            
            'carrega os produtos que estao no grid de itens
            For iIndice = 1 To objGridItens.iLinhasExistentes
                'se o elemento está marcado para ser liberado
                If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Produto_Col))) > 0 Then
                    Prod.AddItem (GridItens.TextMatrix(iIndice, iGrid_Produto_Col))
                End If
            Next

            lErro = CF("Produto_Formata", GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_Prod_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then Error 49433

            For iIndice = 0 To Prod.ListCount - 1
                If sProd = Prod.List(iIndice) Then
                    Prod.ListIndex = iIndice
                    Exit For
                End If
            Next
            
            If iProdutoPreenchido = PRODUTO_VAZIO Or Len(Trim(GridDistribuicao.TextMatrix(iLinha, iGrid_Prod_Col))) = 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If

        'se o controle for CentroCusto ou Almoxarifado ou UnidMed ou Quant ou ContaContabil
        Case CentroCusto.Name, Almoxarifado.Name, Quant.Name, ContaContabil.Name
        
            'se o produto do grid de distribuicao nao estiver preenchido
            lErro = CF("Produto_Formata", GridDistribuicao.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 53190

            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            ElseIf iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                objControl.Enabled = True
            End If

        'se o controle for TipoBloqueio ou Responsavel
        Case TipoBloqueio.Name, ResponsavelBL.Name

            If iLinha <= iGridBloqueiosLinhasBloqueadas Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True

            End If
            
        Case Else
            objControl.Enabled = False

    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 53188, 53189, 53190, 68349, 68350, 68352

        Case 68351
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164422)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras

On Error GoTo Erro_Tela_Extrai

    sTabela = "PedCompraGerado_Fornecedor"
    
    lErro = Move_Tela_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then Error 53191
    
    'Preenche a coleção colCampoValor
    colCampoValor.Add "Codigo", objPedidoCompra.lCodigo, 0, "Codigo"
    colCampoValor.Add "OutrasDespesas", objPedidoCompra.dOutrasDespesas, 0, "OutrasDespesas"
    colCampoValor.Add "Data", objPedidoCompra.dtData, 0, "Data"
    colCampoValor.Add "DataAlteracao", objPedidoCompra.dtDataAlteracao, 0, "DataAlteracao"
    colCampoValor.Add "DataEnvio", objPedidoCompra.dtDataEnvio, 0, "DataEnvio"
    colCampoValor.Add "DataEmissao", objPedidoCompra.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "ValorDesconto", objPedidoCompra.dValorDesconto, 0, "ValorDesconto"
    colCampoValor.Add "ValorFrete", objPedidoCompra.dValorFrete, 0, "ValorFrete"
    colCampoValor.Add "ValorIPI", objPedidoCompra.dValorIPI, 0, "ValorIPI"
    colCampoValor.Add "ValorSeguro", objPedidoCompra.dValorSeguro, 0, "ValorSeguro"
    colCampoValor.Add "ValorTotal", objPedidoCompra.dValorTotal, 0, "ValorTotal"
    colCampoValor.Add "Comprador", objPedidoCompra.iComprador, 0, "Comprador"
    colCampoValor.Add "CondicaoPagto", objPedidoCompra.iCondicaoPagto, 0, "CondicaoPagto"
    colCampoValor.Add "Filial", objPedidoCompra.iFilial, 0, "Filial"
    colCampoValor.Add "FilialDestino", objPedidoCompra.iFilialDestino, 0, "FilialDestino"
    colCampoValor.Add "FilialEmpresa", objPedidoCompra.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "ProxSeqBloqueio", objPedidoCompra.iProxSeqBloqueio, 0, "ProxSeqBloqueio"
    colCampoValor.Add "TipoBaixa", objPedidoCompra.iTipoBaixa, 0, "TipoBaixa"
    colCampoValor.Add "TipoDestino", objPedidoCompra.iTipoDestino, 0, "TipoDestino"
    colCampoValor.Add "FornCliDestino", objPedidoCompra.lFornCliDestino, 0, "FornCliDestino"
    colCampoValor.Add "Fornecedor", objPedidoCompra.lFornecedor, 0, "Fornecedor"
    colCampoValor.Add "NumIntDoc", objPedidoCompra.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "Transportadora", objPedidoCompra.iTransportadora, 0, "Transportadora"
    colCampoValor.Add "Alcada", objPedidoCompra.sAlcada, STRING_BUFFER_MAX_TEXTO, "Alcada"
    colCampoValor.Add "Contato", objPedidoCompra.sContato, STRING_BUFFER_MAX_TEXTO, "Contato"
    colCampoValor.Add "MotivoBaixa", objPedidoCompra.sMotivoBaixa, STRING_BUFFER_MAX_TEXTO, "MotivoBaixa"
    colCampoValor.Add "Observacao", objPedidoCompra.lObservacao, 0, "Observacao"
    colCampoValor.Add "TipoFrete", objPedidoCompra.sTipoFrete, STRING_BUFFER_MAX_TEXTO, "TipoFrete"
'leo
    colCampoValor.Add "Taxa", objPedidoCompra.dTaxa, 0, "Taxa"
    colCampoValor.Add "Embalagem", objPedidoCompra.iEmbalagem, 0, "Embalagem"
    colCampoValor.Add "Moeda", objPedidoCompra.iMoeda, 0, "Moeda"
    colCampoValor.Add "ObsEmbalagem", objPedidoCompra.sObsEmbalagem, STRING_BUFFER_MAX_TEXTO, "ObsEmbalagem"
            
''    colCampoValor.Add "NomeReduzido", Fornecedor.Caption, STRING_BUFFER_MAX_TEXTO, "NomeReduzido"
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "DataEnvio", OP_IGUAL, DATA_NULA

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 53191

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164423)

    End Select

    Exit Sub


End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras

On Error GoTo Erro_Tela_Preenche

    'Carrega objPedidoCompra com os dados passados em colCampoValor
    objPedidoCompra.dOutrasDespesas = colCampoValor.Item("OutrasDespesas").vValor
    objPedidoCompra.dtData = colCampoValor.Item("Data").vValor
    objPedidoCompra.dtDataAlteracao = colCampoValor.Item("DataAlteracao").vValor
    objPedidoCompra.dtDataEmissao = colCampoValor.Item("DataEmissao").vValor
    objPedidoCompra.dtDataEnvio = colCampoValor.Item("DataEnvio").vValor
    objPedidoCompra.dValorDesconto = colCampoValor.Item("ValorDesconto").vValor
    objPedidoCompra.dValorFrete = colCampoValor.Item("ValorFrete").vValor
    objPedidoCompra.dValorIPI = colCampoValor.Item("ValorIPI").vValor
    objPedidoCompra.dValorSeguro = colCampoValor.Item("ValorSeguro").vValor
    objPedidoCompra.dValorTotal = colCampoValor.Item("ValorTotal").vValor
    objPedidoCompra.iComprador = colCampoValor.Item("Comprador").vValor
    objPedidoCompra.iCondicaoPagto = colCampoValor.Item("CondicaoPagto").vValor
    objPedidoCompra.iFilial = colCampoValor.Item("Filial").vValor
    objPedidoCompra.iFilialDestino = colCampoValor.Item("FilialDestino").vValor
    objPedidoCompra.iProxSeqBloqueio = colCampoValor.Item("ProxSeqBloqueio").vValor
    objPedidoCompra.iTipoBaixa = colCampoValor.Item("TipoBaixa").vValor
    objPedidoCompra.iTipoDestino = colCampoValor.Item("TipoDestino").vValor
    objPedidoCompra.lCodigo = colCampoValor.Item("Codigo").vValor
    objPedidoCompra.lFornCliDestino = colCampoValor.Item("FornCliDestino").vValor
    objPedidoCompra.lFornecedor = colCampoValor.Item("Fornecedor").vValor
    objPedidoCompra.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor
    objPedidoCompra.iTransportadora = colCampoValor.Item("Transportadora").vValor
    objPedidoCompra.sAlcada = colCampoValor.Item("Alcada").vValor
    objPedidoCompra.sContato = colCampoValor.Item("Contato").vValor
    objPedidoCompra.sMotivoBaixa = colCampoValor.Item("MotivoBaixa").vValor
    objPedidoCompra.lObservacao = colCampoValor.Item("Observacao").vValor
    objPedidoCompra.sTipoFrete = colCampoValor.Item("TipoFrete").vValor
    objPedidoCompra.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
'leo
    objPedidoCompra.iMoeda = colCampoValor.Item("Moeda").vValor
    objPedidoCompra.iEmbalagem = colCampoValor.Item("Embalagem").vValor
    objPedidoCompra.dTaxa = colCampoValor.Item("Taxa").vValor
    objPedidoCompra.sObsEmbalagem = colCampoValor.Item("ObsEmbalagem").vValor
    
    ' preenche a tela com os elementos do objPedidoCompra
    lErro = Traz_PedidoCompra_Tela(objPedidoCompra)
    If lErro <> SUCESSO Then Error 53192

    iAlterado = 0

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 53192

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164424)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Private Function Move_Tela_Memoria(objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objComprador As New ClassComprador
Dim objEmbalagem As New ClassEmbalagem

On Error GoTo Erro_Move_Tela_Memoria

    'guarda a FilialEmpresa e o Codigo em objPedidoCompra
    objPedidoCompra.lCodigo = StrParaLong(Codigo.Text)
    objPedidoCompra.iFilialEmpresa = giFilialEmpresa

    'Se o Fornecedor estiver preenchido lê o código no BD
    If Len(Trim(Fornecedor.Caption)) > 0 Then

        Set objFornecedor = New ClassFornecedor

        objFornecedor.sNomeReduzido = Fornecedor.Caption
        'Lê o fornecedor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then Error 53389
        If lErro = 6681 Then Error 53390 'Se nao encontrar --> erro

        objPedidoCompra.lFornecedor = objFornecedor.lCodigo

    End If

    'Recolhe os demais campos da Tela
    objPedidoCompra.iFilial = Codigo_Extrai(Filial.Caption)

    objPedidoCompra.iCondicaoPagto = Codigo_Extrai(CondPagto.Caption)

    objPedidoCompra.sContato = Contato.Text

    If Len(Trim(Comprador.Caption)) > 0 Then

        objComprador.sCodUsuario = gsUsuario
        lErro = CF("Comprador_Le_Usuario", objComprador)
        If lErro <> SUCESSO And lErro <> 50059 Then Error 53391
        If lErro = 50059 Then Error 53392

        objPedidoCompra.iComprador = objComprador.iCodigo

    End If

    'recolhe as datas
    objPedidoCompra.dtData = StrParaDate(Data.Caption)
    objPedidoCompra.dtDataAlteracao = StrParaDate(DataAlteracao.Caption)
    objPedidoCompra.dtDataEmissao = StrParaDate(DataEmissao.Caption)
    objPedidoCompra.dtDataEnvio = StrParaDate(DataEnvio.Text)
    objPedidoCompra.dtDataRefFluxo = StrParaDate(DataRefFluxo.Text)

    'recolhe a observacao
    objPedidoCompra.sObservacao = Observacao(0).Text
    objPedidoCompra.sObsEmbalagem = ObsEmbalagem.Text
    objPedidoCompra.dCotacaoMoeda = StrParaDbl(Taxa.Text)

    'Frame Valores
    objPedidoCompra.dValorProdutos = StrParaDbl(ValorProdutos.Caption)
    objPedidoCompra.dValorFrete = StrParaDbl(ValorFrete.Text)
    objPedidoCompra.dValorSeguro = StrParaDbl(ValorSeguro.Text)
    objPedidoCompra.dOutrasDespesas = StrParaDbl(OutrasDespesas.Text)
    objPedidoCompra.dValorDesconto = StrParaDbl(ValorDesc.Text)
    objPedidoCompra.dValorIPI = StrParaDbl(ValorIPI.Text)
    objPedidoCompra.dValorTotal = StrParaDbl(ValorTotal.Caption)

    'Frame Local de Entrega
    If TipoDestino(TIPO_DESTINO_EMPRESA) Then
        objPedidoCompra.iTipoDestino = TIPO_DESTINO_EMPRESA
        objPedidoCompra.iFilialDestino = Codigo_Extrai(FilialEmpresa.Text)
    End If
    If TipoDestino(TIPO_DESTINO_FORNECEDOR) Then
        objPedidoCompra.iTipoDestino = TIPO_DESTINO_FORNECEDOR

        'Preenche o Codigo da Filial Fornecedor
        objPedidoCompra.iFilialDestino = Codigo_Extrai(FilialFornec.Text)
        'codigo da filialfornecedor
        objFornecedor.sNomeReduzido = Fornec.Text
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then Error 53193
        If lErro = 6681 Then Error 53194

        'Preenche o Codigo do Fornecedor
        objPedidoCompra.lFornCliDestino = objFornecedor.lCodigo

    End If

    'Frame Frete
    objPedidoCompra.iTransportadora = Codigo_Extrai(Transportadora.Text)

'Leo
    objPedidoCompra.dTaxa = StrParaDbl(Taxa.Text)
    
    objPedidoCompra.iMoeda = Codigo_Extrai(Moeda.Caption)
    
    If Len(Trim(Embalagem.Text)) > 0 Then

        objEmbalagem.sSigla = Embalagem.Text

        lErro = CF("Embalagem_Le_Sigla", objEmbalagem)
        If lErro <> SUCESSO And lErro <> 95088 Then gError 103400
        
        If lErro <> SUCESSO Then gError 103401
        
        objPedidoCompra.iEmbalagem = objEmbalagem.iCodigo
        
    End If
    
'Leo até aqui

    If TipoFrete.ListIndex <> -1 Then
        objPedidoCompra.sTipoFrete = TipoFrete.ItemData(TipoFrete.ListIndex)
    End If

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case 53193, 53195, 53389, 53391, 103380

        Case 53194, 53390
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", Err, objFornecedor.lCodigo)

        Case 53392
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPRADOR_NAO_CADASTRADO", Err, objComprador.iCodigo)
        
        Case 103400
        
        Case 103401
            Call Rotina_Erro(vbOKOnly, "ERRO_EMBALAGEM_INEXISTENTE", gErr, objEmbalagem.sSigla)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164425)

    End Select

    Exit Function

End Function

Private Function Traz_PedidoCompra_Tela(objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long
Dim iItem As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim objObservacao As New ClassObservacao
Dim objFornecedor As New ClassFornecedor
Dim objTransportadora As New ClassTransportadora
Dim objEmbalagem As New ClassEmbalagem
Dim objMoeda As New ClassMoedas

On Error GoTo Erro_Traz_PedidoCompra_Tela

    lErro = CF("PedidoCompras_Le", objPedidoCompra)
    If lErro <> SUCESSO And lErro <> 56118 Then gError 188600

    ' lê os itens do Pedido de compra
    lErro = CF("ItensPC_Le", objPedidoCompra)
    If lErro <> SUCESSO Then gError 53197

    'Le os Bloqueios do Pedido de Compra
    lErro = CF("BloqueiosPC_Le", objPedidoCompra)
    If lErro <> SUCESSO Then gError 53198
'por leo
    lErro = CF("NotasPedCompras_Le", objPedidoCompra)
    If lErro <> SUCESSO Then gError 103355

    Call Limpa_Tela_PedidoCompras

'Leo
    If objPedidoCompra.dTaxa > 0 Then
        
        Taxa.PromptInclude = False
        Taxa.Text = Format(objPedidoCompra.dTaxa, FORMATO_TAXA_CONVERSAO_MOEDA)
        Taxa.PromptInclude = True
        
        dTaxaInformada = objPedidoCompra.dTaxa
        
    End If
    
    If objPedidoCompra.iEmbalagem > 0 Then
        
        objEmbalagem.iCodigo = objPedidoCompra.iEmbalagem
        
        lErro = CF("Embalagem_Le", objEmbalagem)
        If lErro <> SUCESSO And lErro <> 82763 Then gError 103402
        
        If lErro = SUCESSO Then
             
            Embalagem.Text = objEmbalagem.sSigla
            
        End If
            
    End If
        
    objMoeda.iCodigo = objPedidoCompra.iMoeda
    
    lErro = Moedas_Le(objMoeda)
    If lErro <> SUCESSO And lErro <> 108821 Then gError 103403
    
    If lErro = SUCESSO Then
        Moeda.Caption = objMoeda.iCodigo & SEPARADOR & objMoeda.sNome
    Else
        Moeda.Caption = ""
    End If
    
    'Se a moeda selecionada for = REAL
    If objMoeda.iCodigo = MOEDA_REAL Then
    
        'Limpa a cotacao
        Taxa.Text = ""
        
        'Desabilita a cotacao
        Taxa.Enabled = False
        BotaoTrazCotacao.Enabled = False
        
        bExibirColReal = False
        
    Else
            
        'Habilita a cotacao
        Taxa.Enabled = True
        BotaoTrazCotacao.Enabled = True
        
        bExibirColReal = True
    
    End If
    
    'Coloca os dados na tela
    Codigo.Text = objPedidoCompra.lCodigo
    Contato.Text = objPedidoCompra.sContato

    Data.Caption = Format(objPedidoCompra.dtData, "dd/mm/yyyy")

    objFornecedor.lCodigo = objPedidoCompra.lFornecedor

    'Lê o Fornecedor
    lErro = CF("Fornecedor_Le", objFornecedor)
    If lErro <> SUCESSO And lErro <> 12729 Then gError 53199
    If lErro = 12729 Then gError 53200

    'Coloca o NomeReduzido do Fornecedor na tela
    Fornecedor.Caption = objFornecedor.sNomeReduzido

    'Passa o CodFornecedor e o CodFilial para o objfilialfornecedor
    objFilialFornecedor.lCodFornecedor = objPedidoCompra.lFornecedor
    objFilialFornecedor.iCodFilial = objPedidoCompra.iFilial

    'Lê o filialforncedor
    lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", objFornecedor.sNomeReduzido, objFilialFornecedor)
    If lErro <> SUCESSO And lErro <> 18272 Then gError 53201
    'Se nao encontrou
    If lErro = 18272 Then gError 53202

    'Coloca a filial na tela
    Filial.Caption = objPedidoCompra.iFilial & SEPARADOR & objFilialFornecedor.sNome

   If objPedidoCompra.dtDataAlteracao <> DATA_NULA Then
        DataAlteracao.Caption = Format(objPedidoCompra.dtDataAlteracao, "dd/mm/yy")
    End If
    If objPedidoCompra.dtDataEmissao <> DATA_NULA Then
        DataEmissao.Caption = Format(objPedidoCompra.dtDataEmissao, "dd/mm/yy")
    End If

    'Preenche o TipoDestino
    TipoDestino(objPedidoCompra.iTipoDestino).Value = True

    If iFrameTipoDestinoAtual = TIPO_DESTINO_EMPRESA Then

        Call CF("Filial_Seleciona", FilialEmpresa, objPedidoCompra.iFilialDestino)

    ElseIf iFrameTipoDestinoAtual = TIPO_DESTINO_FORNECEDOR Then

        Fornec.Text = objPedidoCompra.lFornCliDestino
        Fornec_Validate (bSGECancelDummy)

        Call CF("Filial_Seleciona", FilialFornec, objPedidoCompra.iFilialDestino)

    End If

    For iItem = 0 To TipoFrete.ListCount - 1
        If objPedidoCompra.sTipoFrete = TipoFrete.ItemData(iItem) Then
            'coloca na tela
            TipoFrete.Text = TipoFrete.List(iItem)
        End If
    Next

    If objPedidoCompra.iTransportadora <> 0 Then

        objTransportadora.iCodigo = objPedidoCompra.iTransportadora

        'le a transportadora
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError 53217

        'se nao encontrou ---> gErro
        If lErro = 19250 Then gError 53218

        Transportadora.Text = objTransportadora.iCodigo & SEPARADOR & objTransportadora.sNomeReduzido

    End If

    If objPedidoCompra.iCondicaoPagto <> 0 Then

        objCondicaoPagto.iCodigo = objPedidoCompra.iCondicaoPagto

        'lê a cond. de pagto
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 53219

        'se nao encontrou --->erro
        If lErro = 19205 Then gError 53220

        CondPagto.Caption = objPedidoCompra.iCondicaoPagto & SEPARADOR & objCondicaoPagto.sDescReduzida

    End If

    If Len(Trim(DataEnvio.ClipText)) > 0 Then
        DataEnvio.Text = Format(objPedidoCompra.dtDataEnvio, "dd/mm/yyyy")
    End If

    If Len(Trim(DataRefFluxo.ClipText)) > 0 Then
        DataRefFluxo.Text = Format(objPedidoCompra.dtDataRefFluxo, "dd/mm/yyyy")
    End If

    'lê a observacao
    If objPedidoCompra.lObservacao > 0 Then

        objObservacao.lNumInt = objPedidoCompra.lObservacao

        lErro = CF("Observacao_Le", objObservacao)
        If lErro <> SUCESSO And lErro <> 53827 Then gError 53221
        If lErro = 53827 Then gError 53222

        Observacao(0).Text = objObservacao.sObservacao

    End If
    ObsEmbalagem.Text = objPedidoCompra.sObsEmbalagem

    'If objPedidoCompra.dValorProdutos > 0 Then ValorProdutos.Caption = Format(objPedidoCompra.dValorProdutos, "standard")
    If objPedidoCompra.dValorFrete > 0 Then ValorFrete.Text = Format(objPedidoCompra.dValorFrete, "standard")
    If objPedidoCompra.dValorSeguro > 0 Then ValorSeguro.Text = Format(objPedidoCompra.dValorSeguro, "standard")
    If objPedidoCompra.dOutrasDespesas > 0 Then OutrasDespesas.Text = Format(objPedidoCompra.dOutrasDespesas, "standard")
    If objPedidoCompra.dValorDesconto > 0 Then ValorDesc.Text = Format(objPedidoCompra.dValorDesconto, "standard")
    If objPedidoCompra.dValorIPI > 0 Then ValorIPI.Text = Format(objPedidoCompra.dValorIPI, "standard")

    'preenche o Grid com os itens do Pedido Compra
    lErro = Preenche_Grid_Itens(objPedidoCompra)
    If lErro <> SUCESSO Then gError 53223

    ' preenche o Grid de distribuicao atraves do objPedidoCompra
    lErro = Preenche_Grid_Distribuicao(objPedidoCompra)
    If lErro <> SUCESSO Then gError 53224

    'Preenche o GridBloqueio
    lErro = Preenche_Grid_Bloqueio(objPedidoCompra)
    If lErro <> SUCESSO Then gError 53225
'por leo
    'preenche o Grid com as Notas do Pedido Compra
    lErro = Preenche_Grid_Notas(objPedidoCompra)
    If lErro <> SUCESSO Then gError 103358

    'preenche o campo ValorTotal e ValorProdutos
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 53226

    iAlterado = 0

    Traz_PedidoCompra_Tela = SUCESSO

    Exit Function

Erro_Traz_PedidoCompra_Tela:

    Traz_PedidoCompra_Tela = gErr

    Select Case gErr

        Case 53197, 53198, 53199, 53201, 53217, 53219, 53221, 53223 To 53226, 103357, 103358, 103403, 188600

        Case 53200
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)

        Case 53202
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_FORNECEDOR_INEXISTENTE", gErr, objFilialFornecedor.lCodFornecedor, objFilialFornecedor.iCodFilial)

        Case 53220
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)

        Case 53222
            Call Rotina_Erro(vbOKOnly, "ERRO_OBSERVACAO_NAO_CADASTRADA", gErr, objPedidoCompra.lObservacao)

        Case 53218
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", gErr, objTransportadora.sNomeReduzido)

        Case 103402
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164426)

    End Select

    Exit Function

End Function

Private Function Preenche_Grid_Distribuicao(objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoEnxuto As String
Dim dPercDesc As Double
Dim iItem As Integer
Dim objItemPC As New ClassItemPedCompra
Dim objLocalizacao As New ClassLocalizacaoItemPC
Dim sCclMascarado As String
Dim sContaMascarada As String
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_Preenche_Grid_Distribuicao

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridDistribuicao)

    iIndice = 0
    iItem = 0

    For Each objItemPC In objPedidoCompra.colItens

        iItem = iItem + 1

        For Each objLocalizacao In objItemPC.colLocalizacao

            iIndice = iIndice + 1

            'Coloca os dados de distribuicao na tela
            GridDistribuicao.TextMatrix(iIndice, iGrid_Prod_Col) = GridItens.TextMatrix(iItem, iGrid_Produto_Col)
            GridDistribuicao.TextMatrix(iIndice, iGrid_Quant_Col) = Formata_Estoque(objLocalizacao.dQuantidade)
            GridDistribuicao.TextMatrix(iIndice, iGrid_DescProduto_Col) = GridItens.TextMatrix(iItem, iGrid_DescProduto_Col)
            GridDistribuicao.TextMatrix(iIndice, iGrid_UM_Col) = GridItens.TextMatrix(iItem, iGrid_UnidadeMed_Col)

            'Limpa a variável que será utilizada para obter o Ccl
            sCclMascarado = ""
            
            'Se o campo Ccl foi preenchido
            If Len(Trim(objLocalizacao.sCcl)) > 0 Then
                
                'Inicializa a string que receberá o Ccl
                sCclMascarado = String(STRING_CCL, 0)
                
                'Chama a função que irá formatar o Ccl
                lErro = Mascara_RetornaCclEnxuta(objLocalizacao.sCcl, sCclMascarado)
                If lErro <> SUCESSO Then Error 56401
                
                'Preenche o campo Ccl com o Ccl encontrado
                CentroCusto.PromptInclude = False
                CentroCusto.Text = sCclMascarado
                CentroCusto.PromptInclude = True

                'Joga o Ccl no grid
                GridDistribuicao.TextMatrix(iIndice, iGrid_CentroCusto_Col) = CentroCusto.Text

            End If

            If Len(Trim(objLocalizacao.sContaContabil)) <> 0 Then

                lErro = Mascara_RetornaContaEnxuta(objLocalizacao.sContaContabil, sContaMascarada)
                If lErro <> SUCESSO Then Error 56402

                ContaContabil.PromptInclude = False
                ContaContabil.Text = sContaMascarada
                ContaContabil.PromptInclude = True
                
                GridDistribuicao.TextMatrix(iIndice, iGrid_ContaContabil_Col) = ContaContabil.Text

            End If

            If objLocalizacao.iAlmoxarifado > 0 Then
                objAlmoxarifado.iCodigo = objLocalizacao.iAlmoxarifado

                lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25056 Then Error 51336
                If lErro <> SUCESSO Then Error 56403

                GridDistribuicao.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido
            End If
        Next

    Next

    objGridDistribuicao.iLinhasExistentes = iIndice

    Preenche_Grid_Distribuicao = SUCESSO

    Exit Function

Erro_Preenche_Grid_Distribuicao:

    Preenche_Grid_Distribuicao = Err

    Select Case Err

        Case 56401, 56402, 51336

        Case 56403
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", Err, objAlmoxarifado.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164427)

    End Select

    Exit Function

End Function

Private Function Preenche_Grid_Itens(objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoEnxuto As String
Dim dPercDesc As Double
Dim iItem As Integer
Dim objItemPC As New ClassItemPedCompra
Dim dPrecoTotal As Double, objProduto As New ClassProduto
Dim objObservacao As New ClassObservacao

On Error GoTo Erro_Preenche_Grid_Itens

    Set gcolItemPedido = New Collection

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridItens)

    iIndice = 0

    For Each objItemPC In objPedidoCompra.colItens

        iIndice = iIndice + 1

        lErro = Mascara_RetornaProdutoEnxuto(objItemPC.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then Error 53230

        'Mascara o produto enxuto
        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True

        'Calcula o percentual de desconto
        dPercDesc = objItemPC.dValorDesconto / (objItemPC.dPrecoUnitario + objItemPC.dQuantidade - objItemPC.dValorDesconto)
        dPrecoTotal = objItemPC.dPrecoUnitario * objItemPC.dQuantidade

        'Coloca os dados dos itens na tela
        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text

        GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col) = objItemPC.sDescProduto
        GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItemPC.sUM
        If objItemPC.dQuantRecebida > 0 Then GridItens.TextMatrix(iIndice, iGrid_QuantRecebida_Col) = Formata_Estoque(objItemPC.dQuantRecebida)
        If objItemPC.dQuantidade > 0 Then GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemPC.dQuantidade)
        If objItemPC.dPrecoUnitario > 0 Then GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col) = Format(objItemPC.dPrecoUnitario, gobjCOM.sFormatoPrecoUnitario) ' "STANDARD") 'Alterado por Wagner
        If dPercDesc > 0 Then GridItens.TextMatrix(iIndice, iGrid_PercentDesc_Col) = Format(dPercDesc, "Percent")
        If dPrecoTotal > 0 Then GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col) = Format(dPrecoTotal, PrecoTotal.Format) 'Alterado por Wagner
        
        If objItemPC.dtDataLimite <> DATA_NULA Then GridItens.TextMatrix(iIndice, iGrid_DataLimite_Col) = Format(objItemPC.dtDataLimite, "dd/mm/yy")
        If objItemPC.dPercentMaisReceb > 0 Then GridItens.TextMatrix(iIndice, iGrid_PercentMaisReceb_Col) = Format(objItemPC.dPercentMaisReceb, "Percent")
        If objItemPC.dAliquotaIPI > 0 Then GridItens.TextMatrix(iIndice, iGrid_AliquotaIPI_Col) = Format(objItemPC.dAliquotaIPI, "Percent")
        If objItemPC.dAliquotaICMS > 0 Then GridItens.TextMatrix(iIndice, iGrid_AliquotaICMS_Col) = Format(objItemPC.dAliquotaICMS, "Percent")

        'lê a observacao
        If objItemPC.lObservacao > 0 Then

            objObservacao.lNumInt = objItemPC.lObservacao

            lErro = CF("Observacao_Le", objObservacao)
            If lErro <> SUCESSO And lErro <> 53827 Then Error 53231
            If lErro = 53827 Then Error 53232

            GridItens.TextMatrix(iIndice, iGrid_Observacao_Col) = objObservacao.sObservacao

        End If

        For iItem = 0 To RecebForaFaixa.ListCount - 1
            If objItemPC.iRebebForaFaixa = RecebForaFaixa.ItemData(iItem) Then
                'coloca no Grid Itens RecebForaFaixa
                GridItens.TextMatrix(iIndice, iGrid_RecebForaFaixa_Col) = RecebForaFaixa.List(iItem)
            End If
        Next

        If objItemPC.dValorIPI > 0 Then GridItens.TextMatrix(iIndice, iGrid_ValorIPIItem_Col) = objItemPC.dValorIPI

        'Le o produto
        objProduto.sCodigo = objItemPC.sProduto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 56381
        'Se nao encontrou => erro
        If lErro = 28030 Then Error 56437
        
        'Preenche a descrição completa do produto com a ObsFisica do produto na tabela de produtos
        GridItens.TextMatrix(iIndice, iGrid_DescCompleta_Col) = objProduto.sObsFisica
        
        'Armazena os números internos dos itens
        gcolItemPedido.Add objItemPC.lNumIntDoc
    Next

    'Atualiza o número de linhas existentes
    objGridItens.iLinhasExistentes = gcolItemPedido.Count
    
    If Len(Trim(Taxa.Text)) > 0 Then Call ComparativoMoedaReal_Calcula(CDbl(Taxa.Text))

    Exit Function

Erro_Preenche_Grid_Itens:

    Preenche_Grid_Itens = Err

    Select Case Err

        Case 56437
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)
        
        Case 53230, 53231, 56381

        Case 53232
            Call Rotina_Erro(vbOKOnly, "ERRO_OBSERVACAO_NAO_CADASTRADA", Err, objPedidoCompra.lObservacao)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164428)

    End Select

    Exit Function

End Function

Private Function Preenche_Grid_Bloqueio(objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoEnxuto As String
Dim dPercDesc As Double
Dim iItem As Integer
Dim objBloqueioPC As New ClassBloqueioPC
Dim objTipoDeBloqueioPC As New ClassTipoBloqueioPC

On Error GoTo Erro_Preenche_Grid_Bloqueio

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridBloqueio)

    iIndice = 0

    For Each objBloqueioPC In objPedidoCompra.colBloqueiosPC

        iIndice = iIndice + 1
        GridBloqueios.TextMatrix(iIndice, iGrid_CodUsuario_Col) = objBloqueioPC.sCodUsuario
        GridBloqueios.TextMatrix(iIndice, iGrid_ResponsavelBL_Col) = objBloqueioPC.sResponsavel
        GridBloqueios.TextMatrix(iIndice, iGrid_ResponsavelLib_Col) = objBloqueioPC.sCodUsuarioLib
        If objBloqueioPC.dtDataLib <> DATA_NULA Then GridBloqueios.TextMatrix(iIndice, iGrid_DataLiberacao_Col) = Format(objBloqueioPC.dtDataLib, "dd/mm/yy")
        If (objBloqueioPC.dtData <> DATA_NULA) Then GridBloqueios.TextMatrix(iIndice, iGrid_DataBloqueio_Col) = Format(objBloqueioPC.dtData, "dd/mm/yy")

        objTipoDeBloqueioPC.iCodigo = objBloqueioPC.iTipoBloqueio

        lErro = CF("TipoDeBloqueioPC_Le", objTipoDeBloqueioPC)
        If lErro <> SUCESSO And lErro <> 49143 Then Error 57250
        If lErro = 49143 Then Error 57251

        GridBloqueios.TextMatrix(iIndice, iGrid_TipoBloqueio_Col) = objBloqueioPC.iTipoBloqueio & SEPARADOR & objTipoDeBloqueioPC.sNomeReduzido

    Next

    objGridBloqueio.iLinhasExistentes = objPedidoCompra.colBloqueiosPC.Count
    iGridBloqueiosLinhasBloqueadas = objGridBloqueio.iLinhasExistentes

    Preenche_Grid_Bloqueio = SUCESSO

    Exit Function

Erro_Preenche_Grid_Bloqueio:

    Preenche_Grid_Bloqueio = Err

    Select Case Err

        Case 57250

        Case 57251
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODEBLOQUEIOPC_NAO_CADASTRADO", Err, objTipoDeBloqueioPC.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164429)

    End Select

    Exit Function

End Function

Private Sub CodigoLabel_Click()

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras
Dim colSelecao As New Collection

On Error GoTo Erro_CodigoLabel_Click

    'Move os dados da tela
    lErro = Move_Tela_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then Error 53233

    'Chama a Tela de browse
    Call Chama_Tela("PedComprasGeradoLista", colSelecao, objPedidoCompra, objEventoCodigo)

    Exit Sub

Erro_CodigoLabel_Click:

    Select Case Err

        Case 53233

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164430)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objPedidoCompra = obj1

    If Not (objPedidoCompra Is Nothing) Then

        If objPedidoCompra.lNumIntDoc > 0 Then

            'Le o Pedido de Compra
            lErro = CF("PedidoCompras_Le", objPedidoCompra)
            If lErro <> SUCESSO And lErro <> 56118 Then Error 53242
            If lErro = 56118 Then Error 53243

            lErro = Traz_PedidoCompra_Tela(objPedidoCompra)
            If lErro <> SUCESSO Then Error 53244

        End If

    End If
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case Err

        Case 53242, 53244

        Case 53243
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", Err, objPedidoCompra.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164431)

    End Select

    Exit Sub

End Sub

Private Sub ObsLabel_Click()

Dim objObservacao As New ClassObservacao
Dim colSelecao As New Collection

    'Coloca em objObservacao a Observacao da tela
    objObservacao.sObservacao = Observacao(0).Text

    Call Chama_Tela("ObservacaoLista", colSelecao, objObservacao, objEventoObservacao)

    Exit Sub

End Sub

Private Sub objEventoObservacao_evSelecao(obj1 As Object)

Dim objObservacao As ClassObservacao

    Set objObservacao = obj1

    'Coloca observacao na tela
    Observacao(0).Text = objObservacao.sObservacao

    Me.Show

    Exit Sub

End Sub

Private Sub BotaoPedidosGerados_Click()

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoPedidosGerados_Click

    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then Error 53235

    'Chama a tela PedComprasNaoEnvLista
    Call Chama_Tela("PedComprasGeradoLista", colSelecao, objPedidoCompra, objEventoBotaoPedidosGerados)

    Exit Sub

Erro_BotaoPedidosGerados_Click:

    Select Case Err

        Case 53235

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164432)

    End Select

    Exit Sub

End Sub

Private Sub objEventoBotaoPedidosGerados_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras

On Error GoTo Erro_objEventoBotaoPedidosGerados_evSelecao

    Set objPedidoCompra = obj1

    lErro = Traz_PedidoCompra_Tela(objPedidoCompra)
    If lErro <> SUCESSO Then Error 53236

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoBotaoPedidosGerados_evSelecao:

    Select Case Err

        Case 53236

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164433)

    End Select

    Exit Sub

End Sub


Private Sub FornLabel_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As New Collection

    objFornecedor.sNomeReduzido = Fornec.Text

    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornec)

End Sub

Private Sub objEventoFornec_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor
Dim bCancel As Boolean

    Set objFornecedor = obj1

    Fornec.Text = objFornecedor.sNomeReduzido

    'Dispara Validate de Fornec
    Fornec_Validate (bCancel)

    Me.Show

End Sub

Private Sub TransportadoraLabel_Click()

Dim objTransportadora As New ClassTransportadora
Dim colSelecao As New Collection

    'Verifica se transportadora esta preenchida
    If Len(Trim(Transportadora.Text)) > 0 Then

        objTransportadora.iCodigo = Codigo_Extrai(Transportadora.Text)

    End If

    Call Chama_Tela("TransportadoraLista", colSelecao, objTransportadora, objEventoTransportadora)

    Exit Sub

End Sub

Private Sub objEventoTransportadora_evSelecao(obj1 As Object)

Dim objTransportadora As ClassTransportadora

    Set objTransportadora = obj1

    'Coloca a Transportadora na Tela
    Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNomeReduzido

    Me.Show

    Exit Sub

End Sub

Private Sub BotaoAlmoxarifado_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iFilialEmpresa As Integer
Dim iProdutoPreenchido As Integer
Dim iTipoDestino As Integer
Dim colSelecao As New Collection
Dim objEstoqueProdutoTerceiros As New ClassEstoqueProduto

On Error GoTo Erro_BotaoAlmoxarifado_Click

    'Verifica se existe alguma linha selecionada
    If GridDistribuicao.Row = 0 Then gError 89441

    ' se o tipo destino for empresa sai da rotina
    If iTipoDestino <> TIPO_DESTINO_EMPRESA Then Exit Sub

    'Verifica se o produto da linha esta preenchido
    If Len(Trim(GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_Prod_Col))) > 0 Then

        'Passa o produto para o formato do BD
        lErro = CF("Produto_Formata", GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 53134

        iFilialEmpresa = Codigo_Extrai(FilialEmpresa.Text)
        If Len(Trim(iFilialEmpresa)) = 0 Then gError 84610

        'adiciona produto e FiliaEmpresa de Destino em colSelecao
        colSelecao.Add sProdutoFormatado
        colSelecao.Add iFilialEmpresa
        
        Call Chama_Tela("AlmoxarifadoFilialLista", colSelecao, objEstoqueProdutoTerceiros, objEventoAlmoxarifados)

    End If

    Exit Sub

Erro_BotaoAlmoxarifado_Click:

    Select Case gErr

        Case 53134

        Case 84610
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_ENTREGA_NAO_PREENCHIDA", gErr)

        Case 89441
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164434)

    End Select

    Exit Sub

End Sub

Private Sub objEventoAlmoxarifados_evSelecao(obj1 As Object)

Dim objEstoqueProduto As ClassEstoqueProduto

    Set objEstoqueProduto = obj1

    If GridDistribuicao.Row = 0 Then Exit Sub

    GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_Almoxarifado_Col) = objEstoqueProduto.sAlmoxarifadoNomeReduzido
    Almoxarifado.Text = objEstoqueProduto.sAlmoxarifadoNomeReduzido

    Me.Show

    Exit Sub

End Sub

Private Sub BotaoCcl_Click()

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclFormata As String
Dim iCclPreenchida As Integer
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoCcl_Click

    'Se alguma linha do Grid Distribuicao esta selecionada
    If (GridDistribuicao.Row <> 0) Then

        'Verifica se o Ccl esta preenchido
        If Len(Trim(GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_CentroCusto_Col))) > 0 Then

            'Formata o Ccl
            lErro = CF("Ccl_Formata", GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_CentroCusto_Col), sCclFormata, iCclPreenchida)
            If lErro <> SUCESSO Then Error 53238

            'Coloca Ccl formatada em objCcl.sCcl
            objCcl.sCcl = sCclFormata

        End If

        'Chama tela CclLista
        Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoBotaoCcl)

    End If

    Exit Sub

Erro_BotaoCcl_Click:

    Select Case Err

        Case 53238

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164435)

    End Select

    Exit Sub

End Sub

Private Sub objEventoBotaoCcl_evSelecao(obj1 As Object)

Dim lErro As Long
Dim sCclMascarado As String
Dim objCcl As New ClassCcl
Dim iIndice As Integer

On Error GoTo Erro_objEventoBotaoCcl_evSelecao

    Set objCcl = obj1

    'Mascara o Ccl
    lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then Error 53239

    'Coloca ccl mascarado no GridDistribuicao
    CentroCusto.PromptInclude = False
    CentroCusto.Text = sCclMascarado
    CentroCusto.PromptInclude = True
    
    Me.Show

    Exit Sub

Erro_objEventoBotaoCcl_evSelecao:

    Select Case Err

        Case 53239

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164436)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_MascaraCcl() As Long
'Inicializa a mascara do centro de custo

Dim sMascaraCcl As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_mascaraccl

    sMascaraCcl = String(STRING_CCL, 0)

    'le a mascara dos centros de custo/lucro
    lErro = MascaraCcl(sMascaraCcl)
    If lErro <> SUCESSO Then Error 53387

    CentroCusto.Mask = sMascaraCcl

    Inicializa_MascaraCcl = SUCESSO

    Exit Function

Erro_Inicializa_mascaraccl:

    Inicializa_MascaraCcl = Err

    Select Case Err

        Case 53387

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164437)

    End Select

    Exit Function

End Function

Private Sub BotaoContasContabeis_Click()

Dim lErro As Long
Dim objPlanoConta As New ClassPlanoConta
Dim sContaFormata As String
Dim iContaPreenchida As Integer
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoContasContabeis_Click

    'Verifica s alguma linha do grid foi selecionada
    If (GridDistribuicao.Row <> 0) Then

        'Verifica se ContaContabil foi preenchida
        If Len(Trim(GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_ContaContabil_Col))) > 0 Then

            'Formata a conta contabil
            lErro = CF("Conta_Formata", GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_ContaContabil_Col), sContaFormata, iContaPreenchida)
            If lErro <> SUCESSO Then Error 53240

            'Coloca contaFormatada em objPlanoConta
            objPlanoConta.sConta = sContaFormata

        End If

        'Chama tela PlanoContaEstLista
        Call Chama_Tela("PlanoContaEstLista", colSelecao, objPlanoConta, objEventoContaContabil)

    End If

    Exit Sub

Erro_BotaoContasContabeis_Click:

    Select Case Err

        Case 53240

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164438)

    End Select

    Exit Sub

End Sub

Private Sub objEventoContaContabil_evSelecao(obj1 As Object)

Dim lErro As Long
Dim sContaMascarado As String
Dim objPlanoConta As New ClassPlanoConta

On Error GoTo Erro_objEventoContaContabil_evSelecao

    Set objPlanoConta = obj1

    'Mascara a conta contabil
    lErro = Mascara_MascararConta(objPlanoConta.sConta, sContaMascarado)
    If lErro <> SUCESSO Then Error 53241

    'Coloca no Grid a Conta mascarada
    ContaContabil.PromptInclude = False
    ContaContabil.Text = sContaMascarado
    ContaContabil.PromptInclude = True
    
    'Coloca no Grid a Conta mascarada
    GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_ContaContabil_Col) = ContaContabil.Text
    
    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case Err

        Case 53241

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164439)

    End Select

    Exit Sub

End Sub

Public Function Trata_Parametros(Optional objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se algum pedido foi passado por parametro
    If Not (objPedidoCompra Is Nothing) Then

        If objPedidoCompra.lNumIntDoc > 0 Then

            'Le o Pedido de Compra
            lErro = CF("PedidoCompras_Le", objPedidoCompra)
            If lErro <> SUCESSO And lErro <> 56118 Then Error 53242
            If lErro = 56118 Then Error 53243

            lErro = Traz_PedidoCompra_Tela(objPedidoCompra)
            If lErro <> SUCESSO Then Error 53244

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 53242, 53244

        Case 53243
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", Err, objPedidoCompra.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164440)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 53386

    'Limpa a tela
    Call Limpa_Tela_PedidoCompras

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 53386

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164441)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Form_Activate()

   Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

    'Libera as variaveis globais
    Set objEventoAlmoxarifados = Nothing
    Set objEventoBotaoPedidosGerados = Nothing
    Set objEventoCliente = Nothing
    Set objEventoCodigo = Nothing
    Set objEventoContaContabil = Nothing
    Set objEventoObservacao = Nothing
    Set objEventoTransportadora = Nothing
    Set objEventoBotaoCcl = Nothing
    Set objEventoFornec = Nothing
    Set objEventoNotaPC = Nothing
    
    Set objGridItens = Nothing
    Set objGridBloqueio = Nothing
    Set objGridNotas = Nothing
    Set objGridDistribuicao = Nothing
    
    Set objEventoEmbalagens = Nothing 'leo
    
    Set gcolItemPedido = Nothing
        
    '??? por leo
    Set objGridNotas = Nothing
        
    'Libera o comando das setas se estiver aberto
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

End Sub

Sub Limpa_Tela_PedidoCompras()

    'Limpa a tela
    Call Limpa_Tela(Me)

    'Limpa os outros campos da tela
    Codigo.Text = ""
    Fornecedor.Caption = ""
    DataAlteracao.Caption = ""
    DataEmissao.Caption = ""
    Contato.Text = ""
    Data.Caption = ""
    Filial.Caption = ""
    CondPagto.Caption = ""
    ValorTotal.Caption = ""
    ValorProdutos.Caption = ""
    Moeda.Caption = ""

    'Limpa  combobox
    TipoFrete.ListIndex = -1
    Transportadora.ListIndex = -1
    iGridBloqueiosLinhasBloqueadas = 0

    'Limpa os grids
    Call Grid_Limpa(objGridItens)
    Call Grid_Limpa(objGridBloqueio)
    Call Grid_Limpa(objGridDistribuicao)
    Call Grid_Limpa(objGridNotas) '??? leo
    
    ImprimePedido.Value = vbUnchecked 'Inserido por Wagner
    
    Set gcolItemPedido = New Collection

    TipoDestino(TIPO_DESTINO_EMPRESA).Value = True
    
    dTaxaInformada = 0 'leo
    
    Exit Sub

End Sub

Private Sub UM_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UM_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UM_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub UM_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub UM_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = UnidadeMed
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UpDownDataEnvio_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvio_DownClick

    'Diminui um dia em DataEnvio
    lErro = Data_Up_Down_Click(DataEnvio, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 53247

    Exit Sub

Erro_UpDownDataEnvio_DownClick:

    Select Case Err

        Case 53247

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164442)

    End Select

    Exit Sub


End Sub

Private Sub UpDownDataEnvio_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvio_UpClick

    'Aumenta um dia em DataEnvio
    lErro = Data_Up_Down_Click(DataEnvio, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 53248

    Exit Sub

Erro_UpDownDataEnvio_UpClick:

    Select Case Err

        Case 53248

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164443)

    End Select

    Exit Sub

End Sub

Function ValorTotal_Calcula() As Long

Dim dPrecoTotal As Double
Dim dValorTotal As Double
Dim iIndice As Integer

On Error GoTo Erro_ValorTotal_Calcula

   For iIndice = 1 To objGridItens.iLinhasExistentes

        'Calcula a soma dos valores de produtos
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))) > 0 Then

            If StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)) > 0 Then
                dValorTotal = dValorTotal + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
            End If

        End If
        'Calcula Preco Total das linhas do GridItens
        dPrecoTotal = dPrecoTotal + (StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)))

    Next

    'Coloca na tela o valor dos produtos
    ValorProdutos.Caption = Format(dPrecoTotal, PrecoTotal.Format) 'Alterado por Wagner
    dValorTotal = (dPrecoTotal + StrParaDbl(ValorFrete.Text) + StrParaDbl(ValorSeguro.Text) + StrParaDbl(OutrasDespesas.Text) + StrParaDbl(ValorIPI.Text) - StrParaDbl(ValorDesc.Text))

    'Coloca na tela o valor total
    ValorTotal.Caption = Format(dValorTotal, PrecoTotal.Format) 'Alterado por Wagner

    ValorTotal_Calcula = SUCESSO

    Exit Function

Erro_ValorTotal_Calcula:

    ValorTotal_Calcula = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164444)

    End Select

    Exit Function

End Function
'aquii
Private Sub ValorDesc_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dValorDesconto As Double

On Error GoTo Erro_ValorDesconto_Validate

    'Verifica se o ValorDesconto foi preenchido
    If Len(Trim(ValorDesc.Text)) > 0 Then

        'Critica o valor informado
        lErro = Valor_NaoNegativo_Critica(ValorDesc.Text)
        If lErro <> SUCESSO Then Error 53255

        dValorDesconto = StrParaDbl(ValorDesc.Text)

        'Coloca o valor no formato standard da tela
        ValorDesc.Text = Format(dValorDesconto, "Standard")

    End If

    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then Error 53256

    Exit Sub

Erro_ValorDesconto_Validate:

    Cancel = True

    Select Case Err

        Case 53255, 53256

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164445)

    End Select

    Exit Sub

End Sub

Private Sub ValorFrete_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dValorFrete As Double

On Error GoTo Erro_ValorFrete_Validate

    'Verifica se ValorFrete foi preenchido
    If Len(Trim(ValorFrete.Text)) > 0 Then

        'Critica o valor informado
        lErro = Valor_NaoNegativo_Critica(ValorFrete.Text)
        If lErro <> SUCESSO Then Error 53249

        dValorFrete = StrParaDbl(ValorFrete.Text)

        'Coloca o valor no formato standard da tela
        ValorFrete.Text = Format(dValorFrete, "Standard")

    End If

    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then Error 53250

    Exit Sub

Erro_ValorFrete_Validate:

    Cancel = True

    Select Case Err

        Case 53249, 53250

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164446)

    End Select

    Exit Sub

End Sub

Private Sub ValorIPI_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dValorIPI As Double

On Error GoTo Erro_ValorIPI_Validate

    'Verifica se o ValorIPI foi preenchido
    If Len(Trim(ValorIPI.Text)) > 0 Then

        'Critica o valor informado
        lErro = Valor_NaoNegativo_Critica(ValorIPI.Text)
        If lErro <> SUCESSO Then Error 53257

        dValorIPI = StrParaDbl(ValorIPI.Text)

        'Coloca o valor no formato standard da tela
        ValorIPI.Text = Format(dValorIPI, "Standard")

    End If

    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then Error 53258

    Exit Sub

Erro_ValorIPI_Validate:

    Cancel = True

    Select Case Err

        Case 53257, 53258

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164447)

    End Select

    Exit Sub

End Sub

Private Sub TipoDestino_Click(Index As Integer)

Dim lErro As Long

On Error GoTo Erro_TipoDestino_Click

    If Index = iFrameTipoDestinoAtual Then Exit Sub

    'Torna Frame correspondente a Index visivel
    FrameTipo(Index).Visible = True

    'Torna Frame atual invisivel
    FrameTipo(iFrameTipoDestinoAtual).Visible = False

    'Armazena novo valor de iFrameTipoDestinoAtual
    iFrameTipoDestinoAtual = Index

    'Limpa o frame de Endereco
    Call Limpa_Frame_Endereco

    If Index <> TIPO_DESTINO_EMPRESA Then

        'Limpa o Grid de Distribuicao
        Call Grid_Limpa(objGridDistribuicao)
        Frame1(4).Enabled = False

    Else

        Frame1(4).Enabled = True
        Call CF("Filial_Seleciona", FilialEmpresa, giFilialEmpresa)
        Call FilialEmpresa_Click

    End If

    Exit Sub

Erro_TipoDestino_Click:

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164448)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Frame_Endereco()

    Endereco.Caption = ""
    Bairro.Caption = ""
    Cidade.Caption = ""
    CEP.Caption = ""
    Estado.Caption = ""
    Pais.Caption = ""

    Exit Sub

End Sub

Private Sub FilialEmpresa_Click()

Dim objFilialEmpresa As New AdmFiliais
Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_FilialEmpresa_Click

    'Verifica se a filial foi selecionada
    If FilialEmpresa.ListIndex = -1 Then Exit Sub

    'Verifica o tipo destino
    If TipoDestino(TIPO_DESTINO_EMPRESA).Value = True Then

        'Limpa o grid de Distribuicao
        Call Grid_Limpa(objGridDistribuicao)

        'preeenche Codigo da Filial em objFilialEmpresa
        objFilialEmpresa.iCodFilial = Codigo_Extrai(FilialEmpresa.Text)

        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then Error 53260

        'Se nao encontrou => erro
        If lErro = 27378 Then Error 53261

        'Traz endereco para tela
        If objFilialEmpresa.objEnderecoEntrega.lCodigo <> 0 Then
            Call Preenche_Endereco(objFilialEmpresa.objEnderecoEntrega)
        Else
            Call Preenche_Endereco(objFilialEmpresa.objEndereco)
        End If

    End If

    'Verifica se o Grid Itens esta preenchido
    If objGridItens.iLinhasExistentes > 0 Then

        Call Distribuicao_Grid_Atualiza

    End If

    Exit Sub

Erro_FilialEmpresa_Click:

    Select Case Err

        Case 53260

        Case 53261
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err, objFilialEmpresa.iCodFilial)
            FilialEmpresa.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164449)

    End Select

    Exit Sub

End Sub

Private Sub Preenche_Endereco(objEndereco As ClassEndereco)

Dim objPais As New ClassPais
Dim lErro As Long

On Error GoTo Erro_Preenche_Endereco

    objPais.iCodigo = objEndereco.iCodigoPais

    lErro = CF("Paises_Le", objPais)
    If lErro <> SUCESSO And lErro <> 47876 Then Error 53259
    If lErro = 47876 Then Error 53269

    Endereco.Caption = objEndereco.sEndereco
    Bairro.Caption = objEndereco.sBairro
    Estado.Caption = objEndereco.sSiglaEstado
    Cidade.Caption = objEndereco.sCidade
    Pais.Caption = objPais.sNome
    CEP.Caption = objEndereco.sCEP

    Exit Sub

Erro_Preenche_Endereco:

    Select Case Err

        Case 53259

        Case 53269
            Call Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_CADASTRADO", Err, objEndereco.iCodigoPais)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164450)

    End Select

    Exit Sub

End Sub

Function Distribuicao_Grid_Atualiza() As Long

Dim lErro As Long
Dim sProduto As String
Dim iIndice As Integer
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_Distribuicao_Grid_Atualiza

    For iIndice = 1 To objGridItens.iLinhasExistentes

        'Formata o produto
        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then Error 53270

        objProduto.sCodigo = sProduto

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 53271

        'Se nao encontrou
        If lErro = 28030 Then Error 53272

        If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then

            'Atualiza a Distribuicao
            lErro = Distribuicao_Atualiza(iIndice, objProduto)
            If lErro <> SUCESSO Then Error 53273

        End If

    Next

    Distribuicao_Grid_Atualiza = SUCESSO

    Exit Function

Erro_Distribuicao_Grid_Atualiza:

    Distribuicao_Grid_Atualiza = Err

    Select Case Err

        Case 53270, 53271, 53273

        Case 53272
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164451)

    End Select

    Exit Function

End Function

Function Distribuicao_Atualiza(iItem As Integer, objProduto As ClassProduto) As Long

Dim lErro As Long
Dim dQuantidade As Double
Dim iFilialEmpresa As Integer
Dim sProduto As String
Dim iAlmoxarifadoPadrao As Integer
Dim sConta As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim iIndice As Integer, iIndice2 As Integer
Dim sContaMascarada As String
Dim objLocItemPC As ClassLocalizacaoItemPC
Dim objItemPC As ClassItemPedCompra
Dim colItem As New Collection

On Error GoTo Erro_Distribuicao_Atualiza

    'Guarda a quantidade do grid itens
    dQuantidade = StrParaDbl(GridItens.TextMatrix(iItem, iGrid_Quantidade_Col))
   
    For iIndice = 1 To objGridItens.iLinhasExistentes
        
        Set objItemPC = New ClassItemPedCompra
        colItem.Add objItemPC
        
        For iIndice2 = 1 To objGridDistribuicao.iLinhasExistentes
            If (GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = GridDistribuicao.TextMatrix(iIndice2, iGrid_Prod_Col)) And (iIndice <> iItem) Then
                Set objLocItemPC = New ClassLocalizacaoItemPC
                objItemPC.sProduto = GridItens.TextMatrix(iIndice, iGrid_Produto_Col)
                objItemPC.sDescProduto = GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col)
                objItemPC.sObservacao = GridDistribuicao.TextMatrix(iIndice2, iGrid_Almoxarifado_Col)
                objItemPC.sUM = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
                objLocItemPC.dQuantidade = StrParaDbl(GridDistribuicao.TextMatrix(iIndice2, iGrid_Quant_Col))
                objLocItemPC.sCcl = GridDistribuicao.TextMatrix(iIndice2, iGrid_CentroCusto_Col)
                objLocItemPC.sContaContabil = GridDistribuicao.TextMatrix(iIndice2, iGrid_ContaContabil_Col)
                objItemPC.colLocalizacao.Add objLocItemPC
            End If
        Next

    Next

    Set objItemPC = colItem(iItem)

    Call Grid_Limpa(objGridDistribuicao)

    iFilialEmpresa = Codigo_Extrai(FilialEmpresa.Text)

    sProduto = objProduto.sCodigo
    'Busca em FilialProduto o Almoxarifado padrao do produto
    lErro = CF("AlmoxarifadoPadrao_Le", iFilialEmpresa, sProduto, iAlmoxarifadoPadrao)
    If lErro <> SUCESSO And lErro <> 23796 Then gError 56185

    If iAlmoxarifadoPadrao > 0 Then

        objAlmoxarifado.iCodigo = iAlmoxarifadoPadrao

        'Busca a conta contabil na tabela de Almoxarifado
        lErro = CF("Almoxarifado_Le", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25056 Then gError 56342
        If lErro = 25056 Then gError 56436

        sConta = objAlmoxarifado.sContaContabil

    End If

    'Se nao encontrou a conta contabil
    If Len(Trim(objAlmoxarifado.sContaContabil)) = 0 Then

        'Busca a conta contabil do Produto
        sConta = objProduto.sContaContabil

    End If

    'Verifica se a conta contabil esta preenchida
    If Len(Trim(sConta)) > 0 Then

        lErro = Mascara_RetornaContaEnxuta(sConta, sContaMascarada)
        If lErro <> SUCESSO Then gError 79766

        ContaContabil.PromptInclude = False
        ContaContabil.Text = sContaMascarada
        ContaContabil.PromptInclude = True

    End If

    'cria uma linha no grid Distribuicao
    objItemPC.sProduto = GridItens.TextMatrix(iItem, iGrid_Produto_Col)
    objItemPC.sDescProduto = GridItens.TextMatrix(iItem, iGrid_DescProduto_Col)
    objItemPC.sUM = GridItens.TextMatrix(iItem, iGrid_UnidadeMed_Col)
    objItemPC.sObservacao = objAlmoxarifado.sNomeReduzido
    Set objLocItemPC = New ClassLocalizacaoItemPC
    objLocItemPC.dQuantidade = dQuantidade
    objLocItemPC.sContaContabil = ContaContabil.Text
    objItemPC.colLocalizacao.Add objLocItemPC

    For Each objItemPC In colItem
        For Each objLocItemPC In objItemPC.colLocalizacao
            objGridDistribuicao.iLinhasExistentes = objGridDistribuicao.iLinhasExistentes + 1
            'Preenche linha do grid distribuicao
            GridDistribuicao.TextMatrix(objGridDistribuicao.iLinhasExistentes, iGrid_Almoxarifado_Col) = objItemPC.sObservacao
            GridDistribuicao.TextMatrix(objGridDistribuicao.iLinhasExistentes, iGrid_Prod_Col) = objItemPC.sProduto
            GridDistribuicao.TextMatrix(objGridDistribuicao.iLinhasExistentes, iGrid_DescProduto_Col) = objItemPC.sDescProduto
            GridDistribuicao.TextMatrix(objGridDistribuicao.iLinhasExistentes, iGrid_UM_Col) = objItemPC.sUM
            GridDistribuicao.TextMatrix(objGridDistribuicao.iLinhasExistentes, iGrid_Quant_Col) = Formata_Estoque(objLocItemPC.dQuantidade)
            GridDistribuicao.TextMatrix(objGridDistribuicao.iLinhasExistentes, iGrid_ContaContabil_Col) = objLocItemPC.sContaContabil
        Next
    Next

    Distribuicao_Atualiza = SUCESSO

    Exit Function

Erro_Distribuicao_Atualiza:

    Distribuicao_Atualiza = gErr

    Select Case gErr

        Case 56185, 79766, 56319, 56336 To 56345

        Case 56436
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", gErr, objAlmoxarifado.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164452)

    End Select

    Exit Function

End Function

Private Sub Fornec_Change()

    iAlterado = REGISTRO_ALTERADO
    iFornecAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorSeguro_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dValorSeguro As Double

On Error GoTo Erro_ValorSeguro_Validate

    'Verifica se o ValorSeguro esta preenchido
    If Len(Trim(ValorSeguro.Text)) > 0 Then

        'Critica o valor informado
        lErro = Valor_NaoNegativo_Critica(ValorSeguro.Text)
        If lErro <> SUCESSO Then Error 53251

        dValorSeguro = StrParaDbl(ValorSeguro.Text)

        'Coloca o valor no formato standard da tela
        ValorSeguro.Text = Format(dValorSeguro, "Standard")

    End If

    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then Error 53252

    Exit Sub

Erro_ValorSeguro_Validate:

    Cancel = True

    Select Case Err

        Case 53251, 53252

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164453)

    End Select

    Exit Sub

End Sub

Private Sub FilialFornec_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim vbMsgRes As VbMsgBoxResult
Dim objEndereco As New ClassEndereco

On Error GoTo Erro_FilialFornec_Validate

    'Verifica se FilialFornec esta preenchida
    If Len(Trim(FilialFornec.Text)) > 0 Then

        'Verifica se FilialFornec esta selecionada
        If FilialFornec.ListIndex <> -1 Then Exit Sub

        'Seleciona combo box de FilialFornec
        lErro = Combo_Seleciona(FilialFornec, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 53287

        'Se nao encontra o ítem com o código informado
        If lErro = 6730 Then

            'Verifica de o fornecedor foi digitado
            If Len(Trim(Fornec.ClipText)) = 0 Then Error 53288

            objFilialFornecedor.iCodFilial = iCodigo

            'Pesquisa se existe filial com o codigo extraido
            lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornec.Text, objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 18272 Then Error 53289

            If lErro = 18272 Then Error 53290

            'Coloca na tela Codigo e Nome Reduzido de FilialFornec
            FilialFornec.Text = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome

            objEndereco.lCodigo = objFilialFornecedor.lEndereco

            'Le o endereco da FilialFornecedor passada como parametro
            lErro = CF("Endereco_Le", objEndereco)
            If lErro <> SUCESSO And lErro <> 12309 Then Error 53291
            'Se nao encontrou o endereco
            If lErro = 12309 Then Error 53292

            Call Preenche_Endereco(objEndereco)

        End If

        'Não encontrou valor informado que era STRING
        If lErro = 6731 Then Error 53293

    Else

        'Limpa Frame Endereco
        Call Limpa_Frame_Endereco

    End If

    Exit Sub

Erro_FilialFornec_Validate:

    Cancel = True

    Select Case Err

        Case 53287, 53289, 53291

        Case 53288
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 53290
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, FilialFornec.Text)
            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 53293
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA", Err, objFilialFornecedor.sNome)

        Case 53292
            Call Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_NAO_CADASTRADO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164454)

    End Select

    Exit Sub

End Sub

Private Sub TipoFrete_Click()

Dim lErro As Long

    'Se nao Tiver nenhum TipoFrete selecionado
    If TipoFrete.ListIndex = -1 Then

        'Limpa e desabilita combo de Transportadora
        Transportadora.Text = ""
        Transportadora.Enabled = False
        Exit Sub

    Else    'Se tiver um TipoFrete selecionado

        'Se for TIPO_FOB
        If TipoFrete.ItemData(TipoFrete.ListIndex) = TIPO_FOB Then

            'Habilita a combo de Transportadora
            Transportadora.Enabled = True
            TransportadoraLabel.Enabled = True
            

        Else    'Se nao for TIPO_FOB

            'Limpa  e desabilita a combo Transportadora
            Transportadora.Text = ""
            Transportadora.Enabled = False
            TransportadoraLabel.Enabled = False

        End If

    End If

    Exit Sub

End Sub

Private Sub Transportadora_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTransportadora As New ClassTransportadora
Dim iCodigo As Integer

On Error GoTo Erro_Transportadora_Validate

    If Len(Trim(Transportadora.Text)) > 0 Then

        If Transportadora.ListIndex <> -1 Then Exit Sub

        'Verifica se existe o ítem na List da Combo. Se existir seleciona.
        lErro = Combo_Seleciona(Transportadora, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 53301

        'Nao existe o ítem com o CÓDIGO na List da ComboBox
        If lErro = 6730 Then

            objTransportadora.iCodigo = iCodigo

            'Tenta ler Transportadora com esse código no BD
            lErro = CF("Transportadora_Le", objTransportadora)
            If lErro <> SUCESSO And lErro <> 19250 Then Error 53302

            ''Não encontrou Transportadora no BD
            If lErro = 19250 Then Error 53303

            'Encontrou Transportadora no BD, coloca no Text da Combo
            Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNomeReduzido

        End If

    End If

    'Não existe o ítem com a String na List da Combo
    If lErro = 6731 Then Error 53304

    Exit Sub

Erro_Transportadora_Validate:

    Cancel = True

    Select Case Err

        Case 53301, 53302

        Case 53303
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", Err, objTransportadora.iCodigo)

        Case 53304
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", Err, Transportadora.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164455)

    End Select

    Exit Sub

End Sub

Private Sub BotaoImprimir_Click()

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras

On Error GoTo Erro_BotaoImprimir_Click

    lErro = Move_Tela_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then gError 53305

    lErro = PedidoCompra_Imprimir(objPedidoCompra)
    If lErro <> SUCESSO Then gError 140201

    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr
    
        Case 53305, 140201
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164456)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_GridBloqueios(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridBloqueios

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'TipoBloqueio
        Case iGrid_TipoBloqueio_Col
            lErro = Saida_Celula_TipoBloqueio(objGridInt)
            If lErro <> SUCESSO Then Error 53409

        'ResponsavelBL
        Case iGrid_ResponsavelBL_Col
            lErro = Saida_Celula_Responsavel(objGridInt)
            If lErro <> SUCESSO Then Error 53410

    End Select

    Saida_Celula_GridBloqueios = SUCESSO

    Exit Function

Erro_Saida_Celula_GridBloqueios:

    Saida_Celula_GridBloqueios = Err

    Select Case Err

        Case 53409, 53410

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164457)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridDistribuicao(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridDistribuicao

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Prod
        Case iGrid_Prod_Col
            lErro = Saida_Celula_Prod(objGridInt)
            If lErro <> SUCESSO Then Error 53403

        'CentroCusto
        Case iGrid_CentroCusto_Col
            lErro = Saida_Celula_CentroCusto(objGridInt)
            If lErro <> SUCESSO Then Error 53404

        'Almoxarifado
        Case iGrid_Almoxarifado_Col
            lErro = Saida_Celula_Almoxarifado(objGridInt)
            If lErro <> SUCESSO Then Error 53405

        'UM
        Case iGrid_UM_Col
            lErro = Saida_Celula_UnidadeMed(objGridInt)
            If lErro <> SUCESSO Then Error 53406

        'Quant
        Case iGrid_Quant_Col
            lErro = Saida_Celula_Quant(objGridInt)
            If lErro <> SUCESSO Then Error 53407

        'ContaContabil
        Case iGrid_ContaContabil_Col
            lErro = Saida_Celula_ContaContabil(objGridInt)
            If lErro <> SUCESSO Then Error 53408

    End Select

    Saida_Celula_GridDistribuicao = SUCESSO

    Exit Function

Erro_Saida_Celula_GridDistribuicao:

    Saida_Celula_GridDistribuicao = Err

    Select Case Err

        Case 53403 To 53408

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164458)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridItens(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridItens

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        Case iGrid_DescProduto_Col
            lErro = Saida_Celula_Descricao(objGridInt)
            If lErro <> SUCESSO Then gError 86172

        'PercentMaisReceb
        Case iGrid_PercentMaisReceb_Col
            lErro = Saida_Celula_PercentMaisReceb(objGridInt)
            If lErro <> SUCESSO Then gError 53399

        'RecebForaFaixa
        Case iGrid_RecebForaFaixa_Col
            lErro = Saida_Celula_RecebForaFaixa(objGridInt)
            If lErro <> SUCESSO Then gError 53401

        'Observacao
        Case iGrid_Observacao_Col
            lErro = Saida_Celula_Observacao(objGridInt)
            If lErro <> SUCESSO Then gError 53402

        'DataLimite
        Case iGrid_DataLimite_Col
            lErro = Saida_Celula_DataLimite(objGridInt)
            If lErro <> SUCESSO Then gError 56277

        'AliquotaIPI
        Case iGrid_AliquotaIPI_Col
            lErro = Saida_Celula_AliquotaIPI(objGridInt)
            If lErro <> SUCESSO Then gError 56278

        'ValorIPIItem
        Case iGrid_ValorIPIItem_Col
            lErro = Saida_Celula_ValorIPI(objGridInt)
            If lErro <> SUCESSO Then gError 56279

        'AliquotaICMS
        Case iGrid_AliquotaICMS_Col
            lErro = Saida_Celula_AliquotaICMS(objGridInt)
            If lErro <> SUCESSO Then gError 56280

    End Select

    Saida_Celula_GridItens = SUCESSO

    Exit Function

Erro_Saida_Celula_GridItens:

    Saida_Celula_GridItens = gErr

    Select Case gErr

        Case 53399, 53400, 53401, 53402, 56277 To 56280, 86172

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164459)

    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'Se for o GridItens
            Case GridItens.Name
                lErro = Saida_Celula_GridItens(objGridInt)
                If lErro <> SUCESSO Then gError 53395

            'se for o GridBloqueios
            Case GridBloqueios.Name

                lErro = Saida_Celula_GridBloqueios(objGridInt)
                If lErro <> SUCESSO Then gError 53396

            'se for o GridDistribuicao
            Case GridDistribuicao.Name

                lErro = Saida_Celula_GridDistribuicao(objGridInt)
                If lErro <> SUCESSO Then gError 53397

            Case GridNotas.Name '??? por leo
                
                lErro = Saida_Celula_GridNotas(objGridInt)
                If lErro <> SUCESSO Then gError 103351

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 53398

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 53395, 53396, 53397, 53398, 103351
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164460)

    End Select

    Exit Function

End Function

Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_EnterCell()

    Call Grid_Entrada_Celula(objGridItens, iAlterado)

End Sub

Private Sub GridItens_GotFocus()

    Call Grid_Recebe_Foco(objGridItens)

End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_LeaveCell()

    Call Saida_Celula(objGridItens)

End Sub

Private Sub GridItens_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridItens)

End Sub

Private Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGridItens)

End Sub

Private Sub GridItens_Scroll()

    Call Grid_Scroll(objGridItens)

End Sub
Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridItens)

End Sub

Private Sub GridBloqueios_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridBloqueio, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridBloqueio, iAlterado)
    End If

End Sub

Private Sub GridBloqueios_EnterCell()

    Call Grid_Entrada_Celula(objGridBloqueio, iAlterado)

End Sub

Private Sub GridBloqueios_GotFocus()

    Call Grid_Recebe_Foco(objGridBloqueio)

End Sub

Private Sub GridBloqueios_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridBloqueio, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridBloqueio, iAlterado)
    End If

End Sub

Private Sub GridBloqueios_LeaveCell()

    Call Saida_Celula(objGridBloqueio)

End Sub

Private Sub GridBloqueios_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridBloqueio)

End Sub

Private Sub GridBloqueios_RowColChange()

    Call Grid_RowColChange(objGridBloqueio)

End Sub

Private Sub GridBloqueios_Scroll()

    Call Grid_Scroll(objGridBloqueio)

End Sub

Private Sub GridDistribuicao_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridDistribuicao, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDistribuicao, iAlterado)
    End If

End Sub

Private Sub GridDistribuicao_EnterCell()

    Call Grid_Entrada_Celula(objGridDistribuicao, iAlterado)

End Sub

Private Sub GridDistribuicao_GotFocus()

    Call Grid_Recebe_Foco(objGridDistribuicao)

End Sub

Private Sub GridDistribuicao_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridDistribuicao)

End Sub

Private Sub GridDistribuicao_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridDistribuicao, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDistribuicao, iAlterado)
    End If

End Sub

Private Sub GridDistribuicao_LeaveCell()

    Call Saida_Celula(objGridDistribuicao)

End Sub

Private Sub GridDistribuicao_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridDistribuicao)

End Sub

Private Sub GridDistribuicao_RowColChange()

    Call Grid_RowColChange(objGridDistribuicao)

End Sub

Private Sub GridDistribuicao_Scroll()

    Call Grid_Scroll(objGridDistribuicao)

End Sub
Private Function Saida_Celula_PercentMaisReceb(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dPercentMaisReceb As Double

On Error GoTo Erro_Saida_Celula_PercentMaisReceb

    Set objGridInt.objControle = PercentMaisReceb

    'Verifica se PercentMaisReceb foi preenchido
    If Len(Trim(PercentMaisReceb.Text)) > 0 Then

        'Critica porcentagem informada
        lErro = Porcentagem_Critica(PercentMaisReceb.Text)
        If lErro <> SUCESSO Then Error 53308

        dPercentMaisReceb = StrParaDbl(PercentMaisReceb.Text)

        'Coloca no formato fixed da tela
        PercentMaisReceb.Text = Format(dPercentMaisReceb, "Fixed")

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 53309

    Saida_Celula_PercentMaisReceb = SUCESSO

    Exit Function

Erro_Saida_Celula_PercentMaisReceb:

    Saida_Celula_PercentMaisReceb = Err

    Select Case Err

        Case 53308

        Case 53309
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164461)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_RecebForaFaixa(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_RecebForaFaixa

    Set objGridInt.objControle = RecebForaFaixa

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 53312

    Saida_Celula_RecebForaFaixa = SUCESSO

    Exit Function

Erro_Saida_Celula_RecebForaFaixa:

    Saida_Celula_RecebForaFaixa = Err

    Select Case Err

        Case 53312
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164462)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Observacao(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Observacao

    Set objGridInt.objControle = Observ

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 53313

    Saida_Celula_Observacao = SUCESSO

    Exit Function

Erro_Saida_Celula_Observacao:

    Saida_Celula_Observacao = Err

    Select Case Err

        Case 53313
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164463)

    End Select

    Exit Function


End Function

Private Function Saida_Celula_Prod(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim bProdutoPresente As Boolean
Dim iIndice As Integer
Dim iLinha As Integer
Dim sProduto As String

On Error GoTo Erro_Saida_Celula_Prod

    Set objGridInt.objControle = Prod

    bProdutoPresente = False
    
    If Len(Trim(Prod.Text)) > 0 Then

        lErro = CF("Produto_Critica", Prod.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 Then Error 53314

        'Se nao encontrou
       If lErro = 25041 Then Error 53315

        'Verifica se já está em outra linha do Grid
        For iIndice = 1 To objGridItens.iLinhasExistentes
            If (GridItens.TextMatrix(iIndice, iGrid_Produto_Col)) = Prod.Text Then
                bProdutoPresente = True
                Exit For
            End If
        Next

        'Guarda a linha do produto no grid itens
        iLinha = iIndice

        If bProdutoPresente = False Then Error 53316

        'Coloca descricao do produto no Grid Distribuicao
        GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_DescProduto_Col) = objProduto.sDescricao
        GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_UM_Col) = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)

        bProdutoPresente = False

        'Verifica se está presente no Grid de Distribuicao
        For iIndice = 1 To objGridDistribuicao.iLinhasExistentes
            If iIndice <> GridDistribuicao.Row Then
                If GridDistribuicao.TextMatrix(iIndice, iGrid_Prod_Col) = Prod.Text Then
                    bProdutoPresente = True
                End If
            End If
        Next

        If bProdutoPresente = False And Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))) <> 0 Then

            GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_Quant_Col) = GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col)
            GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_UM_Col) = GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)
        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridDistribuicao.Row - GridDistribuicao.FixedRows = objGridDistribuicao.iLinhasExistentes Then
            objGridDistribuicao.iLinhasExistentes = objGridDistribuicao.iLinhasExistentes + 1
        End If

    End If

    'GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_Prod_Col) = Prod.Text
    
    sProduto = Prod.Text
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 53317
    
    GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_Prod_Col) = sProduto
    
    Saida_Celula_Prod = SUCESSO


    Exit Function

Erro_Saida_Celula_Prod:
   
    Saida_Celula_Prod = Err

    Select Case Err

        Case 53314, 53317
            Call Grid_Trata_Erro_Saida_Celula(objGridItens)

        Case 53315
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err, Prod.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridItens)

        Case 53316
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_DESVINCULADO_ITEM", Err, Prod.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridItens)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164464)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_CentroCusto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sCclFormatada As String
Dim objCcl As New ClassCcl

On Error GoTo Erro_Saida_Celula_CentroCusto

    Set objGridInt.objControle = CentroCusto

    'Verifica se CentroCusto foi preenchido
    If Len(Trim(CentroCusto.ClipText)) > 0 Then

        'Critica o Ccl
        lErro = CF("Ccl_Critica", CentroCusto, sCclFormatada, objCcl)
        If lErro <> SUCESSO And lErro <> 5703 Then Error 53318

        'Se o Ccl nao existe
        If lErro = 5703 Then Error 53319

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 53320

    Saida_Celula_CentroCusto = SUCESSO

    Exit Function

Erro_Saida_Celula_CentroCusto:

    Saida_Celula_CentroCusto = Err

    Select Case Err

        Case 53318
             Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 53319
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", Err, CentroCusto.Text)
             Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 53320
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164465)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Almoxarifado(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Almoxarifado

    Set objGridInt.objControle = Almoxarifado

    'Se o Almoxarifado está preenchido
    If Len(Trim(Almoxarifado.Text)) > 0 Then

        'Formata o produto
        lErro = CF("Produto_Formata", GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_Prod_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then Error 53321

        'Valida o Almoxarifado
        lErro = TP_Almoxarifado_Produto_Grid(sProdutoFormatado, Almoxarifado, objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25130 And lErro <> 25135 Then Error 53322

        'Se não for encontrado --> Erro
        If lErro = 25130 Then Error 53323
        If lErro = 25135 Then Error 53324
        If objAlmoxarifado.iFilialEmpresa <> Codigo_Extrai(FilialEmpresa.Text) Then gError 86121

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 53325

    Saida_Celula_Almoxarifado = SUCESSO

    Exit Function

Erro_Saida_Celula_Almoxarifado:

    Saida_Celula_Almoxarifado = Err

    Select Case Err

        Case 53321, 53322, 53325
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 53323
            'Pergunta de deseja criar o Almoxarifado
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ALMOXARIFADO2", Almoxarifado.Text)
            'Se a resposta for sim
            If vbMsg = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                objAlmoxarifado.sNomeReduzido = Almoxarifado.Text

                'Chama a Tela Almoxarifados
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 53324

            'Pergunta se deseja criar o Almoxarifado
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ALMOXARIFADO1", Codigo_Extrai(Almoxarifado.Text))
            'Se a resposta for positiva
            If vbMsg = vbYes Then

                objAlmoxarifado.iCodigo = Codigo_Extrai(Almoxarifado.Text)

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                'Chama a tela de Almoxarifados
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)


            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 86121
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_FILIALENTREGA", gErr, objAlmoxarifado.iCodigo & SEPARADOR & objAlmoxarifado.sNomeReduzido, Codigo_Extrai(FilialEmpresa.Text))
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 164466)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UnidadeMed(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_UnidadeMed

    Set objGridInt.objControle = UnidadeMed

    objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_UM_Col) = UnidadeMed.Text
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 53326

    Saida_Celula_UnidadeMed = SUCESSO

    Exit Function

Erro_Saida_Celula_UnidadeMed:

    Saida_Celula_UnidadeMed = Err

    Select Case Err

        Case 53326
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164467)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quant(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantidade As Double

On Error GoTo Erro_Saida_Celula_Quant

    Set objGridInt.objControle = Quant

    'Se quantidade estiver preenchida
    If Len(Trim(Quant.ClipText)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(Quant.Text)
        If lErro <> SUCESSO Then Error 53327

        dQuantidade = CDbl(Quant.Text)

        'Coloca o valor Formatado na tela
        Quant.Text = Formata_Estoque(dQuantidade)

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 53328

    Saida_Celula_Quant = SUCESSO

    Exit Function

Erro_Saida_Celula_Quant:

    Saida_Celula_Quant = Err

    Select Case Err

        Case 53327, 53328
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164468)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ContaContabil(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaEnxuta As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim iContaPreenchida As Integer

On Error GoTo Erro_Saida_Celula_ContaContabil

    Set objGridInt.objControle = ContaContabil

    'Se a Conta Contábil foi preenchida
    If Len(Trim(ContaContabil.ClipText)) > 0 Then

        'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
        lErro = CF("ContaSimples_Critica_Modulo", sContaFormatada, ContaContabil.ClipText, objPlanoConta, MODULO_COMPRAS)
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then Error 61942

        If lErro = SUCESSO Then

            sContaFormatada = objPlanoConta.sConta

            'mascara a conta
            sContaEnxuta = String(STRING_CONTA, 0)

            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
            If lErro <> SUCESSO Then Error 61945

            ContaContabil.PromptInclude = False
            ContaContabil.Text = sContaEnxuta
            ContaContabil.PromptInclude = True

        'se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then

            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaContabil.Text, sContaFormatada, objPlanoConta, MODULO_COMPRAS)
            If lErro <> SUCESSO And lErro <> 5700 Then Error 61943

            'conta não cadastrada
            If lErro = 5700 Then Error 61946

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 61944

    Saida_Celula_ContaContabil = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaContabil:

    Saida_Celula_ContaContabil = Err

    Select Case Err

        Case 61942, 61943, 61944
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 61945
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 61946
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabil.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164469)

    End Select

End Function

Private Function Saida_Celula_TipoBloqueio(objGridInt As AdmGrid)
'Faz a crítica da célula Tipo Bloqueio que está deixando de ser a corrente

Dim lErro As Long
Dim iCodigo As Integer
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_TipoBloqueio

    Set objGridInt.objControle = TipoBloqueio

    'Verifica se o Tipo foi preenchido
    If Len(Trim(TipoBloqueio.Text)) > 0 Then

        'Tenta selecioná-lo na combo
        lErro = Combo_Seleciona_Grid(TipoBloqueio, iCodigo)
        If lErro <> SUCESSO And lErro <> 25085 And lErro <> 25086 Then Error 56162 '53334

        'Não foi encontrado
        If lErro = 25085 Then Error 56163 '53335
        If lErro = 25086 Then Error 56164 '53336

        'Verifica se já existe um bloqueio, não liberado, do mesmo tipo no grid.
        For iIndice = 1 To objGridInt.iLinhasExistentes

            If iIndice <> GridBloqueios.Row Then
                iCodigo = Codigo_Extrai(GridBloqueios.TextMatrix(iIndice, iGrid_TipoBloqueio_Col))

                If iCodigo = TipoBloqueio.ItemData(TipoBloqueio.ListIndex) Then
                    If Len(Trim(GridBloqueios.TextMatrix(iIndice, iGrid_DataLiberacao_Col))) = 0 Then Error 63852
                End If
            End If

        Next

        'Acrescenta uma linha no Grid se for o caso
        If GridBloqueios.Row - GridBloqueios.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

        'Coloca data e usuário
        GridBloqueios.TextMatrix(GridBloqueios.Row, iGrid_DataBloqueio_Col) = Format(gdtDataAtual, "dd/mm/yy")
        GridBloqueios.TextMatrix(GridBloqueios.Row, iGrid_CodUsuario_Col) = Format(gsUsuario)

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 56165 '53337

    Saida_Celula_TipoBloqueio = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoBloqueio:

    Saida_Celula_TipoBloqueio = Err

    Select Case Err

        Case 53334, 53337
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 53335
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOBLOQUEIO_NAO_ENCONTRADO", Err, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 53336
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOBLOQUEIO_NAO_ENCONTRADO1", Err, TipoBloqueio.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 63852
            Call Rotina_Erro(vbOKOnly, "ERRO_BLOQUEIOPV_REPETIDO", Err, Codigo_Extrai(TipoBloqueio.Text))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164470)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Responsavel(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Responsavel

    Set objGridInt.objControle = ResponsavelBL

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 53338

    Saida_Celula_Responsavel = SUCESSO

    Exit Function

Erro_Saida_Celula_Responsavel:

    Saida_Celula_Responsavel = Err

    Select Case Err

        Case 53338
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164471)

    End Select

    Exit Function

End Function

Private Sub BotaoBaixar_Click()

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoBaixar

    GL_objMDIForm.MousePointer = vbHourglass
    
    'se o numero do pedido nao estiver preenchido ---> erro
    If Len(Trim(Codigo.Text)) = 0 Then Error 53379

    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then Error 53380

    've se o Pedido de Compra foi baixado
    lErro = CF("PedidoCompraBaixado_Le_Numero", objPedidoCompra)
    If lErro <> SUCESSO And lErro <> 56137 Then Error 53381
    'se foi ---> erro
    If lErro = SUCESSO Then Error 53382

    'procura na tabela de Pedido de Compras
    lErro = CF("PedidoCompra_Le_Numero", objPedidoCompra)
    If lErro <> SUCESSO And lErro <> 56142 Then Error 53383

    'se nao encontrar ---> erro
    If lErro = 56142 Then Error 53384

    'Pede a confirmação da baixa do pedido
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_BAIXA_PEDIDOCOMPRAS", objPedidoCompra.lCodigo)

    If vbMsgRes = vbYes Then

        'baixa o Pedido de Compras
        lErro = CF("PedidoCompra_Baixar", objPedidoCompra)
        If lErro <> SUCESSO Then Error 53385

        'limpa a tela
        Call Limpa_Tela_PedidoCompras

        'Fecha o comando das setas se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)

        iAlterado = 0

    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoBaixar:

    Select Case Err

        Case 53379
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_PEDIDO_NAO_PREENCHIDO", Err)

        Case 53380, 53381, 53383, 53385

        Case 53382
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDCOMPRA_BAIXADO", Err, objPedidoCompra.lCodigo)

        Case 53384
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", Err, objPedidoCompra.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164472)

    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

End Sub
Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 53339

    'Limpa a tela
    Call Limpa_Tela_PedidoCompras

    'Fecha o comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 53339

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164473)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim dValorTotal As Double
Dim objPedidoCompra As New ClassPedidoCompras

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se Codigo foi preenchido
    If (Len(Trim(Codigo.Text))) = 0 Then gError 53339

    'se o TipoDestino for FilialEmpresa
    If TipoDestino(TIPO_DESTINO_EMPRESA) Then

        'verifica se a FilialEmpresa foi informada
        If Len(Trim(FilialEmpresa.Text)) = 0 Then gError 53340

    'se o TipoDestino for fornecedor
    ElseIf TipoDestino(TIPO_DESTINO_FORNECEDOR) Then

        'verifica se o fornecedor esta preenchido
        If Len(Trim(Fornec.Text)) = 0 Then gError 53341
        'verifica se a filial do fornecedor esta preenchida
        If Len(Trim(FilialFornec.Text)) = 0 Then gError 53342

    End If

    'Verifica se ValorTotal foi preenchido
    dValorTotal = StrParaDbl(ValorTotal.Caption)
    If dValorTotal < 0 Then gError 53345

    'Valida os dados do Grid Distribuicao
    lErro = Valida_Grid_Distribuicao()
    If lErro <> SUCESSO Then gError 53346

    'Valida os dados do Grid Bloqueios
    lErro = Valida_Grid_Bloqueios()
    If lErro <> SUCESSO Then gError 53347

    'Recolhe os dados da tela
    lErro = Move_Pedido_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then gError 53348

    'Grava o Pedido de Compra
    lErro = CF("PedidoCompraGerado_Grava", objPedidoCompra)
    If lErro <> SUCESSO Then gError 53349

    '########################################
    'Inserido por Wagner
    If ImprimePedido.Value = vbChecked Then
    
        lErro = PedidoCompra_Imprimir(objPedidoCompra)
        If lErro <> SUCESSO Then gError 140202
    
    End If
    '########################################
    
    Gravar_Registro = SUCESSO

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 53339
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_PREENCHIDO", gErr)

        Case 53340
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_DESTINO_NAO_PREENCHIDA", gErr)

         Case 53341
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_DESTINO_NAO_PREENCHIDO", gErr)

        Case 53342
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORN_DESTINO_NAO_PREENCHIDA", gErr)

        Case 53343
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_DESTINO_NAO_PREENCHIDO", gErr)

        Case 53344
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_INFORMADA", gErr)

        Case 53345
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_PC_NEGATIVO", gErr)

        Case 53346, 53347, 53348, 53349, 140202
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164474)

    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

End Function
Private Function Valida_Grid_Distribuicao() As Long

Dim lErro As Long
Dim iIndice As Integer, iIndice2 As Integer
Dim dQuantidade As Double

On Error GoTo Erro_Valida_Grid_Distribuicao

    For iIndice = 1 To objGridDistribuicao.iLinhasExistentes

        'Verifica se Produto foi preenchido
        If Len(Trim(GridDistribuicao.TextMatrix(iIndice, iGrid_Prod_Col))) = 0 Then gError 53350

        'Verifica se Almoxarifado foi preenchido
        If Len(Trim(GridDistribuicao.TextMatrix(iIndice, iGrid_Almoxarifado_Col))) = 0 Then gError 53351

    Next

    For iIndice2 = 1 To objGridItens.iLinhasExistentes
        
        dQuantidade = 0
        
        For iIndice = 1 To objGridDistribuicao.iLinhasExistentes
            'Acumula a quantidade distribuida
            If GridDistribuicao.TextMatrix(iIndice, iGrid_Prod_Col) = GridItens.TextMatrix(iIndice2, iGrid_Produto_Col) Then dQuantidade = dQuantidade + StrParaDbl(GridDistribuicao.TextMatrix(iIndice, iGrid_Quant_Col))
        Next
        
        If dQuantidade > StrParaDbl(GridItens.TextMatrix(iIndice2, iGrid_Quantidade_Col)) Then gError 81627
    Next

    'verificar se nao há repeticao de produto-almoxarifado-ccl
    For iIndice = 1 To objGridDistribuicao.iLinhasExistentes
    
        For iIndice2 = 1 To objGridDistribuicao.iLinhasExistentes
        
            If iIndice2 <> iIndice And _
                GridDistribuicao.TextMatrix(iIndice, iGrid_Prod_Col) = GridDistribuicao.TextMatrix(iIndice2, iGrid_Produto_Col) And _
                GridDistribuicao.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = GridDistribuicao.TextMatrix(iIndice2, iGrid_Almoxarifado_Col) And _
                GridDistribuicao.TextMatrix(iIndice, iGrid_CentroCusto_Col) = GridDistribuicao.TextMatrix(iIndice2, iGrid_CentroCusto_Col) Then gError 81626
        
        Next
    
    Next
    
    Valida_Grid_Distribuicao = SUCESSO

    Exit Function

Erro_Valida_Grid_Distribuicao:

    Valida_Grid_Distribuicao = gErr

    Select Case gErr

        Case 53350
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_ITEM_DISTRIBUICAO_VAZIO", gErr)

        Case 53351
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_ITEM_DISTRIBUICAO_VAZIO", gErr)

        Case 81627
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTDISTRIBUIDA_SUPERIOR_ITEM", gErr, iIndice2)

        Case 81626
            Call Rotina_Erro(vbOKOnly, "ERRO_LOC_ITEM_PC_DUPLICADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164475)

    End Select

    Exit Function

End Function

Private Function Valida_Grid_Bloqueios() As Long

Dim iIndice As Integer

On Error GoTo Erro_Valida_Grid_Bloqueios

    For iIndice = 1 To objGridBloqueio.iLinhasExistentes

        'Testa se preencheu tipo de bloqueio
        If Len(Trim(GridBloqueios.TextMatrix(iIndice, iGrid_TipoBloqueio_Col))) = 0 Then Error 53352

    Next

    Valida_Grid_Bloqueios = SUCESSO

    Exit Function

Erro_Valida_Grid_Bloqueios:

    Valida_Grid_Bloqueios = Err

    Select Case Err

        Case 53352
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOBLOQUEIO_BLOQUEIO_NAO_INFORMADO", Err, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164476)

    End Select

    Exit Function

End Function

Private Function Move_Pedido_Memoria(objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Pedido_Memoria

    lErro = Move_Tela_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then gError 53353

    lErro = Move_GridItens_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then gError 53354

    lErro = Move_GridDistribuicao_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then gError 53355

    lErro = Move_GridBloqueio_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then gError 53356

    lErro = Move_GridNotas_Memoria(objPedidoCompra) '??? por leo
    If lErro <> SUCESSO Then gError 103354

    Move_Pedido_Memoria = SUCESSO

    Exit Function

Erro_Move_Pedido_Memoria:

    Move_Pedido_Memoria = gErr

    Select Case gErr

        Case 53353 To 53356, 103354
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164477)

    End Select

    Exit Function

End Function

Private Function Move_GridItens_Memoria(objPedidoCompra As ClassPedidoCompras) As Long
'Recolhe do Grid os dados do item pedido no parametro

Dim lErro As Long
Dim sProduto As String
Dim objItemPedCompra As New ClassItemPedCompra
Dim iPreenchido As Integer
Dim iIndice As Integer
Dim iItem As Integer

On Error GoTo Erro_Move_GridItens_Memoria

    For iIndice = 1 To objGridItens.iLinhasExistentes

        Set objItemPedCompra = New ClassItemPedCompra

        'Verifica se o Produto está preenchido
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Produto_Col))) > 0 Then

            'Formata o produto
            lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProduto, iPreenchido)
            If lErro <> SUCESSO Then Error 53357

            objItemPedCompra.sProduto = sProduto
        End If

        'Armazena os dados do item
        objItemPedCompra.sUM = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
        objItemPedCompra.dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        objItemPedCompra.dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col))
        objItemPedCompra.dtDataLimite = StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataLimite_Col))
        objItemPedCompra.sObservacao = GridItens.TextMatrix(iIndice, iGrid_Observacao_Col)
        objItemPedCompra.dPercentMaisReceb = PercentParaDbl(GridItens.TextMatrix(iIndice, iGrid_PercentMaisReceb_Col))
        objItemPedCompra.dAliquotaICMS = PercentParaDbl(GridItens.TextMatrix(iIndice, iGrid_AliquotaICMS_Col))
        objItemPedCompra.dAliquotaIPI = PercentParaDbl(GridItens.TextMatrix(iIndice, iGrid_AliquotaIPI_Col))
        objItemPedCompra.dValorIPI = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorIPIItem_Col))

        For iItem = 0 To RecebForaFaixa.ListCount - 1
            If GridItens.TextMatrix(iIndice, iGrid_RecebForaFaixa_Col) = RecebForaFaixa.List(iItem) Then

                'coloca no Grid Itens RecebForaFaixa
                objItemPedCompra.iRebebForaFaixa = RecebForaFaixa.ItemData(iItem)

            End If
        Next

        objItemPedCompra.sDescProduto = GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col)
        objItemPedCompra.lNumIntDoc = gcolItemPedido.Item(iIndice)

        'Adiciona o item na colecao de itens do pedido de compra
        objPedidoCompra.colItens.Add objItemPedCompra

    Next

    Move_GridItens_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItens_Memoria:

    Move_GridItens_Memoria = Err

    Select Case Err

        Case 53357

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164478)

    End Select

    Exit Function

End Function

Private Function Move_GridDistribuicao_Memoria(objPedidoCompra As ClassPedidoCompras) As Long
'Recolhe do Grid os dados do item pedido no parametro

Dim lErro As Long
Dim objItemPC As ClassItemPedCompra
Dim objLocalizacao As New ClassLocalizacaoItemPC
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim sCcl As String
Dim iCclPreenchida As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sContaFormatada As String
Dim iContaPreenchida As Integer

On Error GoTo Erro_Move_GridDistribuicao_Memoria

    iIndice = 0

    For Each objItemPC In objPedidoCompra.colItens

        iIndice = iIndice + 1

        For iIndice2 = 1 To objGridDistribuicao.iLinhasExistentes

            If GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = GridDistribuicao.TextMatrix(iIndice2, iGrid_Prod_Col) Then
                
                Set objLocalizacao = New ClassLocalizacaoItemPC
                
                'preenche objLocalizacao com os dados do grid
                objLocalizacao.dQuantidade = StrParaDbl(GridDistribuicao.TextMatrix(iIndice2, iGrid_Quant_Col))

                If Len(Trim(GridDistribuicao.TextMatrix(iIndice2, iGrid_Almoxarifado_Col))) > 0 Then

                    objAlmoxarifado.sNomeReduzido = GridDistribuicao.TextMatrix(iIndice2, iGrid_Almoxarifado_Col)
                    lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                    If lErro <> SUCESSO And lErro <> 25060 Then Error 53360

                    'se nao encontrou ---> erro
                    If lErro = 25060 Then Error 53359

                    objLocalizacao.iAlmoxarifado = objAlmoxarifado.iCodigo

                End If

                If Len(Trim(GridDistribuicao.TextMatrix(iIndice2, iGrid_CentroCusto_Col))) > 0 Then

                    lErro = CF("Ccl_Formata", GridDistribuicao.TextMatrix(iIndice2, iGrid_CentroCusto_Col), sCcl, iCclPreenchida)
                    If lErro <> SUCESSO Then Error 53361

                    objLocalizacao.sCcl = sCcl

                End If

                If Len(Trim(GridDistribuicao.TextMatrix(iIndice2, iGrid_ContaContabil_Col))) > 0 Then

                    lErro = CF("Conta_Formata", GridDistribuicao.TextMatrix(iIndice2, iGrid_ContaContabil_Col), sContaFormatada, iContaPreenchida)
                    If lErro <> SUCESSO Then Error 53388

                    objLocalizacao.sContaContabil = sContaFormatada

                End If

                Set objItemPC = objPedidoCompra.colItens.Item(iIndice)

                objItemPC.colLocalizacao.Add objLocalizacao

            End If

        Next

    Next

    Move_GridDistribuicao_Memoria = SUCESSO

    Exit Function

Erro_Move_GridDistribuicao_Memoria:

    Move_GridDistribuicao_Memoria = Err

    Select Case Err

        Case 53360, 53361, 53388

        Case 53359
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", Err, objAlmoxarifado.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164479)

    End Select

    Exit Function

End Function

Private Function Move_GridBloqueio_Memoria(objPedidoCompra As ClassPedidoCompras) As Long
'Recolhe do Grid os dados do bloqueio no parametro

Dim lErro As Long
Dim iIndice As Integer
Dim objBloqueioPC As New ClassBloqueioPC

On Error GoTo Erro_Move_GridBloqueio_Memoria

    For iIndice = iGridBloqueiosLinhasBloqueadas + 1 To objGridBloqueio.iLinhasExistentes

        Set objBloqueioPC = New ClassBloqueioPC

        'Armazena os dados do bloqueio
        objBloqueioPC.dtData = StrParaDate(GridBloqueios.TextMatrix(iIndice, iGrid_DataBloqueio_Col))
        objBloqueioPC.dtDataLib = StrParaDate(GridBloqueios.TextMatrix(iIndice, iGrid_DataLiberacao_Col))
        objBloqueioPC.iTipoBloqueio = Codigo_Extrai(GridBloqueios.TextMatrix(iIndice, iGrid_TipoBloqueio_Col))
        objBloqueioPC.sCodUsuario = GridBloqueios.TextMatrix(iIndice, iGrid_CodUsuario_Col)
        objBloqueioPC.sCodUsuarioLib = GridBloqueios.TextMatrix(iIndice, iGrid_ResponsavelLib_Col)
        objBloqueioPC.sResponsavel = GridBloqueios.TextMatrix(iIndice, iGrid_ResponsavelBL_Col)
        objBloqueioPC.iFilialEmpresa = giFilialEmpresa

        'Adiciona o item na colecao de itens do pedido de compra
        objPedidoCompra.colBloqueiosPC.Add objBloqueioPC

    Next

    Move_GridBloqueio_Memoria = SUCESSO

    Exit Function

Erro_Move_GridBloqueio_Memoria:

    Move_GridBloqueio_Memoria = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164480)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Pedido de Compra Gerado - Não Enviado"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "PedComprasGerado"
    
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

Private Sub Filial_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Filial, Source, X, Y)
End Sub

Private Sub Filial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Filial, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label6(Index), Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6(Index), Button, Shift, X, Y)
End Sub

Private Sub FornecedorLabel_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(FornecedorLabel, Source, X, Y)
End Sub

Private Sub FornecedorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel, Button, Shift, X, Y)
End Sub

Private Sub Fornecedor_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Fornecedor, Source, X, Y)
End Sub

Private Sub Fornecedor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Fornecedor, Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label15(Index), Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15(Index), Button, Shift, X, Y)
End Sub


Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub CodigoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CodigoLabel, Source, X, Y)
End Sub

Private Sub CodigoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CodigoLabel, Button, Shift, X, Y)
End Sub

Private Sub CondPagtoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CondPagtoLabel, Source, X, Y)
End Sub

Private Sub CondPagtoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CondPagtoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
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

Private Sub Label29_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label29, Source, X, Y)
End Sub

Private Sub Label29_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label29, Button, Shift, X, Y)
End Sub

Private Sub DataEmissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataEmissao, Source, X, Y)
End Sub

Private Sub DataEmissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataEmissao, Button, Shift, X, Y)
End Sub

Private Sub Label30_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label30, Source, X, Y)
End Sub

Private Sub Label30_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label30, Button, Shift, X, Y)
End Sub

Private Sub DataAlteracao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataAlteracao, Source, X, Y)
End Sub

Private Sub DataAlteracao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataAlteracao, Button, Shift, X, Y)
End Sub

Private Sub CondPagto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CondPagto, Source, X, Y)
End Sub

Private Sub CondPagto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CondPagto, Button, Shift, X, Y)
End Sub

Private Sub Codigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Codigo, Source, X, Y)
End Sub

Private Sub Codigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Codigo, Button, Shift, X, Y)
End Sub

Private Sub Data_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Data, Source, X, Y)
End Sub

Private Sub Data_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Data, Button, Shift, X, Y)
End Sub

Private Sub Label41_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label41, Source, X, Y)
End Sub

Private Sub Label41_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label41, Button, Shift, X, Y)
End Sub

Private Sub ValorProdutos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorProdutos, Source, X, Y)
End Sub

Private Sub ValorProdutos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorProdutos, Button, Shift, X, Y)
End Sub

Private Sub Label19_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label19, Source, X, Y)
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label19, Button, Shift, X, Y)
End Sub

Private Sub Label20_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label20, Source, X, Y)
End Sub

Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label20, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub ValorTotal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorTotal, Source, X, Y)
End Sub

Private Sub ValorTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorTotal, Button, Shift, X, Y)
End Sub

Private Sub Label25_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label25, Source, X, Y)
End Sub

Private Sub Label25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label25, Button, Shift, X, Y)
End Sub

Private Sub Label40_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label40, Source, X, Y)
End Sub

Private Sub Label40_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label40, Button, Shift, X, Y)
End Sub

Private Sub FornLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornLabel, Source, X, Y)
End Sub

Private Sub FornLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornLabel, Button, Shift, X, Y)
End Sub

Private Sub Label73_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label73, Source, X, Y)
End Sub

Private Sub Label73_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label73, Button, Shift, X, Y)
End Sub

Private Sub Label72_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label72, Source, X, Y)
End Sub

Private Sub Label72_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label72, Button, Shift, X, Y)
End Sub

Private Sub Label71_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label71, Source, X, Y)
End Sub

Private Sub Label71_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label71, Button, Shift, X, Y)
End Sub

Private Sub Label70_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label70, Source, X, Y)
End Sub

Private Sub Label70_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label70, Button, Shift, X, Y)
End Sub

Private Sub Label65_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label65, Source, X, Y)
End Sub

Private Sub Label65_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label65, Button, Shift, X, Y)
End Sub

Private Sub Label63_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label63, Source, X, Y)
End Sub

Private Sub Label63_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label63, Button, Shift, X, Y)
End Sub

Private Sub Endereco_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Endereco, Source, X, Y)
End Sub

Private Sub Endereco_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Endereco, Button, Shift, X, Y)
End Sub

Private Sub Bairro_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Bairro, Source, X, Y)
End Sub

Private Sub Bairro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Bairro, Button, Shift, X, Y)
End Sub

Private Sub Cidade_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Cidade, Source, X, Y)
End Sub

Private Sub Cidade_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Cidade, Button, Shift, X, Y)
End Sub

Private Sub CEP_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CEP, Source, X, Y)
End Sub

Private Sub CEP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CEP, Button, Shift, X, Y)
End Sub

Private Sub Estado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Estado, Source, X, Y)
End Sub

Private Sub Estado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Estado, Button, Shift, X, Y)
End Sub

Private Sub Pais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Pais, Source, X, Y)
End Sub

Private Sub Pais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Pais, Button, Shift, X, Y)
End Sub

Private Sub Label31_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label31, Source, X, Y)
End Sub

Private Sub Label31_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label31, Button, Shift, X, Y)
End Sub

Private Sub TransportadoraLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TransportadoraLabel, Source, X, Y)
End Sub

Private Sub TransportadoraLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TransportadoraLabel, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label4(Index), Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4(Index), Button, Shift, X, Y)
End Sub

Private Sub ObsLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ObsLabel, Source, X, Y)
End Sub

Private Sub ObsLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ObsLabel, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub
Private Function Saida_Celula_DataLimite(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dtData As Date
Dim dtDataLimite As Date

On Error GoTo Erro_Saida_Celula_DataLimite

    Set objGridInt.objControle = DataLimite

    'Verifica se a data limite foi preenchida
    If Len(Trim(DataLimite.ClipText)) > 0 Then

        'Critica a data limite
        lErro = Data_Critica(DataLimite.Text)
        If lErro <> SUCESSO Then Error 56078

        dtDataLimite = CDate(DataLimite.Text)

        'Verifica se a data do pedido foi preenchida
        If Len(Trim(Data.Caption)) > 0 Then

            dtData = CDate(Data.Caption)

            'comparacao entre data limite e data do pedido
            If dtDataLimite < dtData Then Error 56079

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 56080

    Saida_Celula_DataLimite = SUCESSO

    Exit Function

Erro_Saida_Celula_DataLimite:

    Saida_Celula_DataLimite = Err

    Select Case Err

        Case 56078
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56079
            Call Rotina_Erro(vbOKOnly, "ERRO_DATALIMITE_INFERIOR_DATAPEDIDO", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56080
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164481)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_AliquotaICMS(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dAliquotaICMS As Double

On Error GoTo Erro_Saida_Celula_AliquotaICMS

    Set objGridInt.objControle = AliquotaICM

    'Verifica se AliquotaICMS foi preenchida
    If Len(Trim(AliquotaICM.Text)) > 0 Then

        'Critica a aliquota informada
        lErro = Porcentagem_Critica(AliquotaICM.Text)
        If lErro <> SUCESSO Then Error 49475

        dAliquotaICMS = StrParaDbl(AliquotaICM.Text)

        'Coloca aliquota no formato fixed da tela
        AliquotaICM.Text = Format(dAliquotaICMS, "Fixed")

        If AliquotaICM.Text = 100# Then Error 49476

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 49477

    Saida_Celula_AliquotaICMS = SUCESSO

    Exit Function

Erro_Saida_Celula_AliquotaICMS:

    Saida_Celula_AliquotaICMS = Err

    Select Case Err

        Case 49475
            'Erro tratado na rotina chamada
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 49476
            Call Rotina_Erro(vbOKOnly, "ERRO_ALIQUOTA_IGUAL_100", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 49477
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164482)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_AliquotaIPI(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dAliquotaIPI As Double
Dim bAliquotaIPIIgual As Boolean
Dim dPrecoTotal As Double
Dim dValorTotal As Double
Dim dValorIPI As Double
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_AliquotaIPI

    Set objGridInt.objControle = AliquotaIPI

    If Len(Trim(AliquotaIPI.Text)) > 0 Then

        lErro = Porcentagem_Critica(AliquotaIPI.Text)
        If lErro <> SUCESSO Then Error 56099

        AliquotaIPI.Text = Format(AliquotaIPI.Text, "Fixed")

        'Verifica se AliquotaIPI foi alterada
        If PercentParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_AliquotaIPI_Col)) = StrParaDbl(AliquotaIPI.Text) Then bAliquotaIPIIgual = True

        If bAliquotaIPIIgual = False And Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col))) <> 0 Then

            dAliquotaIPI = StrParaDbl(AliquotaIPI.Text)
            If dAliquotaIPI = 100 Then Error 56100

            dPrecoTotal = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col))

            dValorIPI = dPrecoTotal * dAliquotaIPI / 100
            GridItens.TextMatrix(GridItens.Row, iGrid_ValorIPIItem_Col) = Format(dValorIPI, "Standard")
        End If

    Else  'se aliquota nao estiver preenchida

        GridItens.TextMatrix(GridItens.Row, iGrid_ValorIPIItem_Col) = ""

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 56101

    'Calcula valorIPI para o pedido
    For iIndice = 1 To objGridItens.iLinhasExistentes

        dValorIPI = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorIPIItem_Col))
        dValorTotal = dValorTotal + dValorIPI

    Next

    ValorIPI = Format(dValorTotal, "standard")

    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then Error 56122

    Saida_Celula_AliquotaIPI = SUCESSO

    Exit Function


Erro_Saida_Celula_AliquotaIPI:

    Saida_Celula_AliquotaIPI = Err

    Select Case Err

        Case 56099, 56101, 56122
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56100
            Call Rotina_Erro(vbOKOnly, "ERRO_ALIQUOTA_IGUAL_100", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164483)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorIPI(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dValorIPI As Double
Dim dPrecoTotal As Double
Dim dAliquotaIPI As Double
Dim bValorIPIIgual As Boolean
Dim iIndice As Integer
Dim dValorTotal As Double

On Error GoTo Erro_Saida_Celula_ValorIPI

    Set objGridInt.objControle = ValorIPIItem

    bValorIPIIgual = False

    'Verifica se ValorIPI foi preenchido
    If Len(Trim(ValorIPIItem.Text)) > 0 Then

        'Critica o valor informado
        lErro = Valor_Positivo_Critica(ValorIPIItem.Text)
        If lErro <> SUCESSO Then Error 56102

        'Coloca no formato standard da tela
        ValorIPIItem.Text = Format(ValorIPIItem.Text, "Standard")

        dValorIPI = StrParaDbl(ValorIPIItem)

    End If

    'Comparação com Valor IPI anterior
    If StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorIPIItem_Col)) = StrParaDbl(ValorIPIItem.Text) Then bValorIPIIgual = True

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 56104

    If bValorIPIIgual = False Then

        dPrecoTotal = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col))

        'Verifica se Preco Total foi preenchido
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col))) <> 0 And dPrecoTotal > 0 Then

            If dValorIPI > dPrecoTotal Then Error 56103

            dAliquotaIPI = dValorIPI / dPrecoTotal

            'coloca no formato percent
            GridItens.TextMatrix(GridItens.Row, iGrid_AliquotaIPI_Col) = Format(dAliquotaIPI, "Percent")

        End If

        'Calcula o valor IPI para o pedido
         For iIndice = 1 To objGridItens.iLinhasExistentes

            dValorIPI = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorIPIItem_Col))
            dValorTotal = dValorTotal + dValorIPI

        Next

        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then Error 56123

    End If

    Saida_Celula_ValorIPI = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorIPI:

    Saida_Celula_ValorIPI = Err

    Select Case Err

        Case 56102, 56104, 56123
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56103
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORIPI_MAIOR_OU_IGUAL_PRECO_TOTAL", Err, dValorIPI, dPrecoTotal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164484)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Descricao(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Descricao

    Set objGridInt.objControle = DescProduto
    
    If Len(Trim(DescProduto.Text)) = 0 Then gError 86171
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 53313

    For iIndice = 1 To objGridDistribuicao.iLinhasExistentes
        If GridDistribuicao.TextMatrix(iIndice, iGrid_Prod_Col) = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) Then
            GridDistribuicao.TextMatrix(iIndice, iGrid_DescProduto_Col) = DescProduto.Text
        End If
    Next

    Saida_Celula_Descricao = SUCESSO

    Exit Function

Erro_Saida_Celula_Descricao:

    Saida_Celula_Descricao = gErr

    Select Case gErr

        Case 53313
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 86171
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164485)

    End Select

    Exit Function


End Function

'??? Leo daqui p/ baixo
Private Sub GridNotas_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridNotas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridNotas, iAlterado)
    End If

End Sub

Private Sub GridNotas_EnterCell()

    Call Grid_Entrada_Celula(objGridNotas, iAlterado)

End Sub

Private Sub GridNotas_GotFocus()

    Call Grid_Recebe_Foco(objGridNotas)

End Sub

Private Sub GridNotas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridNotas, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridNotas, iAlterado)
    End If

End Sub

Private Sub GridNotas_LeaveCell()

    Call Saida_Celula(objGridNotas)

End Sub

Private Sub GridNotas_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridNotas)

End Sub

Private Sub GridNotas_RowColChange()

    Call Grid_RowColChange(objGridNotas)

End Sub

Private Sub GridNotas_Scroll()

    Call Grid_Scroll(objGridNotas)

End Sub

Private Sub GridNotas_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridNotas)

End Sub

Private Function Inicializa_GridNotas(objGridInt As AdmGrid) As Long
'Executa a Inicialização do gridNotas

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Nota")
    
    ' campos de edição do grid
    objGridInt.colCampo.Add (NotaPC.Name)
    
    'indica onde estao situadas as colunas do grid
    iGrid_NotaPC_Col = 1

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridNotas

    'Linhas do grid
    objGridInt.objGrid.Rows = 20

    GridBloqueios.ColWidth(0) = 300

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 16

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGridInt.iProibidoIncluir = PROIBIDO_INCLUIR
    
    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridNotas = SUCESSO

    Exit Function

End Function

Private Sub NotaPC_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NotaPC_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridNotas)

End Sub

Private Sub NotaPC_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridNotas)

End Sub

Private Sub NotaPC_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridNotas.objControle = NotaPC
    lErro = Grid_Campo_Libera_Foco(objGridNotas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_GridNotas(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridNotas
    
    lErro = Saida_Celula_NotasPC(objGridInt)
    If lErro <> SUCESSO Then gError 103349

    Saida_Celula_GridNotas = SUCESSO

    Exit Function

Erro_Saida_Celula_GridNotas:

    Saida_Celula_GridNotas = gErr

    Select Case gErr

        Case 103349

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164486)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_NotasPC(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_NotasPC

    Set objGridInt.objControle = NotaPC

    'Verifica se a nota foi informada
    If Len(Trim(NotaPC.Text)) > 0 Then
        
        GridNotas.TextMatrix(GridNotas.Row, iGrid_NotaPC_Col) = NotaPC.Text
        
        lErro = Grid_Abandona_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 103350
    
    End If

    Saida_Celula_NotasPC = SUCESSO

    Exit Function

Erro_Saida_Celula_NotasPC:

    Saida_Celula_NotasPC = gErr

    Select Case gErr

        Case 103350
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164487)

    End Select

    Exit Function

End Function

Private Function Move_GridNotas_Memoria(objPedidoCompra As ClassPedidoCompras) As Long
'Recolhe do Grid os dados das notas

Dim iIndice As Integer
Dim sNota As String

    For iIndice = 1 To objGridNotas.iLinhasExistentes

        'Armazena os dados das Notas
        sNota = GridNotas.TextMatrix(iIndice, iGrid_NotaPC_Col)
       
        'Adiciona o item na colecao de NotasPedCompras
        objPedidoCompra.colNotasPedCompras.Add sNota

    Next

    Move_GridNotas_Memoria = SUCESSO

End Function

Private Function Preenche_Grid_Notas(objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Preenche_Grid_Notas

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridNotas)

    iIndice = 0

    For iIndice = 1 To objPedidoCompra.colNotasPedCompras.Count

        GridNotas.TextMatrix(iIndice, iGrid_NotaPC_Col) = objPedidoCompra.colNotasPedCompras.Item(iIndice)
        
    Next

    objGridNotas.iLinhasExistentes = iIndice - 1

    Preenche_Grid_Notas = SUCESSO

    Exit Function

Erro_Preenche_Grid_Notas:

    Preenche_Grid_Notas = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164488)

    End Select

    Exit Function

End Function

'??? por leo
Private Sub BotaoIncluirNota_Click()

    If Len(Trim(Nota.Text)) > 0 Then
    
        objGridNotas.iLinhasExistentes = objGridNotas.iLinhasExistentes + 1
        
        GridNotas.TextMatrix(objGridNotas.iLinhasExistentes, iGrid_NotaPC_Col) = Trim(Nota.Text)
        
        Nota.Text = ""
        
    End If

End Sub

Private Sub objEventoNotaPC_evSelecao(obj1 As Object)

Dim objNotasPC As ClassNotaPC

    Set objNotasPC = obj1

    'Coloca a Nota na tela
    Nota.Text = objNotasPC.sNota

    Me.Show

    Exit Sub

End Sub

'??? Já existe na tela de moedas
Public Function Moedas_Le(objMoedas As ClassMoedas) As Long

Dim lComando As Long
Dim lErro As Long
Dim sNome As String
Dim sSimbolo As String

On Error GoTo Erro_Moedas_Le

    'Abre Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 108818

    'Inicializa as strings
    sNome = String(STRING_NOME_MOEDA, 0)
    sSimbolo = String(STRING_SIMBOLO_MOEDA, 0)
    
    'Verifica se existe moeda com o codigo passado
    lErro = Comando_Executar(lComando, "SELECT Nome, Simbolo FROM Moedas WHERE Codigo = ?", sNome, sSimbolo, objMoedas.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 108819

    'Busca o registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 108820

    'Se nao encontrou => Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 108821

    'Transfere os dados
    objMoedas.sNome = sNome
    objMoedas.sSimbolo = sSimbolo
    
    'Fecha Comando
    Call Comando_Fechar(lComando)

    Moedas_Le = SUCESSO
    
    Exit Function

Exit Function

Erro_Moedas_Le:

    Moedas_Le = gErr

    Select Case gErr

        Case 108818
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 108819, 108820
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOEDAS", gErr)

        Case 108821

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164489)

    End Select

    Call Comando_Fechar(lComando)

End Function

Private Sub BotaoTrazCotacao_Click()
'Traz a última cotação da moeda selecionada

Dim lErro As Long
Dim objCotacao As New ClassCotacaoMoeda
Dim objCotacaoAnterior As New ClassCotacaoMoeda

On Error GoTo Erro_BotaoTrazCotacao_Click

    'Carrega objCotacao
    objCotacao.dtData = gdtDataAtual
    
    'Se a moeda não foi selecionada => Erro
    If Len(Trim(Moeda.Caption)) = 0 Then gError 108943
        
    'Preeche com a Moeda selecionada
    objCotacao.iMoeda = Codigo_Extrai(Moeda.Caption)
    objCotacaoAnterior.iMoeda = Codigo_Extrai(Moeda.Caption)

    'Chama função de leitura
    lErro = CF("CotacaoMoeda_Le_UltimasCotacoes", objCotacao, objCotacaoAnterior)
    If lErro <> SUCESSO Then gError 108944
    
    'Se nao existe cotacao para a data informada => Mostra a última.
    Taxa.Text = IIf(objCotacao.dValor <> 0, Format(objCotacao.dValor, FORMATO_TAXA_CONVERSAO_MOEDA), Format(objCotacaoAnterior.dValor, FORMATO_TAXA_CONVERSAO_MOEDA))

    Exit Sub
    
Erro_BotaoTrazCotacao_Click:

    Select Case gErr
    
        Case 108943
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDA_NAO_PREENCHIDA", gErr)
            '??? Falta cadastrar: ERRO_MOEDA_NAO_PREENCHIDA - "Para trazer a cotação a moeda deve ser selecionada antes."
            
        Case 108944
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164490)
    
    End Select
    
End Sub

Private Sub ComparativoMoedaReal_Calcula(ByVal dTaxa As Double)
'Preenche as colunas INFORMATIVAS de proporção da moeda R$.

Dim iIndice As Integer

On Error GoTo Erro_ComparativoMoedaReal_Calcula

    'Para cada linha do grid de Itens será claculado o correspondente em R$
    For iIndice = 1 To objGridItens.iLinhasExistentes
        
        'Preço Unitário em R$ = Preço Unitário na Moeda selecionada dividido pela taxa de conversão
        GridItens.TextMatrix(iIndice, iGrid_PrecoUnitarioMoedaReal_Col) = Format(StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col)) * dTaxa, gobjCOM.sFormatoPrecoUnitario) 'Alterado por Wagner
        
        'Preço Total em R$ = Preço Unitário em R$ x Quantidade do produto
        GridItens.TextMatrix(iIndice, iGrid_TotalMoedaReal_Col) = Format(StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoUnitarioMoedaReal_Col)) * StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col)), TotalMoedaReal.Format) 'Alterado por Wagner
        
    Next

    Exit Sub
    
Erro_ComparativoMoedaReal_Calcula:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164491)

    End Select

End Sub

'##############################################
'Inserido por Wagner
Private Sub Formata_Controles()

    PrecoUnitario.Format = gobjCOM.sFormatoPrecoUnitario
    PrecoUnitarioMoedaReal.Format = gobjCOM.sFormatoPrecoUnitario

End Sub
'##############################################

'##############################################
'Inserido por Wagner 13/09/05
Private Function PedidoCompra_Imprimir(objPedidoCompra) As Long

Dim lErro As Long, sFiltro As String
Dim objRelatorio As New AdmRelatorio
Dim objBloqueioPC As ClassBloqueioPC

On Error GoTo Erro_PedidoCompra_Imprimir

    If objPedidoCompra.lCodigo = 0 Then gError 76054

    'Verifica se o Pedido de Compra informado existe
    lErro = CF("PedidoCompra_Le_Numero", objPedidoCompra)
    If lErro <> SUCESSO And lErro <> 56142 Then gError 76027
    
    'Se o Pedido não existe ==> erro
    If lErro = 56142 Then gError 76028
        
    lErro = CF("BloqueiosPC_Le", objPedidoCompra)
    If lErro <> SUCESSO Then gError 76056
    
    'If objPedidoCompra.colBloqueiosPC.Count > 0 Then gError 76050
    
    For Each objBloqueioPC In objPedidoCompra.colBloqueiosPC
            
        If objBloqueioPC.dtDataLib = DATA_NULA Then gError 76050
    
    Next
    
    'Preenche a Data de Entrada com a Data Atual
    DataEmissao.Caption = Format(gdtDataHoje, "dd/mm/yy")

    objPedidoCompra.dtDataEmissao = gdtDataHoje

    'Atualiza data de emissao no BD para a data atual
    lErro = CF("PedidoCompra_Atualiza_DataEmissao", objPedidoCompra)
    If lErro <> SUCESSO And lErro <> 56348 Then gError 53306

    'se nao encontrar ---> erro
    If lErro = 56348 Then gError 53307
    
    'Alterado por Wagner
    sFiltro = "REL_PCOM.PC_NumIntDoc = @NPEDCOM"
    lErro = CF("Relatorio_ObterFiltro", "Pedido de Compra Gerado", sFiltro)
    If lErro <> SUCESSO Then gError 76029
    
    'Executa o relatório
    lErro = objRelatorio.ExecutarDireto("Pedido de Compra Gerado", sFiltro, 0, "PEDCOM", "NPEDCOM", objPedidoCompra.lNumIntDoc)
    If lErro <> SUCESSO Then gError 76029
    
    PedidoCompra_Imprimir = SUCESSO

    Exit Function

Erro_PedidoCompra_Imprimir:

    PedidoCompra_Imprimir = gErr

    Select Case gErr
    
        Case 53306, 76056

        Case 53307
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", gErr, objPedidoCompra.lCodigo)

        Case 76021
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
            
        Case 76027, 76029
        
        Case 76028
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", gErr, objPedidoCompra.lCodigo)
        
        Case 76050
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_BLOQUEADO", gErr, objPedidoCompra.lCodigo)
            
        Case 76054
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDCOMPRA_IMPRESSAO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164492)

    End Select

    Exit Function

End Function
'##############################################

Private Sub DataRefFluxo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataRefFluxo, iAlterado)
    
End Sub

Private Sub DataRefFluxo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataRefFluxo_Validate

    'Verifica se a data foi preenchida
    If Len(DataRefFluxo.ClipText) > 0 Then

        'Critica a data informada
        lErro = Data_Critica(DataRefFluxo.Text)
        If lErro <> SUCESSO Then gError 188595

        If Len(Trim(Data.Caption)) = 0 Then
            If (CDate(DataRefFluxo.Text) < DATA_NULA) Then gError 188596
        Else
            If (CDate(DataRefFluxo.Text) < CDate(Data.Caption)) Then gError 188597
        End If

    End If

    Exit Sub

Erro_DataRefFluxo_Validate:

    Cancel = True

    Select Case gErr

        Case 188595

        Case 188596, 188597
            Call Rotina_Erro(vbOKOnly, "ERRO_DataRefFluxo_INFERIOR_DATAPEDIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188598)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataRefFluxo_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataRefFluxo_DownClick

    'Diminui um dia em DataRefFluxo
    lErro = Data_Up_Down_Click(DataRefFluxo, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 188596

    Exit Sub

Erro_UpDownDataRefFluxo_DownClick:

    Select Case gErr

        Case 188596

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188597)

    End Select

    Exit Sub


End Sub

Private Sub UpDownDataRefFluxo_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataRefFluxo_UpClick

    'Aumenta um dia em DataRefFluxo
    lErro = Data_Up_Down_Click(DataRefFluxo, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 188598

    Exit Sub

Erro_UpDownDataRefFluxo_UpClick:

    Select Case gErr

        Case 188598

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188599)

    End Select

    Exit Sub

End Sub

