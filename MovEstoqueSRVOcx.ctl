VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl MovEstoqueSRVOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5310
      Index           =   1
      Left            =   195
      TabIndex        =   19
      Top             =   645
      Width           =   9225
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
         Height          =   315
         Left            =   0
         TabIndex        =   11
         Top             =   4965
         Width           =   1110
      End
      Begin VB.CommandButton BotaoEstoque 
         Caption         =   "Estoque"
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
         Left            =   1277
         TabIndex        =   12
         Top             =   4965
         Width           =   1035
      End
      Begin VB.CommandButton BotaoSerie 
         Caption         =   "Núm.Séries ..."
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
         Left            =   7515
         TabIndex        =   18
         Top             =   4965
         Width           =   1545
      End
      Begin VB.CommandButton BotaoLote 
         Caption         =   "Lotes"
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
         Left            =   4523
         TabIndex        =   15
         Top             =   4965
         Width           =   840
      End
      Begin VB.CommandButton BotaoOP 
         Caption         =   "OS"
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
         Left            =   2479
         TabIndex        =   13
         Top             =   4965
         Width           =   795
      End
      Begin VB.CommandButton BotaoCcls 
         Caption         =   "Ccl"
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
         Left            =   5530
         TabIndex        =   16
         Top             =   4965
         Width           =   810
      End
      Begin VB.CommandButton BotaoPlanoConta 
         Caption         =   "Contas"
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
         Left            =   6507
         TabIndex        =   17
         Top             =   4965
         Width           =   840
      End
      Begin VB.CommandButton BotaoServicos 
         Caption         =   "Serviços"
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
         Left            =   3441
         TabIndex        =   14
         Top             =   4965
         Width           =   915
      End
      Begin VB.Frame Frame4 
         Caption         =   "Outros"
         Height          =   825
         Left            =   45
         TabIndex        =   97
         Top             =   1200
         Width           =   9105
         Begin VB.TextBox Observacao 
            Height          =   585
            Left            =   765
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   9
            Top             =   165
            Width           =   8220
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "OBS:"
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
            TabIndex        =   98
            Top             =   225
            Width           =   450
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Padrões"
         Height          =   540
         Left            =   45
         TabIndex        =   92
         Top             =   645
         Width           =   9105
         Begin VB.TextBox OP 
            Height          =   285
            Left            =   780
            MaxLength       =   9
            TabIndex        =   5
            Top             =   195
            Width           =   1230
         End
         Begin MSMask.MaskEdBox CclPadrao 
            Height          =   315
            Left            =   4800
            TabIndex        =   7
            Top             =   165
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox AlmoxPadrao 
            Height          =   315
            Left            =   6930
            TabIndex        =   8
            Top             =   165
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdutoOPGera 
            Height          =   285
            Left            =   2820
            TabIndex        =   6
            Top             =   195
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label ProdutoOPLabel 
            AutoSize        =   -1  'True
            Caption         =   "Serviço:"
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
            Left            =   2070
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   96
            Top             =   225
            Width           =   705
         End
         Begin VB.Label OPLabel 
            AutoSize        =   -1  'True
            Caption         =   "OS:"
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
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   95
            Top             =   240
            Width           =   330
         End
         Begin VB.Label AlmoxPadraoLabel 
            AutoSize        =   -1  'True
            Caption         =   "Almox.:"
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
            Left            =   6270
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   94
            Top             =   195
            Width           =   630
         End
         Begin VB.Label CclPadraoLabel 
            AutoSize        =   -1  'True
            Caption         =   "Ccl:"
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
            Left            =   4440
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   93
            Top             =   210
            Width           =   345
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Identificação"
         Height          =   615
         Left            =   30
         TabIndex        =   88
         Top             =   0
         Width           =   9150
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   1635
            Picture         =   "MovEstoqueSRVOcx.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Numeração Automática"
            Top             =   240
            Width           =   300
         End
         Begin MSComCtl2.UpDown UpDownData 
            Height          =   300
            Left            =   3915
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   225
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox Data 
            Height          =   315
            Left            =   2835
            TabIndex        =   2
            Top             =   225
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   315
            Left            =   795
            TabIndex        =   0
            Top             =   225
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Hora 
            Height          =   300
            Left            =   4815
            TabIndex        =   4
            Top             =   225
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   " "
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Data:"
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
            Left            =   2325
            TabIndex        =   91
            Top             =   270
            Width           =   480
         End
         Begin VB.Label CodigoLabel 
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
            Left            =   120
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   90
            Top             =   270
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Hora:"
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
            Left            =   4335
            TabIndex        =   89
            Top             =   270
            Width           =   480
         End
      End
      Begin VB.ComboBox FilialOP 
         Height          =   315
         Left            =   1755
         TabIndex        =   33
         Top             =   3090
         Width           =   1575
      End
      Begin MSMask.MaskEdBox Lote 
         Height          =   270
         Left            =   2520
         TabIndex        =   32
         Top             =   2520
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.TextBox Descricao 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   5400
         MaxLength       =   50
         TabIndex        =   27
         Top             =   2700
         Width           =   2340
      End
      Begin VB.ComboBox UnidadeMed 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   2085
         Width           =   645
      End
      Begin VB.ComboBox TipoMov 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "MovEstoqueSRVOcx.ctx":00EA
         Left            =   4275
         List            =   "MovEstoqueSRVOcx.ctx":00EC
         TabIndex        =   23
         Top             =   2130
         Width           =   2835
      End
      Begin VB.CheckBox Estorno 
         Height          =   210
         Left            =   7650
         TabIndex        =   29
         Top             =   2955
         Width           =   870
      End
      Begin MSMask.MaskEdBox ContaContabilDebito 
         Height          =   240
         Left            =   6615
         TabIndex        =   31
         Top             =   3465
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
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
      Begin MSMask.MaskEdBox ContaContabilCredito 
         Height          =   240
         Left            =   4845
         TabIndex        =   30
         Top             =   3435
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
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
      Begin MSMask.MaskEdBox OPCodigo 
         Height          =   225
         Left            =   7260
         TabIndex        =   26
         Top             =   2430
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
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
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ProdutoOP 
         Height          =   225
         Left            =   5640
         TabIndex        =   28
         Top             =   2925
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   20
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
      Begin MSMask.MaskEdBox Quantidade 
         Height          =   225
         Left            =   3225
         TabIndex        =   22
         Top             =   2130
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
         Left            =   360
         TabIndex        =   20
         Top             =   2130
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
         Left            =   7440
         TabIndex        =   24
         Top             =   2070
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Ccl 
         Height          =   225
         Left            =   4545
         TabIndex        =   25
         Top             =   2565
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   10
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
      Begin MSFlexGridLib.MSFlexGrid GridMovimentos 
         Height          =   2040
         Left            =   15
         TabIndex        =   10
         Top             =   2265
         Width           =   9210
         _ExtentX        =   16245
         _ExtentY        =   3598
         _Version        =   393216
         Rows            =   11
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
         AllowUserResizing=   1
      End
      Begin MSMask.MaskEdBox Custo 
         Height          =   225
         Left            =   675
         TabIndex        =   87
         Top             =   2640
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
         Format          =   "#,##0.00##"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Movimentos de Estoque"
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
         Left            =   75
         TabIndex        =   62
         Top             =   2025
         Width           =   2040
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade Disponível:"
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
         Left            =   120
         TabIndex        =   63
         Top             =   4665
         Width           =   2025
      End
      Begin VB.Label QuantDisponivel 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2205
         TabIndex        =   64
         Top             =   4620
         Width           =   1545
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5160
      Index           =   2
      Left            =   180
      TabIndex        =   34
      Top             =   750
      Visible         =   0   'False
      Width           =   9225
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4680
         TabIndex        =   86
         Tag             =   "1"
         Top             =   2520
         Width           =   870
      End
      Begin VB.CommandButton CTBBotaoModeloPadrao 
         Caption         =   "Modelo Padrão"
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
         Left            =   6420
         TabIndex        =   40
         Top             =   345
         Width           =   2700
      End
      Begin VB.CommandButton CTBBotaoLimparGrid 
         Caption         =   "Limpar Grid"
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
         Left            =   6420
         TabIndex        =   38
         Top             =   30
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   6450
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   870
         Width           =   2700
      End
      Begin VB.CommandButton CTBBotaoImprimir 
         Caption         =   "Imprimir"
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
         Left            =   7875
         TabIndex        =   39
         Top             =   30
         Width           =   1245
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4920
         TabIndex        =   48
         Top             =   1320
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin VB.CheckBox CTBAglutina 
         Height          =   210
         Left            =   4470
         TabIndex        =   50
         Top             =   2010
         Width           =   870
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   49
         Top             =   1620
         Width           =   1770
      End
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2790
         Left            =   6330
         TabIndex        =   52
         Top             =   1485
         Visible         =   0   'False
         Width           =   2685
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   59
         Top             =   3420
         Width           =   5895
         Begin VB.Label CTBCclLabel 
            AutoSize        =   -1  'True
            Caption         =   "Centro de Custo:"
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
            Left            =   240
            TabIndex        =   65
            Top             =   660
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.Label CTBLabel7 
            AutoSize        =   -1  'True
            Caption         =   "Conta:"
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
            Left            =   1125
            TabIndex        =   66
            Top             =   300
            Width           =   570
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   67
            Top             =   285
            Width           =   3720
         End
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   68
            Top             =   645
            Visible         =   0   'False
            Width           =   3720
         End
      End
      Begin VB.CheckBox CTBLancAutomatico 
         Caption         =   "Recalcula Automaticamente"
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
         Left            =   3465
         TabIndex        =   43
         Top             =   945
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin MSMask.MaskEdBox CTBConta 
         Height          =   225
         Left            =   525
         TabIndex        =   44
         Top             =   1305
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBDebito 
         Height          =   225
         Left            =   3435
         TabIndex        =   47
         Top             =   1335
         Width           =   1155
         _ExtentX        =   2037
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
      Begin MSMask.MaskEdBox CTBCredito 
         Height          =   225
         Left            =   2280
         TabIndex        =   46
         Top             =   1275
         Width           =   1155
         _ExtentX        =   2037
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
      Begin MSMask.MaskEdBox CTBCcl 
         Height          =   225
         Left            =   1545
         TabIndex        =   45
         Top             =   1320
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   10
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
      Begin MSComCtl2.UpDown CTBUpDown 
         Height          =   300
         Left            =   1650
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   510
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox CTBDataContabil 
         Height          =   300
         Left            =   570
         TabIndex        =   37
         Top             =   525
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBLote 
         Height          =   300
         Left            =   5580
         TabIndex        =   36
         Top             =   135
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBDocumento 
         Height          =   300
         Left            =   3795
         TabIndex        =   35
         Top             =   120
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid CTBGridContabil 
         Height          =   1860
         Left            =   0
         TabIndex        =   51
         Top             =   1170
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   3281
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin MSComctlLib.TreeView CTBTvwCcls 
         Height          =   2985
         Left            =   6330
         TabIndex        =   53
         Top             =   1485
         Visible         =   0   'False
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   5265
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView CTBTvwContas 
         Height          =   2985
         Left            =   6330
         TabIndex        =   54
         Top             =   1485
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   5265
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin VB.Label CTBLabel1 
         AutoSize        =   -1  'True
         Caption         =   "Modelo:"
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
         Left            =   6480
         TabIndex        =   41
         Top             =   660
         Width           =   690
      End
      Begin VB.Label CTBLabel21 
         Caption         =   "Origem:"
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
         Height          =   255
         Left            =   45
         TabIndex        =   69
         Top             =   165
         Width           =   720
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   70
         Top             =   120
         Width           =   1530
      End
      Begin VB.Label CTBLabel14 
         Caption         =   "Período:"
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
         Left            =   4230
         TabIndex        =   71
         Top             =   600
         Width           =   735
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5010
         TabIndex        =   72
         Top             =   570
         Width           =   1185
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   73
         Top             =   555
         Width           =   1185
      End
      Begin VB.Label CTBLabel13 
         Caption         =   "Exercício:"
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
         Left            =   1995
         TabIndex        =   74
         Top             =   585
         Width           =   870
      End
      Begin VB.Label CTBLabel5 
         AutoSize        =   -1  'True
         Caption         =   "Lançamentos"
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
         TabIndex        =   75
         Top             =   945
         Width           =   1140
      End
      Begin VB.Label CTBLabelHistoricos 
         Caption         =   "Históricos"
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
         Left            =   6345
         TabIndex        =   76
         Top             =   1275
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label CTBLabelContas 
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
         Height          =   255
         Left            =   6345
         TabIndex        =   77
         Top             =   1275
         Width           =   2340
      End
      Begin VB.Label CTBLabelCcl 
         Caption         =   "Centros de Custo / Lucro"
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
         Left            =   6345
         TabIndex        =   78
         Top             =   1275
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.Label CTBLabelTotais 
         Caption         =   "Totais:"
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
         Height          =   225
         Left            =   1800
         TabIndex        =   79
         Top             =   3045
         Width           =   615
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   80
         Top             =   3030
         Width           =   1155
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2460
         TabIndex        =   81
         Top             =   3030
         Width           =   1155
      End
      Begin VB.Label CTBLabel8 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
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
         TabIndex        =   82
         Top             =   555
         Width           =   480
      End
      Begin VB.Label CTBLabelDoc 
         AutoSize        =   -1  'True
         Caption         =   "Documento:"
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
         Left            =   2700
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   83
         Top             =   165
         Width           =   1035
      End
      Begin VB.Label CTBLabelLote 
         AutoSize        =   -1  'True
         Caption         =   "Lote:"
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
         Left            =   5100
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   84
         Top             =   165
         Width           =   450
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7275
      ScaleHeight     =   495
      ScaleWidth      =   2100
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   30
      Width           =   2160
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "MovEstoqueSRVOcx.ctx":00EE
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "MovEstoqueSRVOcx.ctx":0278
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "MovEstoqueSRVOcx.ctx":03D2
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "MovEstoqueSRVOcx.ctx":0904
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5670
      Left            =   120
      TabIndex        =   61
      Top             =   315
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   10001
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Movimentos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Contabilização"
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
Attribute VB_Name = "MovEstoqueSRVOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Public gobjAnotacao As ClassAnotacoes

'inicio contabilidade

Dim objGrid1 As AdmGrid
Dim objContabil As New ClassContabil

Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1
Private WithEvents objEventoRastroLote As AdmEvento
Attribute objEventoRastroLote.VB_VarHelpID = -1

'Mnemônicos
Private Const CODIGO1 As String = "Codigo"
Private Const DATA1 As String = "Data"
Private Const CCL1 As String = "Ccl"
Private Const Produto1 As String = "Produto_Codigo"
Private Const ALMOXARIFADO1 As String = "Almoxarifado"
Private Const ESTORNO1 As String = "Estorno"
Private Const UNIDADE_MED As String = "Unidade_Med"
Private Const QUANTIDADE1 As String = "Quantidade"
Private Const DESCRICAO1 As String = "Descricao_Item"
Private Const OP_CODIGO As String = "OS_Codigo"
Private Const PRODUTO_OP As String = "Servico_OS"
Private Const TIPO_MOV As String = "Tipo_Mov"
Private Const CONTACONTABILDEBITO1 As String = "ContaContabilDebito"
Private Const CONTACONTABILCREDITO1 As String = "ContaContabilCredito"
Private Const QUANT_ESTOQUE As String = "Quant_Estoque"
'fim da contabilidade

Dim gcolcolRastreamentoSerie As Collection
Dim iTipoMovtoAnt As Integer

'Declaração de Variáveis Globais
Public iAlterado As Integer
Dim iFrameAtual As Integer
Dim lCodigoAntigo As Long

Dim colTipoMovEst As colTipoMovEst
Dim colEntradaOuSaida As AdmColCodigoNome
Dim colItensNumIntDoc As Collection

Dim objGrid As AdmGrid
Dim iLinhaAntiga As Integer
Dim iGrid_Sequencial_Col As Integer
Dim iGrid_Produto_Col As Integer
Dim iGrid_Descricao_Col As Integer
Dim iGrid_UnidadeMed_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_TipoMov_Col As Integer
Dim iGrid_Almoxarifado_Col As Integer
Dim iGrid_Ccl_Col As Integer
Dim iGrid_OP_Col As Integer
Dim iGrid_FilialOP_Col As Integer
Dim iGrid_ProdutoOP_Col As Integer
Dim iGrid_Estorno_Col As Integer
Dim iGrid_ContaContabilCredito_Col As Integer
Dim iGrid_ContaContabilDebito_Col As Integer
Dim iGrid_Lote_Col As Integer
Dim iGrid_Custo_Col As Integer

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoCclPadrao As AdmEvento
Attribute objEventoCclPadrao.VB_VarHelpID = -1
Private WithEvents objEventoAlmoxPadrao As AdmEvento
Attribute objEventoAlmoxPadrao.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoEstoque As AdmEvento
Attribute objEventoEstoque.VB_VarHelpID = -1
Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1
Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1
Private WithEvents objEventoOP As AdmEvento
Attribute objEventoOP.VB_VarHelpID = -1
Private WithEvents objEventoProdutoOP_Grid As AdmEvento
Attribute objEventoProdutoOP_Grid.VB_VarHelpID = -1
Private WithEvents objEventoOP_OP As AdmEvento
Attribute objEventoOP_OP.VB_VarHelpID = -1
Private WithEvents objEventoProdutoOP_OP As AdmEvento
Attribute objEventoProdutoOP_OP.VB_VarHelpID = -1
Private WithEvents objEventoOP_OP_Todos As AdmEvento
Attribute objEventoOP_OP_Todos.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_Movimentos = 1
Private Const TAB_Contabilizacao = 2

Private Sub BotaoExcluir_Click()

Dim objMovEstoque As New ClassMovEstoque
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'Verifica se o código foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 89874

    objMovEstoque.lCodigo = CLng(Codigo.Text)
    objMovEstoque.iFilialEmpresa = giFilialEmpresa
    
    'trata a exclusão do movimento de estoque
    lErro = CF("MovEstoque_Exclui", objMovEstoque, objContabil)
    If lErro <> SUCESSO Then gError 89890

    Call Limpa_Tela_MovEstoque

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 89874
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 89890
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162800)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    lErro = CF("MovEstoque_Automatico", giFilialEmpresa, lCodigo)
    If lErro <> SUCESSO Then gError 57521

    Codigo.PromptInclude = False
    Codigo.Text = CStr(lCodigo)
    Codigo.PromptInclude = True

    Codigo.SetFocus
    
    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 57521
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162801)
    
    End Select

    Exit Sub

End Sub

Private Sub Almoxarifado_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub AlmoxPadrao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub AlmoxPadrao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_AlmoxPadrao_Validate

    'Verifica se Almoxarifado Padrão foi informado
    If Len(Trim(AlmoxPadrao.Text)) <> 0 Then

        lErro = TP_Almoxarifado_Filial_Le(AlmoxPadrao, objAlmoxarifado, 0)
        If lErro <> SUCESSO And lErro <> 25136 And lErro <> 25143 Then gError 30263

        If lErro = 25136 Then gError 22916

        If lErro = 25143 Then gError 22917

    End If

    Exit Sub

Erro_AlmoxPadrao_Validate:

    Cancel = True
    
    Select Case gErr

        Case 22916, 22917
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, AlmoxPadrao.Text)
    
        Case 30263
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162802)

    End Select

    Exit Sub

End Sub

Private Sub AlmoxPadraoLabel_Click()

Dim colSelecao As New Collection
Dim objAlmoxarifado As ClassAlmoxarifado
    
    Call Chama_Tela("AlmoxarifadoLista_Consulta", colSelecao, objAlmoxarifado, objEventoAlmoxPadrao)

End Sub

Private Sub BotaoCcls_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCcl As ClassCcl

On Error GoTo Erro_BotaoCcls_Click

    If GridMovimentos.Row = 0 Then gError 43764

    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))) = 0 Then gError 43765
    
    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCcl)

    Exit Sub
    
Erro_BotaoCcls_Click:
    
    Select Case gErr
    
        Case 43764
             Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 43765
             Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162803)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoEstoque_Click()

Dim lErro As Long
Dim objEstoqueProduto As ClassEstoqueProduto
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoEstoque_Click

    If GridMovimentos.Row = 0 Then gError 43712

    sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 22919

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        colSelecao.Add sProdutoFormatado

        Call Chama_Tela("EstoqueProdutoFilialLista", colSelecao, objEstoqueProduto, objEventoEstoque)
    Else
        Error 43736
    End If

    Exit Sub

Erro_BotaoEstoque_Click:

    Select Case gErr

        Case 22919
        
        Case 43712
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 43736
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, , Error$, 162804)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 30334

    Call Limpa_Tela_MovEstoque

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 30334

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162805)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 30346

    Call Limpa_Tela_MovEstoque

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 30346

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162806)

    End Select

    Exit Sub

End Sub

Private Sub BotaoOP_Click()

Dim objOrdemProducao As New ClassOS
Dim colSelecao As New Collection
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_BotaoOP_Click
    
    If (GridMovimentos.Row = 0) Then gError 52245

    sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 52246

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 52247
    
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_OP_Col))) > 0 Then
        objOrdemProducao.sCodigo = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_OP_Col)
    Else
        objOrdemProducao.sCodigo = ""
    End If
    
    Call Chama_Tela("OSLista", colSelecao, objOrdemProducao, objEventoOP)
   
    Exit Sub

Erro_BotaoOP_Click:

    Select Case gErr
        
        Case 52245
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 52246
        
        Case 52247
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162807)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)

End Sub

Private Sub Lote_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Lote_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Lote_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Lote_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Lote
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub FilialOP_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub FilialOP_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub FilialOP_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub FilialOP_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = FilialOP
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub BotaoPlanoConta_Click()

Dim lErro As Long
Dim iContaPreenchida As Integer
Dim sConta As String
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoPlanoConta_Click

    If GridMovimentos.Row = 0 Then gError 43751
    
    If GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col) = "" Then gError 43752

    sConta = String(STRING_CONTA, 0)
    
    If GridMovimentos.Col = iGrid_ContaContabilCredito_Col Then
        
        lErro = CF("Conta_Formata", ContaContabilCredito.Text, sConta, iContaPreenchida)
        If lErro <> SUCESSO Then gError 43753
    
    ElseIf GridMovimentos.Col = iGrid_ContaContabilDebito_Col Then
        
        lErro = CF("Conta_Formata", ContaContabilDebito.Text, sConta, iContaPreenchida)
        If lErro <> SUCESSO Then gError 43763
    
    End If
    
    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    'Chama PlanoContaESTLista
    Call Chama_Tela("PlanoContaESTLista", colSelecao, objPlanoConta, objEventoContaContabil)
    
    Exit Sub

Erro_BotaoPlanoConta_Click:

    Select Case gErr

        Case 43751
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 43752
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 43753, 43763

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162808)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProdutos_Click()

Dim objProduto As New ClassProduto
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long
Dim objItemOP As New ClassItemOS
Dim sSelecao As String

On Error GoTo Erro_BotaoProdutos_Click
    
    'se não selecionou nenhuma linha do grid entao - - - -> Erro
    If GridMovimentos.Row = 0 Then gError 43711
              
    ' Formata o Produto
    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 52255
        
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then objProduto.sCodigo = sProdutoFormatado
    
    sSelecao = "ControleEstoque<>?"
    colSelecao.Add PRODUTO_CONTROLE_SEM_ESTOQUE
    
   'Lista de produtos que  podem participar da producao e inventariados
   Call Chama_Tela("ProdutoEstoqueLista", colSelecao, objProduto, objEventoProduto, sSelecao)
   
   Exit Sub
   
Erro_BotaoProdutos_Click:

     Select Case gErr

        Case 43711
             Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 52254, 52255
        
        Case 55388
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTOOP_NAO_PREENCHIDO_OP", gErr)
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162809)

     End Select

    Exit Sub

End Sub

Private Sub Ccl_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CclPadrao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub BotaoServicos_Click()

Dim colSelecao As New Collection
Dim lErro As Long
Dim objItemOP As New ClassOS
Dim sOS As String
    
On Error GoTo Erro_BotaoServicos_Click

    If GridMovimentos.Row = 0 Then gError 206251
    
    sOS = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_OP_Col)
        
    'se a OP estiver preenchida, mostra só os produtos da OP em questão
    If Len(Trim(sOS)) = 0 Then gError 206254
        
    colSelecao.Add Trim(sOS)
    
    Call Chama_Tela("ItemOSLista", colSelecao, objItemOP, objEventoProdutoOP_Grid, "Codigo = ?")
        
    Exit Sub
   
Erro_BotaoServicos_Click:

    Select Case gErr
              
        Case 206251 To 206253
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 206254
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTOOP_NAO_PREENCHIDO_OP", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206255)
        
    End Select
    
    Exit Sub
   
End Sub

Private Sub CclPadrao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sCclFormatada As String
Dim objCcl As New ClassCcl

On Error GoTo Erro_CclPadrao_Validate

    'Verifica se o CclPadrao foi Preenchida
    If Len(Trim(CclPadrao.Text)) <> 0 Then

        lErro = CF("Ccl_Critica", CclPadrao.Text, sCclFormatada, objCcl)
        If lErro <> SUCESSO And lErro <> 5703 Then gError 30261

        If lErro = 5703 Then gError 30262

    End If

    Exit Sub

Erro_CclPadrao_Validate:

    Cancel = True


    Select Case gErr

        Case 30261

        Case 30262
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, CclPadrao.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162810)

    End Select

    Exit Sub

End Sub

Private Sub CclPadraoLabel_Click()

Dim objCcl As ClassCcl
Dim colSelecao As New Collection

    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCclPadrao)

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Limpa_colItensNumIntDoc(colItensNumIntDoc As Collection)

Dim iCount As Integer
Dim iIndice As Integer

    iCount = colItensNumIntDoc.Count
    Set colItensNumIntDoc = New Collection

    For iIndice = 0 To iCount - 1

        colItensNumIntDoc.Add 0
        'GridMovimentos.TextMatrix(iIndice, iGrid_Estorno_Col) = "0"

    Next

End Sub

Private Function Verifica_TipoTelaAtual(iTipoMov As Integer) As Long

Dim iIndice As Integer
Dim iAchou As Integer

    iAchou = 0
    For iIndice = 0 To TipoMov.ListCount - 1
        If TipoMov.ItemData(iIndice) = iTipoMov Then
            iAchou = 1
            Exit For
        End If
    Next

    Verifica_TipoTelaAtual = iAchou

End Function

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long, iIndice As Integer
Dim objMovEstoque As New ClassMovEstoque
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Codigo_Validate

    If Len(Trim(Codigo.Text)) > 0 Then
        
        lErro = Valor_Positivo_Critica(Codigo.Text)
        If lErro <> SUCESSO Then gError 57763

    End If
    
    'se o codigo foi trocado
    If lCodigoAntigo <> StrParaLong(Trim(Codigo.Text)) Then
    
        If Len(Trim(Codigo.ClipText)) > 0 Then

            Call Limpa_colItensNumIntDoc(colItensNumIntDoc)
    
            objMovEstoque.lCodigo = Codigo.Text
    
            'Le o Movimento de Estoque e Verifica se ele já foi estornado
            lErro = CF("MovEstoqueItens_Le_Verifica_Estorno", objMovEstoque)
            If lErro <> SUCESSO And lErro <> 78883 And lErro <> 78885 Then gError 34904
            
            'Se todos os Itens do Movimento foram estornados
            If lErro = 78885 Then gError 78888
            
            If lErro = SUCESSO Then
    
                If Verifica_TipoTelaAtual(objMovEstoque.iTipoMov) <> MOVESTOQUEINTSRV Then gError 41000
    
                vbMsg = Rotina_Aviso(vbYesNo, "AVISO_PREENCHER_TELA")
    
                If vbMsg = vbNo Then gError 34905
    
                lErro = Preenche_Tela(objMovEstoque)
                If lErro <> SUCESSO Then gError 34906
    
            End If

        End If

        lCodigoAntigo = StrParaLong(Trim(Codigo.Text))

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 34904, 34906, 57763

        Case 34905

        Case 41000
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_INCOMPATIVEL_MOVESTOQUE", gErr, objMovEstoque.lCodigo)
            lCodigoAntigo = 0
        
        Case 78888
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOESTOQUE_ESTORNADO", gErr, giFilialEmpresa, objMovEstoque.lCodigo)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162811)

    End Select

    Exit Sub

End Sub

Private Sub CodigoLabel_Click()

Dim objMovEstoque As New ClassMovEstoque
Dim colSelecao As New Collection

    'Preenche objeto com Código de Tela
    If Len(Trim(Codigo.Text)) <> 0 Then objMovEstoque.lCodigo = CLng(Codigo.Text)

    Call Chama_Tela("MovEstoqueInterno2Lista", colSelecao, objMovEstoque, objEventoCodigo)

End Sub

Private Sub Data_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Estorno_Click()

Dim lErro As Long

On Error GoTo Erro_Estorno_Click

    iAlterado = REGISTRO_ALTERADO

    'Carrega as séries na coleção global
    lErro = Carrega_Series(gcolcolRastreamentoSerie.Item(GridMovimentos.Row), StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), GridMovimentos.Row)
    If lErro <> SUCESSO Then gError 177299

    Exit Sub
    
Erro_Estorno_Click:

    Select Case gErr
    
        Case 177299
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 177300)
    
    End Select
    
    Exit Sub
    
End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim sMascaraCclPadrao As String
Dim objTipoMovEst As ClassTipoMovEst

On Error GoTo Erro_From_Load

    iFrameAtual = 1

    Set gcolcolRastreamentoSerie = New Collection
    Set colTipoMovEst = New colTipoMovEst
    Set colItensNumIntDoc = New Collection
    Set colEntradaOuSaida = New AdmColCodigoNome

    Set objEventoCodigo = New AdmEvento
    Set objEventoCclPadrao = New AdmEvento
    Set objEventoAlmoxPadrao = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoEstoque = New AdmEvento
    Set objEventoCcl = New AdmEvento
    Set objEventoContaContabil = New AdmEvento
    Set objEventoOP = New AdmEvento
    Set objEventoProdutoOP_Grid = New AdmEvento
    Set objEventoRastroLote = New AdmEvento
    Set objEventoOP_OP = New AdmEvento
    Set objEventoProdutoOP_OP = New AdmEvento
    Set objEventoOP_OP_Todos = New AdmEvento
    
    'Lê os tipos de Movimento de Estoque Interno
    lErro = CF("TiposMovEstoque_Le", colTipoMovEst)
    If lErro <> SUCESSO And lErro <> 30352 Then gError 30242

    'Preenche TipoMov com Codigo e Descrição de colTipoMovEst
    'que têm ValidoMovInt=2 e Inativo=ATIVO
    For Each objTipoMovEst In colTipoMovEst

        If objTipoMovEst.iValidoMovInt = MOVESTOQUEINTSRV And objTipoMovEst.iInativo = Ativo Then
                               
            TipoMov.AddItem CStr(objTipoMovEst.iCodigo) & "-" & objTipoMovEst.sDescricao
            
            'Coloca o código em ItemData
            TipoMov.ItemData(TipoMov.NewIndex) = objTipoMovEst.iCodigo
            
            colEntradaOuSaida.Add objTipoMovEst.iCodigo, objTipoMovEst.sEntradaOuSaida
            
        End If

    Next

    'Coloca a Data Atual na Tela
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True

    'Carrega a combo de Filial O.P.
    lErro = Carrega_FilialOP()
    If lErro <> SUCESSO Then gError 78691
    
    'Inicializa Máscara de Produto e ProdutoOP
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 30243

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoOP)
    If lErro <> SUCESSO Then gError 30244

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoOPGera)
    If lErro <> SUCESSO Then gError 30244

    'Inicializa mascara de ContaContabilCredito e ContaContabilDebito
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabilCredito)
    If lErro <> SUCESSO Then gError 49624

    'Inicializa mascara de ContaContabilCredito e ContaContabilDebito
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabilDebito)
    If lErro <> SUCESSO Then gError 49625

    'Inicializa Máscara para CclPadrao e Ccl
    sMascaraCclPadrao = String(STRING_CCL, 0)

    lErro = MascaraCcl(sMascaraCclPadrao)
    If lErro <> SUCESSO Then gError 30245

    Ccl.Mask = sMascaraCclPadrao
    CclPadrao.Mask = sMascaraCclPadrao

    'Formata Quantidade
    Quantidade.Format = FORMATO_ESTOQUE

    'Inicialização do GridMovimentos
    Set objGrid = New AdmGrid

    lErro = Inicializa_GridMovimentos(objGrid)
    If lErro <> SUCESSO Then gError 30246

    'inicializacao da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_SERVICOS)
    If lErro <> SUCESSO Then gError 39612

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_From_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 30242, 30243, 30244, 30245, 30246, 39612, 49624, 49625, 78691

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162812)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Carrega_FilialOP() As Long
'Carrega a combobox FilialOP

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_FilialOP

    'Lê o Código e o Nome de toda FilialOP do BD
    lErro = CF("Cod_Nomes_Le_FilEmp", colCodigoNome)
    If lErro <> SUCESSO Then gError 78690

    'Carrega a combo de Filial Empresa com código e nome
    For Each objCodigoNome In colCodigoNome
        FilialOP.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        FilialOP.ItemData(FilialOP.NewIndex) = objCodigoNome.iCodigo
    Next

    Carrega_FilialOP = SUCESSO

    Exit Function

Erro_Carrega_FilialOP:

    Carrega_FilialOP = gErr

    Select Case gErr

        Case 78690 'Erro já tratado

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162813)

    End Select

    Exit Function

End Function

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

'Extrai os campos da tela que correspondem aos campos no Banco de Dados
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objTipoMovEst As ClassTipoMovEst
Dim objMovEstoque As New ClassMovEstoque

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "MovEstoqueInterno2"

    'Lê os atributos de objMovEstoque que aparecem na Tela
    If Len(Trim(Codigo.Text)) <> 0 Then objMovEstoque.lCodigo = CLng(Codigo.Text)

    If Len(Data.ClipText) <> 0 Then
        objMovEstoque.dtData = CDate(Data.Text)

    Else
        objMovEstoque.dtData = DATA_NULA

    End If

    If Len(Trim(Hora.ClipText)) > 0 Then
        objMovEstoque.dtHora = CDate(Hora.Text)
    Else
        objMovEstoque.dtHora = 0
    End If

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do Banco de Dados), tamanho do campo
    'no Banco de Dados no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objMovEstoque.lCodigo, 0, "Codigo"
    colCampoValor.Add "Data", objMovEstoque.dtData, 0, "Data"
    colCampoValor.Add "Hora", CDbl(objMovEstoque.dtHora), 0, "Hora"

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162814)

    End Select

    Exit Sub

End Sub

'Preenche os campos da tela com os correspondentes do Banco de Dados
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objMovEstoque As New ClassMovEstoque

On Error GoTo Erro_Tela_Preenche

    'Passa os dados da coleção para objReserva
    objMovEstoque.lCodigo = colCampoValor.Item("Codigo").vValor
    objMovEstoque.dtData = colCampoValor.Item("Data").vValor
    objMovEstoque.dtHora = colCampoValor.Item("Hora").vValor

    lErro = Preenche_Tela(objMovEstoque)
    If lErro <> SUCESSO Then gError 30247

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 30247

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162815)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoCodigo = Nothing
    Set objEventoCclPadrao = Nothing
    Set objEventoAlmoxPadrao = Nothing
    Set objEventoProduto = Nothing
    Set objEventoEstoque = Nothing
    Set objEventoCcl = Nothing
    Set objEventoContaContabil = Nothing
    Set objEventoOP = Nothing
    Set objEventoProdutoOP_Grid = Nothing
    Set objEventoRastroLote = Nothing
    Set objEventoOP_OP = Nothing
    Set objEventoProdutoOP_OP = Nothing
    Set objEventoOP_OP_Todos = Nothing
    
    Set colTipoMovEst = Nothing
    Set colItensNumIntDoc = Nothing
    
    Set colEntradaOuSaida = Nothing

    'Eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing

    Set objGrid = Nothing
    Set objGrid1 = Nothing
    Set objContabil = Nothing
    Set gobjAnotacao = Nothing
     
    Set gcolcolRastreamentoSerie = Nothing
     
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    Call ComandoSeta_Liberar(Me.Name)
   
End Sub

Private Sub GridMovimentos_RowColChange()

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_GridMovimentos_RowColChange

    Call Grid_RowColChange(objGrid)

    If (GridMovimentos.Row <> iLinhaAntiga) Then

        'Guarda a Linha usada
        iLinhaAntiga = GridMovimentos.Row

        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 141946
            
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 141947

        If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col))) = 0 Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
            lErro = QuantDisponivel_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col))
            If lErro <> SUCESSO Then gError 30274
        Else
            lErro = QuantDisponivel_Lote_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col)))
            If lErro <> SUCESSO Then gError 78681
        End If
        
    End If

    Exit Sub

Erro_GridMovimentos_RowColChange:

    Select Case gErr

        Case 30274, 78681

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162816)

    End Select

    Exit Sub

End Sub

Private Sub objEventoAlmoxPadrao_evSelecao(obj1 As Object)

Dim objAlmoxarifado As New ClassAlmoxarifado

    Set objAlmoxarifado = obj1

    AlmoxPadrao.Text = objAlmoxarifado.sNomeReduzido

    Me.Show

End Sub

Private Sub objEventoCcl_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclMascarado As String
Dim sCclFormatada As String

On Error GoTo Erro_objEventoCcl_evSelecao

    Set objCcl = obj1

    'Se o produto da linha corrente estiver preenchido e Linha corrente diferente da Linha fixa
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))) <> 0 And GridMovimentos.Row <> 0 Then

        'Mascara para colocar na tela
        sCclMascarado = String(STRING_CCL, 0)

        lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then gError 30381

        'Coloca o valor do Ccl na coluna correspondente
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Ccl_Col) = sCclMascarado

        Ccl.PromptInclude = False
        Ccl.Text = sCclMascarado
        Ccl.PromptInclude = True

    End If

    Me.Show

    Exit Sub

Erro_objEventoCcl_evSelecao:

    Select Case gErr

        Case 30381

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162817)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCclPadrao_evSelecao(obj1 As Object)

Dim objCcl As New ClassCcl
Dim sCclFormatada As String
Dim sCclMascarado As String
Dim lErro As Long

On Error GoTo Erro_objEventoCclPadrao_evSelecao

    Set objCcl = obj1

    sCclMascarado = String(STRING_CCL, 0)

    lErro = Mascara_RetornaCclEnxuta(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then gError 30249

    CclPadrao.PromptInclude = False
    CclPadrao.Text = sCclMascarado
    CclPadrao.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoCclPadrao_evSelecao:

    Select Case gErr

        Case 30249
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", gErr, objCcl.sCcl)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162818)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objMovEstoque As New ClassMovEstoque

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objMovEstoque = obj1

    lErro = Preenche_Tela(objMovEstoque)
    If lErro <> SUCESSO Then gError 30248

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 30248

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162819)

    End Select

    Exit Sub

End Sub

Private Sub objEventoContaContabil_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaContabil_evSelecao

    Set objPlanoConta = obj1

    If objPlanoConta.sConta <> "" Then
    
        If GridMovimentos.Col <> iGrid_ContaContabilCredito_Col And GridMovimentos.Col <> iGrid_ContaContabilDebito_Col Then
            Me.Show
            Exit Sub
        End If
        
        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then gError 43761
        
        If GridMovimentos.Col = iGrid_ContaContabilCredito_Col Then
            ContaContabilCredito.PromptInclude = False
            ContaContabilCredito.Text = sContaEnxuta
            ContaContabilCredito.PromptInclude = True
        Else
            ContaContabilDebito.PromptInclude = False
            ContaContabilDebito.Text = sContaEnxuta
            ContaContabilDebito.PromptInclude = True
        End If
        
        GridMovimentos.TextMatrix(GridMovimentos.Row, GridMovimentos.Col) = objGrid.objControle.Text
    
    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case gErr

        Case 43761
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162820)

    End Select

    Exit Sub

End Sub

Private Sub objEventoEstoque_evselecao(obj1 As Object)

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim sCodProduto As String

On Error GoTo Erro_objEventoEstoque_evselecao

    Set objEstoqueProduto = obj1

    sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 30407

    'Verifica se o produto está preenchido e se a linha corrente é diferente da linha fixa
    If iProdutoPreenchido = PRODUTO_PREENCHIDO And GridMovimentos.Row <> 0 Then

        'Preenche o Nome do Almoxarifado
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col) = objEstoqueProduto.sAlmoxarifadoNomeReduzido

        Almoxarifado.Text = objEstoqueProduto.sAlmoxarifadoNomeReduzido

        If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col))) = 0 Then
            'Calcula a Quantidade Disponível nesse Almoxarifado
            lErro = QuantDisponivel_Calcula(sCodProduto, objEstoqueProduto.sAlmoxarifadoNomeReduzido)
            If lErro <> SUCESSO Then gError 30255
        Else
            'Calcula a Quantidade Disponível nesse Almoxarifado - Lote
            lErro = QuantDisponivel_Lote_Calcula(sCodProduto, objEstoqueProduto.sAlmoxarifadoNomeReduzido)
            If lErro <> SUCESSO Then gError 78682
        End If
        
        'Conta contabil ---> vem como PADRAO da  tabela EstoqueProduto se o Produto e o Almoxarifado estiverem Preenchidos
        lErro = Preenche_Contas()
        If lErro <> SUCESSO Then gError 49596

    End If

    Me.Show

    Exit Sub

Erro_objEventoEstoque_evselecao:

    Select Case gErr

        Case 30255
            QuantDisponivel.Caption = ""
            Almoxarifado.Text = ""
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col) = ""
        
        Case 30407, 49596, 78682

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162821)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String
Dim objTipoDeProduto As New ClassTipoDeProduto

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))) = 0 Then
    
        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 30250
    
        If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then
    
            sProdutoMascarado = String(STRING_PRODUTO, 0)
    
            lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 30251
    
            'Lê os demais atributos do Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 30385
    
            If lErro = 28030 Then gError 30386
                       
            Produto.PromptInclude = False
            Produto.Text = sProdutoMascarado
            Produto.PromptInclude = True
    
            If Not (Me.ActiveControl Is Produto) Then
    
                GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col) = sProdutoMascarado
                           
                'Preenche a Linha do Grid
                lErro = ProdutoLinha_Preenche(objProduto)
                If lErro <> SUCESSO Then gError 30252
        
        
                If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col))) = 0 Then
        
                    lErro = QuantDisponivel_Calcula1(sProdutoMascarado, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col), objProduto)
                    If lErro <> SUCESSO Then gError 30408
                
                Else
                    
                    lErro = QuantDisponivel_Lote_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col)))
                    If lErro <> SUCESSO Then gError 78683
                            
                End If
                
            End If
    
        Else
    
            iProdutoPreenchido = 0
            sProdutoMascarado = ""
    
            lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 30253
    
            If GridMovimentos.Col = iGrid_ProdutoOP_Col And iProdutoPreenchido <> PRODUTO_PREENCHIDO Then
    
                sProdutoMascarado = String(STRING_PRODUTO, 0)
    
                lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
                If lErro <> SUCESSO Then gError 30254
    
                GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ProdutoOP_Col) = sProdutoMascarado
    
                ProdutoOP.PromptInclude = False
                ProdutoOP.Text = sProdutoMascarado
                ProdutoOP.PromptInclude = True
    
            End If
        End If
    End If

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 30250, 30252, 30253, 30385, 30408, 78683

        Case 30251, 30254
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objProduto.sCodigo)

        Case 30386
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162822)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoOP_Grid_evSelecao(obj1 As Object)

Dim objItemOP As ClassOS
Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String
Dim objProduto As New ClassProduto

On Error GoTo Erro_objEventoProdutoOP_Grid_evSelecao

    Set objItemOP = obj1
    
    If GridMovimentos.Row > 0 Then
    
        lErro = CF("Produto_Formata", objItemOP.sServico, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 52256

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            If (GridMovimentos.Col = iGrid_ProdutoOP_Col) And (iProdutoPreenchido = PRODUTO_PREENCHIDO) Then

                sProdutoMascarado = String(STRING_PRODUTO, 0)

                lErro = Mascara_MascararProduto(objItemOP.sServico, sProdutoMascarado)
                If lErro <> SUCESSO Then gError 52257

                GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ProdutoOP_Col) = sProdutoMascarado

                ProdutoOP.PromptInclude = False
                ProdutoOP.Text = sProdutoMascarado
                ProdutoOP.PromptInclude = True
            
            End If
            
        End If
        
    End If

    Me.Show

    Exit Sub

Erro_objEventoProdutoOP_Grid_evSelecao:

    Select Case gErr

        Case 52256
        
        Case 52257
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objItemOP.sServico)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162823)

    End Select

End Sub

Function Trata_Parametros(Optional objMovEstoque As ClassMovEstoque) As Long

Dim lErro As Long
Dim lCodigo As Long
Dim objTipoMovEst As ClassTipoMovEst

On Error GoTo Erro_Trata_Parametros

    'Se há um Movestoque passado como parâmetro
    If Not objMovEstoque Is Nothing Then

        objMovEstoque.iFilialEmpresa = giFilialEmpresa

        'Lê MovEstoque no Banco de Dados
        lErro = CF("MovEstoque_Le", objMovEstoque)
        If lErro <> SUCESSO And lErro <> 30128 Then gError 30256

        If lErro <> 30128 Then 'Se ele existe

            Set objTipoMovEst = New ClassTipoMovEst

            objTipoMovEst.iCodigo = objMovEstoque.iTipoMov

            lErro = CF("TipoMovEstoque_Le", objTipoMovEst)
            If lErro <> SUCESSO And lErro <> 30372 Then gError 30257

            If lErro = 30372 Then gError 30373

            If objTipoMovEst.iValidoMovInt <> MOVESTOQUEINTSRV Then gError 30258

            lErro = Preenche_Tela(objMovEstoque)
            If lErro <> SUCESSO Then gError 30259

        Else
            'Se ele não existe exibe apenas o código
            Codigo.PromptInclude = False
            Codigo.Text = CStr(objMovEstoque.lCodigo)
            Codigo.PromptInclude = True

            lCodigoAntigo = objMovEstoque.lCodigo

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 30256, 30257, 30259

        Case 30258
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_NAO_INTERNO", gErr, objMovEstoque.lCodigo)

        Case 30373
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMOVEST_NAO_CADASTRADO", gErr, objTipoMovEst.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162824)

    End Select

    iAlterado = 0

    Exit Function

End Function

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    If Len(Data.ClipText) = 0 Then Exit Sub

    lErro = Data_Critica(Data.Text)
    If lErro <> SUCESSO Then gError 30264

    Exit Sub

Erro_Data_Validate:

    Cancel = True


    Select Case gErr

        Case 30264

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162825)

    End Select

    Exit Sub

End Sub

'hora
Public Sub Hora_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Hora, iAlterado)

End Sub

'hora
Public Sub Hora_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

'hora
Public Sub Hora_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Hora_Validate

    'Verifica se a hora foi digitada
    If Len(Trim(Hora.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Hora_Critica(Hora.Text)
    If lErro <> SUCESSO Then gError 89805

    Exit Sub

Erro_Hora_Validate:

    Cancel = True

    Select Case gErr

        Case 89805

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162826)

    End Select

    Exit Sub

End Sub

Private Sub Opcao_Click()

Dim iEstorno As Integer
Dim iLinha As Integer

    'Se frame selecionado não for o atual
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        'Esconde o frame atual, mostra o novo
        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index

        'se estiver selecionando o tabstrip de contabilidade e o usuário não alterou a contabilidade ==> carrega o modelo padrao
        If Opcao.SelectedItem.Caption = TITULO_TAB_CONTABILIDADE Then
            
            'trata os casos de estorno da versao light
            Call MovEstoque_Trata_Estorno_Versao_Light
        
            Call objContabil.Contabil_Carga_Modelo_Padrao

        End If
        
        Select Case iFrameAtual
        
            Case TAB_Movimentos
                Parent.HelpContextID = IDH_MOVIMENTOS_ESTOQUE_MOVIMENTO
                
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_MOVIMENTOS_ESTOQUE_CONTABILIZACAO
                        
        End Select

    End If

End Sub

Private Sub OPCodigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProdutoOP_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoMov_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UnidadeMed_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    If Len(Data.ClipText) = 0 Then Exit Sub

    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 30266

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 30266

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162827)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    If Len(Data.ClipText) = 0 Then Exit Sub

    lErro = lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 30265

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 30265

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162828)

    End Select

    Exit Sub

End Sub

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim iIndice As Integer, lNumIntDoc As Long
Dim sUnidadeMed As String
Dim sCodProduto As String
Dim objProduto As New ClassProduto
Dim objClasseUM As New ClassClasseUM
Dim objUnidadeDeMedida As ClassUnidadeDeMedida
Dim colSiglas As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iTipoMov As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    'Verifica se produto está preenchido
    sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 30267

    If colItensNumIntDoc.Count >= GridMovimentos.Row Then
        lNumIntDoc = colItensNumIntDoc.Item(GridMovimentos.Row)
    Else
        lNumIntDoc = 0
    End If
    
    iTipoMov = Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoMov_Col))

    If objControl.Name = "Produto" Then

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objControl.Enabled = False

        Else
            objControl.Enabled = True

        End If

    ElseIf objControl.Name = "UnidadeMed" Then

        If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then
            objControl.Enabled = False

        Else
            objControl.Enabled = True

            objProduto.sCodigo = sProdutoFormatado

            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 30269

            If lErro = 28030 Then gError 30270

            objClasseUM.iClasse = objProduto.iClasseUM

            'Preenche a List da Combo UnidadeMed com as UM's do Produto
            lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
            If lErro <> SUCESSO Then gError 30271

            'Guardo o valor da Unidade de Medida da Linha
            sUnidadeMed = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col)

            'Limpar as Unidades utilizadas anteriormente
            UnidadeMed.Clear

            For Each objUnidadeDeMedida In colSiglas
                UnidadeMed.AddItem objUnidadeDeMedida.sSigla
            Next

            UnidadeMed.AddItem ""

            'Tento selecionar na Combo a Unidade anterior
            If UnidadeMed.ListCount <> 0 Then

                For iIndice = 0 To UnidadeMed.ListCount - 1

                    If UnidadeMed.List(iIndice) = sUnidadeMed Then
                        UnidadeMed.ListIndex = iIndice
                        Exit For
                    End If
                Next
            End If


            If lNumIntDoc = 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If

        End If

    ElseIf objControl.Name = "TipoMov" Or objControl.Name = "Almoxarifado" Or objControl.Name = "Ccl" Or objControl.Name = "OPCodigo" Or objControl.Name = "ProdutoOP" Or objControl.Name = "ContaContabilCredito" Or objControl.Name = "ContaContabilDebito" Then

            If iProdutoPreenchido = PRODUTO_PREENCHIDO And lNumIntDoc = 0 Then

                objControl.Enabled = True

            Else
                objControl.Enabled = False

            End If
            
    ElseIf objControl.Name = "Quantidade" Then
    
        If iTipoMov >= 300 And iTipoMov <= 304 Then
            objControl.Enabled = False
        Else
            If iProdutoPreenchido = PRODUTO_PREENCHIDO And lNumIntDoc = 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
        End If
    
    ElseIf objControl.Name = "Custo" Then
    
        If iTipoMov >= 300 And iTipoMov <= 304 Then
            objControl.Enabled = True
        Else
            objControl.Enabled = False
        End If

    ElseIf objControl.Name = "Lote" Then

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            objProduto.sCodigo = sProdutoFormatado
    
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 78302
    
            If lErro = 28030 Then gError 78303
        
            If objProduto.iRastro = PRODUTO_RASTRO_NENHUM Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
        Else
            objControl.Enabled = False
        End If
    
    ElseIf objControl.Name = "FilialOP" Then

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            objProduto.sCodigo = sProdutoFormatado

            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 78720

            If lErro = 28030 Then gError 78721

            If objProduto.iRastro = PRODUTO_RASTRO_OP Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
        Else
            objControl.Enabled = False
        End If

    ElseIf objControl.Name = "Estorno" Then

        If lNumIntDoc = 0 Then

            objControl.Enabled = False

        Else
            objControl.Enabled = True

        End If

    End If

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 30267, 30268, 30269, 30270, 30271, 30272, 30273, 78302, 78303, 78317, 78318, 78720, 78721

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162829)

    End Select

    Exit Sub

End Sub

Private Sub Testa_TipoMovimento(ByVal iIndice As Integer, iEntradaSaida As Integer)
'Testa se é movimento de entrada ou saída de material. bEntradaSaida = ESTOQUE_ENTRADA (Entrada de Material no Estoque) ESTOQUE_SAIDA (Saida de Material do Estoque)

Dim lErro As Long
Dim iTipoMov As Integer
Dim iCodigoLinha As Integer
Dim objEntradaOuSaida As AdmCodigoNome
Dim iEstorno As Integer
Dim sEntradaOuSaida As String

On Error GoTo Erro_Testa_TipoMovimento

    iEntradaSaida = -1

'    If GridMovimentos.TextMatrix(iIndice, iGrid_Estorno_Col) = "1" Then
'        iEstorno = MOVIMENTO_ESTORNO
'    Else
        iEstorno = MOVIMENTO_NORMAL
'    End If

    If Len(GridMovimentos.TextMatrix(iIndice, iGrid_TipoMov_Col)) <> 0 Then
        
        iCodigoLinha = Codigo_Extrai(GridMovimentos.TextMatrix(iIndice, iGrid_TipoMov_Col))
        For Each objEntradaOuSaida In colEntradaOuSaida
            If iCodigoLinha = objEntradaOuSaida.iCodigo Then
                sEntradaOuSaida = objEntradaOuSaida.sNome
                Exit For
            End If
        Next
        
        Select Case iEstorno
        
            Case MOVIMENTO_ESTORNO
        
                If sEntradaOuSaida = TIPO_ENTRADA Then
                    iEntradaSaida = ESTOQUE_SAIDA
                Else
                    iEntradaSaida = ESTOQUE_ENTRADA
                End If
        
            Case MOVIMENTO_NORMAL
        
                If sEntradaOuSaida = TIPO_ENTRADA Then
                    iEntradaSaida = ESTOQUE_ENTRADA
                Else
                    iEntradaSaida = ESTOQUE_SAIDA
                End If
        End Select
    
    End If
    
    Exit Sub

Erro_Testa_TipoMovimento:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162830)

    End Select

    Exit Sub

End Sub

Private Sub ContaContabilDebito_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ContaContabilDebito_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub ContaContabilDebito_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub ContaContabilDebito_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGrid.objControle = ContaContabilDebito
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ContaContabilCredito_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ContaContabilCredito_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub ContaContabilCredito_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub ContaContabilCredito_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGrid.objControle = ContaContabilCredito
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 39613

        If objGridInt.objGrid Is GridMovimentos Then

            Select Case objGridInt.objGrid.Col

                Case iGrid_Almoxarifado_Col
                    lErro = Saida_Celula_Almoxarifado(objGridInt)
                    If lErro <> SUCESSO Then gError 30288

                Case iGrid_Ccl_Col
                    lErro = Saida_Celula_Ccl(objGridInt)
                    If lErro <> SUCESSO Then gError 30289

                Case iGrid_Produto_Col
                    lErro = Saida_Celula_Produto(objGridInt)
                    If lErro <> SUCESSO Then gError 30291

                Case iGrid_Quantidade_Col
                    lErro = Saida_Celula_Quantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 30292

'                Case iGrid_Estorno_Col
'                    lErro = Saida_Celula_Estorno(objGridInt)
'                    If lErro <> SUCESSO Then gError 30293

                Case iGrid_UnidadeMed_Col
                    lErro = Saida_Celula_UnidadeMed(objGridInt)
                    If lErro <> SUCESSO Then gError 30294

                Case iGrid_TipoMov_Col
                    lErro = Saida_Celula_TipoMov(objGridInt)
                    If lErro <> SUCESSO Then gError 30296

                Case iGrid_OP_Col
                    lErro = Saida_Celula_OP(objGridInt)
                    If lErro <> SUCESSO Then gError 30297

                Case iGrid_ProdutoOP_Col
                    lErro = Saida_Celula_ProdutoOP(objGridInt)
                    If lErro <> SUCESSO Then gError 30298
                    
                Case iGrid_ContaContabilCredito_Col
                    lErro = Saida_Celula_ContaContabilCredito(objGridInt)
                    If lErro <> SUCESSO Then gError 49626
                    
                Case iGrid_ContaContabilDebito_Col
                    lErro = Saida_Celula_ContaContabilDebito(objGridInt)
                    If lErro <> SUCESSO Then gError 49627
                    
                Case iGrid_Lote_Col
                    lErro = Saida_Celula_Lote(objGridInt)
                    If lErro <> SUCESSO Then gError 78310
                
                Case iGrid_FilialOP_Col
                    lErro = Saida_Celula_FilialOP(objGridInt)
                    If lErro <> SUCESSO Then gError 78703
                    
                Case iGrid_Custo_Col
                    lErro = Saida_Celula_Custo(objGridInt)
                    If lErro <> SUCESSO Then gError 132078
                    
            End Select

        End If

    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30299

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 30288, 30289, 30291, 30292, 30293, 30294, 30296, 30297, 30298, 49626, 49627, 78310, 78703

        Case 30299
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 39613, 132078

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162831)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Lote(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objRastroLote As New ClassRastreamentoLote
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objOrdemProducao As New ClassOS
Dim iEntradaSaida As Integer

On Error GoTo Erro_Saida_Celula_Lote

    Set objGridInt.objControle = Lote
    
    If Len(Trim(Lote.Text)) > 0 Then
        
        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 78451
            
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 78452
            
        If lErro = 28030 Then gError 78453
                
        'Se o Produto foi preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
            'Se for rastro por lote
            If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then

                objRastroLote.sCodigo = Lote.Text
                objRastroLote.sProduto = sProdutoFormatado

                'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                lErro = CF("RastreamentoLote_Le", objRastroLote)
                If lErro <> SUCESSO And lErro <> 75710 Then gError 78705

                'Se não encontrou --> Erro
                If lErro = 75710 Then gError 78706

                'Preenche a Quantidade do Lote
                lErro = QuantDisponivel_Lote_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col), Lote.Text, Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col)))
                If lErro <> SUCESSO Then gError 78735
                    
            'Se for rastro por OP
            ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then

                If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col))) > 0 Then

                    objRastroLote.sCodigo = Lote.Text
                    objRastroLote.sProduto = sProdutoFormatado
                    objRastroLote.iFilialOP = Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col))

                    'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                    lErro = CF("RastreamentoLote_Le", objRastroLote)
                    If lErro <> SUCESSO And lErro <> 75710 Then gError 78710

                    'Se não encontrou --> Erro
                    If lErro = 75710 Then gError 78711

                    'Preenche a Quantidade do Lote
                    lErro = QuantDisponivel_Lote_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col), Lote.Text, Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col)))
                    If lErro <> SUCESSO Then gError 78727
                
                Else
        
                    'Preenche a Quantidade do Lote
                    lErro = QuantDisponivel_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col))
                    If lErro <> SUCESSO Then gError 78732
                
                End If

            'Se for rastro por OP
            ElseIf objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
                    
                'Preenche a Quantidade do Lote
                lErro = QuantDisponivel_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col))
                If lErro <> SUCESSO Then gError 78732
                            
            End If
            
        End If
    
    Else
    
        'Preenche a Quantidade do Lote
        lErro = QuantDisponivel_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col))
        If lErro <> SUCESSO Then gError 78728

    End If
            
    'Testa se é movimento de estorno , entrada , saída
    Call Testa_TipoMovimento(GridMovimentos.Row, iEntradaSaida)
        
    If Len(Trim(QuantDisponivel.Caption)) > 0 And iEntradaSaida = ESTOQUE_SAIDA Then

        lErro = Testa_QuantRequisitada(StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)))
        If lErro <> SUCESSO Then gError 78730

    End If

    'Carrega as séries na coleção global
    lErro = Carrega_Series(gcolcolRastreamentoSerie.Item(GridMovimentos.Row), StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)), Lote.Text, StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), GridMovimentos.Row)
    If lErro <> SUCESSO Then gError 141912

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 78458

    Saida_Celula_Lote = SUCESSO

    Exit Function

Erro_Saida_Celula_Lote:

    Saida_Celula_Lote = gErr

    Select Case gErr

        Case 78451, 78452, 78454, 78458, 78705, 78710, 78727, 78728, 78730, 78732, 78735, 141912 '78707
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78453
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78706, 78711
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("RastreamentoLote", objRastroLote)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162832)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_FilialOP(objGridInt As AdmGrid) As Long
'Faz a saida de celula da Filial da Ordem de Produção

Dim lErro As Long
Dim objFilialOP As New AdmFiliais
Dim iCodigo As Integer
Dim objOrdemProducao As New ClassOS
Dim objRastroLote As New ClassRastreamentoLote
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim iEntradaSaida As Integer

On Error GoTo Erro_Saida_Celula_FilialOP

    Set objGridInt.objControle = FilialOP

    If Len(Trim(FilialOP.Text)) <> 0 Then
            
        'Verifica se é uma FilialOP selecionada
        If FilialOP.Text <> FilialOP.List(FilialOP.ListIndex) Then
        
            'Tenta selecionar na combo
            lErro = Combo_Seleciona(FilialOP, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 78692
    
            'Se não encontrou o ítem com o código informado
            If lErro = 6730 Then
    
                objFilialOP.iCodFilial = iCodigo
    
                'Pesquisa se existe FilialOP com o codigo extraido
                lErro = CF("FilialEmpresa_Le", objFilialOP)
                If lErro <> SUCESSO And lErro <> 27378 Then gError 78693
        
                'Se não encontrou a FilialOP
                If lErro = 27378 Then gError 78694
        
                'coloca na tela
                FilialOP.Text = iCodigo & SEPARADOR & objFilialOP.sNome
            
            End If
    
            'Não encontrou valor informado que era STRING
            If lErro = 6731 Then gError 78695
                    
        End If
        
        If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col))) > 0 Then

            lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 78699
                                
            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
                objRastroLote.sCodigo = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col)
                objRastroLote.sProduto = sProdutoFormatado
                objRastroLote.iFilialOP = Codigo_Extrai(FilialOP.Text)
            
                'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                lErro = CF("RastreamentoLote_Le", objRastroLote)
                If lErro <> SUCESSO And lErro <> 75710 Then gError 78700
                
                'Se não encontrou --> Erro
                If lErro = 75710 Then gError 78701
                            
                'Preenche a Quantidade do Lote
                lErro = QuantDisponivel_Lote_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), Codigo_Extrai(FilialOP.Text))
                If lErro <> SUCESSO Then gError 78725
                
            End If
            
        End If
        
    Else
    
        'Preenche a Quantidade do Lote
        lErro = QuantDisponivel_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col))
        If lErro <> SUCESSO Then gError 78726
    
    End If
        
    'Testa se é movimento de estorno , entrada , saída
    Call Testa_TipoMovimento(GridMovimentos.Row, iEntradaSaida)
        
    If Len(Trim(QuantDisponivel.Caption)) > 0 And iEntradaSaida = ESTOQUE_SAIDA Then

        lErro = Testa_QuantRequisitada(StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)))
        If lErro <> SUCESSO Then gError 78731

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 78702

    Saida_Celula_FilialOP = SUCESSO

    Exit Function

Erro_Saida_Celula_FilialOP:

    Saida_Celula_FilialOP = gErr

    Select Case gErr

        Case 78692, 78693, 78696, 78699, 78700, 78702, 78725, 78726, 78731
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 78694
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, FilialOP.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78695
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, FilialOP.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78697
            Call Rotina_Erro(vbYesNo, "ERRO_OSCODIGO_NAO_CADASTRADO", gErr, objOrdemProducao.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78701
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("RastreamentoLote", objRastroLote)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162833)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ContaContabilDebito(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaMascarada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim iContaPreenchida As Integer

On Error GoTo Erro_Saida_Celula_ContaContabilDebito
    
    Set objGridInt.objControle = ContaContabilDebito
    
    If Len(Trim(ContaContabilDebito.ClipText)) > 0 Then
    
        'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
        lErro = CF("ContaSimples_Critica_Modulo", sContaFormatada, ContaContabilDebito.ClipText, objPlanoConta, MODULO_ESTOQUE)
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 49632
        
        If lErro = SUCESSO Then
        
            sContaFormatada = objPlanoConta.sConta
            
            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
            If lErro <> SUCESSO Then gError 49633
            
            ContaContabilDebito.PromptInclude = False
            ContaContabilDebito.Text = sContaMascarada
            ContaContabilDebito.PromptInclude = True
        
        'se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then
    
            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaContabilDebito.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
            If lErro <> SUCESSO And lErro <> 5700 Then gError 49634
    
            'conta não cadastrada
            If lErro = 5700 Then gError 49635
             
        End If
    
    Else
        
        ContaContabilDebito.PromptInclude = False
        ContaContabilDebito.Text = ""
        ContaContabilDebito.PromptInclude = True
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 49647
    
    Saida_Celula_ContaContabilDebito = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaContabilDebito:

    Saida_Celula_ContaContabilDebito = gErr

    Select Case gErr

        Case 49632, 49634, 49647
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 49633
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
            
        Case 49635
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabilDebito.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
                            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162834)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ContaContabilCredito(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaMascarada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim iContaPreenchida As Integer

On Error GoTo Erro_Saida_Celula_ContaContabilCredito

    Set objGridInt.objControle = ContaContabilCredito

    If Len(Trim(ContaContabilCredito.ClipText)) > 0 Then
    
        'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
        lErro = CF("ContaSimples_Critica_Modulo", sContaFormatada, ContaContabilCredito.ClipText, objPlanoConta, MODULO_ESTOQUE)
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 49628
        
        If lErro = SUCESSO Then
        
            sContaFormatada = objPlanoConta.sConta
            
            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
            If lErro <> SUCESSO Then gError 49629
            
            ContaContabilCredito.PromptInclude = False
            ContaContabilCredito.Text = sContaMascarada
            ContaContabilCredito.PromptInclude = True
        
        'se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then
    
            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaContabilCredito.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
            If lErro <> SUCESSO And lErro <> 5700 Then gError 49630
    
            'conta não cadastrada
            If lErro = 5700 Then gError 49631
             
        End If
    
    Else
        
        ContaContabilCredito.PromptInclude = False
        ContaContabilCredito.Text = ""
        ContaContabilCredito.PromptInclude = True
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 49643
    
   Saida_Celula_ContaContabilCredito = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaContabilCredito:

    Saida_Celula_ContaContabilCredito = gErr

    Select Case gErr

        Case 49628, 49630, 49643
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 49629
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 49631
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabilCredito.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
                            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162835)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult
Dim objTipoDeProduto As New ClassTipoDeProduto
Dim sProduto As String

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto
    
    iProdutoPreenchido = PRODUTO_VAZIO

    If Len(Produto.ClipText) <> 0 Then

        sProduto = Produto.Text

        lErro = CF("Trata_Segmento_Produto", sProduto)
        If lErro <> SUCESSO Then gError 199348

        Produto.Text = sProduto

        lErro = CF("Produto_Critica_Estoque", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25077 Then gError 30300

        If lErro = 25077 Then gError 30301

    End If

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
        If giTipoVersao = VERSAO_FULL Then
            If Trim(objProduto.sContaContabil) = "" Then
                
                objTipoDeProduto.iTipo = objProduto.iTipo
                
                lErro = CF("TipoDeProduto_Le", objTipoDeProduto)
                If lErro <> SUCESSO And lErro <> 22531 Then gError 52007
                
                If lErro = 22531 Then gError 52008
                
                objProduto.sContaContabil = objTipoDeProduto.sContaContabil
                            
            End If
        End If
        
        lErro = ProdutoLinha_Preenche(objProduto)
        If lErro <> SUCESSO Then gError 30303

        If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col))) = 0 Then
            'Calcula a Quantidade Disponível
            lErro = QuantDisponivel_Calcula1(Produto.Text, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col), objProduto)
            If lErro <> SUCESSO Then gError 30304
        Else
            'Calcula a Quantidade Disponível do lote
            lErro = QuantDisponivel_Lote_Calcula1(Produto.Text, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col)), objProduto)
            If lErro <> SUCESSO Then gError 78686
        End If
        
        If objProduto.iRastro = PRODUTO_RASTRO_OP Then
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col) = giFilialEmpresa & SEPARADOR & gsNomeFilialEmpresa
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30302

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 30300, 30302, 30303, 30304, 52007, 78686, 199348
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 30301
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)

            If vbMsg = vbYes Then
                objProduto.sCodigo = Produto.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Produto", objProduto)


            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If
        
        Case 52008
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_NAO_CADASTRADO", gErr, objTipoDeProduto.iTipo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162836)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iEntradaSaida As Integer

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    If Len(Trim(Quantidade.Text)) > 0 Then

        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 30305

        'Caso QuantDisponivel estiver preenchida verificar se é maior
        If colItensNumIntDoc.Item(GridMovimentos.Row) = 0 Then
            
            'Testa se é movimento de estorno , entrada , saída
            Call Testa_TipoMovimento(GridMovimentos.Row, iEntradaSaida)
            
            If Len(Trim(QuantDisponivel.Caption)) > 0 And iEntradaSaida = ESTOQUE_SAIDA Then

                lErro = Testa_QuantRequisitada(CDbl(Quantidade.Text))
                If lErro <> SUCESSO Then gError 30176

            End If

        End If

        Quantidade.Text = Formata_Estoque(Quantidade.Text)

    End If

    'Carrega as séries na coleção global
    lErro = Carrega_Series(gcolcolRastreamentoSerie.Item(GridMovimentos.Row), StrParaDbl(Quantidade.Text), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), GridMovimentos.Row)
    If lErro <> SUCESSO Then gError 141911

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30308

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 30176, 30305, 30308, 141911
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162837)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Function Saida_Celula_TipoMov(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iCodigo As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim iEntradaSaida As Integer
Dim dQuantidade As Double

On Error GoTo Erro_Saida_Celula_TipoMov

    Set objGridInt.objControle = TipoMov

    If Len(Trim(TipoMov.Text)) <> 0 Then
    
        lErro = Combo_Seleciona_Grid(TipoMov, iCodigo)
        If lErro <> SUCESSO And lErro <> 25085 And lErro <> 25086 Then gError 30309

        If lErro = 25085 Or lErro = 25086 Then gError 30310
        
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoMov_Col) = TipoMov.Text

        If Codigo_Extrai(TipoMov.Text) >= 300 And Codigo_Extrai(TipoMov.Text) <= 304 Then
            Custo.Enabled = True
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col) = ""
        Else
            Quantidade.Enabled = True
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Custo_Col) = ""
        End If

        If colItensNumIntDoc.Item(GridMovimentos.Row) = 0 Then
            
            'Testa se é movimento de entrada ou saída
            Call Testa_TipoMovimento(GridMovimentos.Row, iEntradaSaida)
            
            If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))) > 0 And iEntradaSaida = ESTOQUE_SAIDA Then
    
                dQuantidade = CDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))
    
                lErro = Testa_QuantRequisitada(dQuantidade)
                If lErro <> SUCESSO Then gError 55380
    
            End If
    
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30311
    
    'Carrega as séries na coleção global
    lErro = Carrega_Series(gcolcolRastreamentoSerie.Item(GridMovimentos.Row), StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), GridMovimentos.Row)
    If lErro <> SUCESSO Then gError 177288
    
    'preenche as contas de acordo com o Tipo de Movimento
    lErro = Preenche_Contas()
    If lErro <> SUCESSO Then gError 52123
    
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col))) = 0 Then
        lErro = QuantDisponivel_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col))
        If lErro <> SUCESSO Then gError 52123
    Else
        lErro = QuantDisponivel_Lote_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col)))
        If lErro <> SUCESSO Then gError 52123
    End If
    
    Saida_Celula_TipoMov = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoMov:

    Saida_Celula_TipoMov = gErr

    Select Case gErr

        Case 30309, 30311, 55377, 55378, 55380, 177288
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 30310
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMOVINT_NAO_CADASTRADO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 52123
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 55375
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PCP", gErr, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 55376
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL", gErr, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 55379
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162838)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Preenche_Contas() As Long
'preenche as contas de acordo com o tipo de movimento de estoque

Dim lErro As Long
Dim sContaEnxutaCredito As String
Dim sContaEnxutaDebito As String
Dim iCodigoTipo As Integer

On Error GoTo Erro_Preenche_Contas
        
    'Limpa a conta Debito
    ContaContabilDebito.PromptInclude = False
    ContaContabilDebito.Text = ""
    ContaContabilDebito.PromptInclude = True
    
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ContaContabilDebito_Col) = ""
    
    'limpa a conta Credito
    ContaContabilCredito.PromptInclude = False
    ContaContabilCredito.Text = ""
    ContaContabilCredito.PromptInclude = True
    
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ContaContabilCredito_Col) = ""
    
    'se o movimento esta preenchido preenche as contas
    If GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoMov_Col) <> "" Then
        
        iCodigoTipo = Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoMov_Col))
        
        'seleciona de acordo com o movimento
        Select Case iCodigoTipo
            
            'se for um movimento MOV_EST_REQ_PRODUCAO, MOV_EST_PRODUCAO, MOV_EST_CONSUMO
            Case MOV_EST_REQ_PRODUCAO, MOV_EST_CONSUMO, MOV_EST_PROCESSO, MOV_EST_OUTRAS_REQ
                 
                If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col))) > 0 And Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))) > 0 Then
         
                    'preenche a conta de Crédito
                    lErro = ContaContabilEst_Le(sContaEnxutaCredito)
                    If lErro <> SUCESSO Then gError 52088
                    
                    'preenche a conta credito com Conta estoque
                    ContaContabilCredito.PromptInclude = False
                    ContaContabilCredito.Text = sContaEnxutaCredito
                    ContaContabilCredito.PromptInclude = True
                    
                    If Len(Trim(sContaEnxutaCredito)) > 0 Then
                        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ContaContabilCredito_Col) = ContaContabilCredito.Text
                    End If
                    
                End If
                
                'preenche conta de Débito
                If iCodigoTipo = MOV_EST_REQ_PRODUCAO Or iCodigoTipo = MOV_EST_PROCESSO Then
                
                    lErro = ContaContabilProd_Le(sContaEnxutaDebito)
                    If lErro <> SUCESSO Then gError 52089
                    
                Else
                
                    lErro = ContaContabilAplic_Le(sContaEnxutaDebito)
                    If lErro <> SUCESSO Then gError 55372
                
                End If
                
                'preenche a conta Debito com a conta aplicacao
                ContaContabilDebito.PromptInclude = False
                ContaContabilDebito.Text = sContaEnxutaDebito
                ContaContabilDebito.PromptInclude = True
                
                If Trim(Len(sContaEnxutaDebito)) Then
                    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ContaContabilDebito_Col) = ContaContabilDebito.Text
                End If
                
            'se for um movimento do tipo MOV_EST_DEV_CONSUMO, MOV_EST_DEV_MATERIAL_PRODUCAO
            Case MOV_EST_PRODUCAO, MOV_EST_DEV_CONSUMO, MOV_EST_DEV_MATERIAL_PRODUCAO, MOV_EST_DEV_PROCESSO, MOV_EST_OUTRAS_DEV
                            
                If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col))) > 0 And Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))) > 0 Then
                             
                    'preenche a conta de Debito
                    lErro = ContaContabilEst_Le(sContaEnxutaDebito)
                    If lErro <> SUCESSO Then gError 52090
                    
                    'preenche a conta debito com a conta estoque
                    ContaContabilDebito.PromptInclude = False
                    ContaContabilDebito.Text = sContaEnxutaDebito
                    ContaContabilDebito.PromptInclude = True
                    
                    If Trim(Len(sContaEnxutaDebito)) Then
                        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ContaContabilDebito_Col) = ContaContabilDebito.Text
                    End If
                    
                End If
            
                'preenche conta de Credito
                If iCodigoTipo = MOV_EST_PRODUCAO Or iCodigoTipo = MOV_EST_DEV_MATERIAL_PRODUCAO Or iCodigoTipo = MOV_EST_DEV_PROCESSO Then
                
                    lErro = ContaContabilProd_Le(sContaEnxutaCredito)
                    If lErro <> SUCESSO Then gError 55373
                    
                Else
                
                    lErro = ContaContabilAplic_Le(sContaEnxutaCredito)
                    If lErro <> SUCESSO Then gError 52091
                    
                End If
                
                'preenche a conta Credito com a conta de aplicacao
                ContaContabilCredito.PromptInclude = False
                ContaContabilCredito.Text = sContaEnxutaCredito
                ContaContabilCredito.PromptInclude = True
                
                If Len(Trim(sContaEnxutaCredito)) > 0 Then
                    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ContaContabilCredito_Col) = ContaContabilCredito.Text
                End If
                
        End Select
    
    End If
    
    Exit Function
    
Preenche_Contas = SUCESSO

Erro_Preenche_Contas:

Preenche_Contas = gErr

    Select Case gErr
        
        Case 52088, 52089, 52090, 52091, 55372, 55373
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162839)
        
    End Select
    
    Exit Function
        
End Function

Private Function Saida_Celula_Almoxarifado(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim vbMsg As VbMsgBoxResult
Dim objProduto As New ClassProduto 'Inserido por Wagner 15/03/2006

On Error GoTo Erro_Saida_Celula_Almoxarifado

    Set objGridInt.objControle = Almoxarifado

    If Len(Trim(Almoxarifado.Text)) <> 0 Then

        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 30383

        lErro = TP_Almoxarifado_Filial_Produto_Grid(sProdutoFormatado, Almoxarifado, objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25157 And lErro <> 25162 Then gError 30312

        If lErro = 25157 Then gError 30313

        If lErro = 25162 Then gError 30314

        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido
        
        'Formata o Produto para o BD
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 141948
        
        If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col))) = 0 Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then   'Alterado por Wagner 15/03/2006
            lErro = QuantDisponivel_Calcula(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), Almoxarifado.Text)
            If lErro <> SUCESSO Then gError 30316
        Else
            'Calcula a Quantidade Disponível do lote
            lErro = QuantDisponivel_Lote_Calcula(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), Almoxarifado.Text)
            If lErro <> SUCESSO Then gError 78687
        End If
        
   Else

        'Limpa a Quantidade Disponível da Tela
        QuantDisponivel.Caption = ""

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30315
    
    'Conta contabil ---> vem como PADRAO da  tabela EstoqueProduto se o Produto e o Almoxarifado estiverem Preenchidos
    lErro = Preenche_Contas()
    If lErro <> SUCESSO Then gError 49706
   
    Saida_Celula_Almoxarifado = SUCESSO

    Exit Function

Erro_Saida_Celula_Almoxarifado:

    Saida_Celula_Almoxarifado = gErr

    Select Case gErr

        Case 30313

            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_INEXISTENTE", Almoxarifado.Text)

            If vbMsg = vbYes Then

                objAlmoxarifado.sNomeReduzido = Almoxarifado.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 30314

            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_INEXISTENTE1", CInt(Almoxarifado.Text))

            If vbMsg = vbYes Then

                objAlmoxarifado.iCodigo = CInt(Almoxarifado.Text)

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)


            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 30312, 30315, 30316, 30383, 49706, 78687, 141948
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162840)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Ccl(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sCclFormatada As String
Dim objCcl As New ClassCcl
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Ccl

    Set objGridInt.objControle = Ccl

    If Len(Ccl.Text) <> 0 Then

        lErro = CF("Ccl_Critica", Ccl.Text, sCclFormatada, objCcl)
        If lErro <> SUCESSO And lErro <> 5703 Then gError 30325

        If lErro = 5703 Then gError 30326

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30327

    Saida_Celula_Ccl = SUCESSO

    Exit Function

Erro_Saida_Celula_Ccl:

    Saida_Celula_Ccl = gErr

    Select Case gErr

        Case 30325, 30327
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 30326
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", Ccl.Text)
            If vbMsg = vbYes Then
            
                objCcl.sCcl = sCclFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("CclTela", objCcl)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162841)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_OP(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim vbMsg As VbMsgBoxResult
Dim objOrdemProducao As New ClassOS
Dim objItemOP As New ClassItemOS
Dim iProdutoOPPreenchido As Integer
Dim sProdutoOPFormatado As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim objRastroLote As New ClassRastreamentoLote

On Error GoTo Erro_Saida_Celula_OP

    Set objGridInt.objControle = OPCodigo

    If Len(Trim(OPCodigo.Text)) > 0 Then

        objOrdemProducao.iFilialEmpresa = giFilialEmpresa
        objOrdemProducao.sCodigo = OPCodigo.Text

        lErro = CF("OS_Le", objOrdemProducao)
        If lErro <> SUCESSO And lErro <> 195102 Then gError 30328

        If lErro = 195102 Then gError 30329

        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ProdutoOP_Col), sProdutoOPFormatado, iProdutoOPPreenchido)
        If lErro <> SUCESSO Then gError 19489

        If iProdutoOPPreenchido = PRODUTO_PREENCHIDO Then

            objItemOP.iFilialEmpresa = giFilialEmpresa
            objItemOP.sCodigo = OPCodigo.Text
            objItemOP.sServico = sProdutoOPFormatado

            lErro = CF("ItemOS_Le", objItemOP)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 19488

            If lErro = ERRO_LEITURA_SEM_DADOS Then gError 19487

        End If

    Else

        ProdutoOP.PromptInclude = False
        ProdutoOP.Text = ""
        ProdutoOP.PromptInclude = True

        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ProdutoOP_Col) = ""

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30330

    Saida_Celula_OP = SUCESSO

    Exit Function

Erro_Saida_Celula_OP:

    Saida_Celula_OP = gErr

    Select Case gErr

        Case 19487
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PARTICIPA_OS", gErr, objItemOP.sServico, objItemOP.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 19488, 19489, 30328, 30330
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 30329
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_OSCODIGO_NAO_CADASTRADO", objOrdemProducao.sCodigo)

            If vbMsg = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("OrdemProducao", objOrdemProducao)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162842)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ProdutoOP(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objItemOP As New ClassItemOS
Dim objKit As New ClassKit
Dim iProdutoPreenchido As Integer
Dim iProdutoOPPreenchido As Integer
Dim sProduto As String
Dim sProdutoOP As String
Dim sProdutoFormatado As String
Dim sProdutoOPFormatado As String
Dim iAchou As Integer, iIndice As Integer
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_ProdutoOP

    iAchou = 0

    Set objGridInt.objControle = ProdutoOP
    
    If Len(Trim(ProdutoOP.ClipText)) > 0 Then

        lErro = CF("Produto_Critica", ProdutoOP.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25077 Then gError 30331

        If lErro = 25077 Then gError 30332

        If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_OP_Col))) > 0 Then

            objItemOP.sCodigo = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_OP_Col)
            objItemOP.iFilialEmpresa = giFilialEmpresa

            objItemOP.sServico = objProduto.sCodigo
            
            'Lê o Item da Ordem de Produção
            lErro = CF("ItemOS_Le", objItemOP)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 19491
            
            'Se não encontrou o Item --> Erro
            If lErro = ERRO_LEITURA_SEM_DADOS Then gError 19490

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30333

    Saida_Celula_ProdutoOP = SUCESSO

    Exit Function

Erro_Saida_Celula_ProdutoOP:

    Saida_Celula_ProdutoOP = gErr

    Select Case gErr

        Case 19490
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PARTICIPA_OS", gErr, objItemOP.sServico, objItemOP.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 19491, 30331, 30333
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 30332
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", ProdutoOP.Text)

            If vbMsg = vbYes Then
                objProduto.sCodigo = ProdutoOP.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Produto", objProduto)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162843)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim iAchou As Integer
Dim iIndice As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iTipoMov As Integer
Dim objItemOP As ClassItemOS
Dim objMovEstoque As New ClassMovEstoque
Dim iEstorno As Integer
Dim iLinha As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o código foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 30336

    'Verifica se a Data foi preenchida
    If Len(Data.ClipText) = 0 Then gError 30337

    'Verifica se há Algum Ítem de Movimento de Estoque Informado no GridMovimentos
    If objGrid.iLinhasExistentes = 0 Then gError 30338

    'Para cada MovEstoque
    For iIndice = 1 To objGrid.iLinhasExistentes

        'Verifica se a Unidade de Medida foi preenchida
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_UnidadeMed_Col))) = 0 Then gError 55381

        If Codigo_Extrai(GridMovimentos.TextMatrix(iIndice, iGrid_TipoMov_Col)) >= 300 And Codigo_Extrai(GridMovimentos.TextMatrix(iIndice, iGrid_TipoMov_Col)) <= 304 Then
            'Verifica se a Quantidade foi informada
            If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Custo_Col))) = 0 Then gError 132079
        Else
            'Verifica se a Quantidade foi informada
            If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 30339
        End If
        
        'Verifica se o Almoxarifado foi informado
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Almoxarifado_Col))) = 0 Then gError 30340

        'Verifica se TipoMov foi Preenchido
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_TipoMov_Col))) = 0 Then gError 30343

        'Verifica se TipoMov está ligado a Produção
        iTipoMov = Codigo_Extrai(GridMovimentos.TextMatrix(iIndice, iGrid_TipoMov_Col))

        'Verifica se OPCodigo foi preenchido
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_OP_Col))) = 0 Then gError 30359

        iProdutoPreenchido = 0
        sProdutoFormatado = ""

        'Verifica se ProdutoOP foi preenchido
        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(iIndice, iGrid_ProdutoOP_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 30360

        If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 30361

        Set objItemOP = New ClassItemOS

        objItemOP.sCodigo = GridMovimentos.TextMatrix(iIndice, iGrid_OP_Col)
        objItemOP.iFilialEmpresa = giFilialEmpresa
        objItemOP.sServico = sProdutoFormatado

        'Verifica se Produto faz parte da Ordem de Produção ==>>
        lErro = CF("ItemOS_Le", objItemOP)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 30395

        '==>> Se não fizer Erro!
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 30402

    Next

    objMovEstoque.lCodigo = CLng(Codigo.Text)
    objMovEstoque.iFilialEmpresa = giFilialEmpresa

    lErro = CF("MovEstoque_Le", objMovEstoque)
    If lErro <> SUCESSO And lErro <> 30128 Then gError 30880

    If lErro = SUCESSO Then
        
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_MOVIMENTO_ESTOQUE_ALTERACAO_CAMPOS2")
        If vbMsgRes = vbNo Then gError 78737
    
    End If

    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(Data.Text))
    If lErro <> SUCESSO Then gError 92026

    'Tipo de Movimento só deve estar preenchido a nível de Itens
    objMovEstoque.iTipoMov = 0

    lErro = Move_Tela_Memoria(objMovEstoque)
    If lErro <> SUCESSO Then gError 30344

    'trata os casos de estorno da versao light
    Call MovEstoque_Trata_Estorno_Versao_Light

    'Grava no BD (inclusive os dados contabeis) (contabilidade)
    lErro = CF("MovEstoque_Interno_Grava", objMovEstoque, objContabil)
    If lErro <> SUCESSO Then gError 30345

    'gravar anotacao, se houver
    If Not (gobjAnotacao Is Nothing) Then
    
        If Len(Trim(gobjAnotacao.sTextoCompleto)) <> 0 Or Len(Trim(gobjAnotacao.sTitulo)) <> 0 Then
        
            gobjAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_MOVESTOQUE
            gobjAnotacao.sID = CStr(objMovEstoque.iFilialEmpresa) & "," & CStr(objMovEstoque.lCodigo)
            gobjAnotacao.dtDataAlteracao = gdtDataHoje
            
            lErro = CF("Anotacoes_Grava", gobjAnotacao)
            If lErro <> SUCESSO Then gError 30345
            
        End If
        
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 30181
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTORNO_MOVTO_ESTOQUE_NAO_CADASTRADO", gErr, objMovEstoque.lCodigo)

        Case 30182
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVTO_ESTOQUE_CADASTRADO", gErr, objMovEstoque.lCodigo)

        Case 30295
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTORNO_ITEM_NAO_CADASTRADO", gErr, iIndice)

        Case 30336
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 30337
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)

        Case 30338
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVESTOQUE_NAO_INFORMADO", gErr)

        Case 30339
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA", gErr, iIndice)

        Case 30340
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_PREENCHIDO", gErr, iIndice)

        Case 30343
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMOV_NAO_PREENCHIDO", gErr, iIndice)

        Case 30359
            Call Rotina_Erro(vbOKOnly, "ERRO_OS_NAO_PREENCHIDO_GRID", gErr, iIndice)

        Case 30361
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTOOP_NAO_PREENCHIDO", gErr, iIndice)

        Case 30344, 30345, 30360, 30395, 30880, 78737, 92026

        Case 30402
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_FORA_OS", gErr, GridMovimentos.TextMatrix(iIndice, iGrid_ProdutoOP_Col), iIndice, GridMovimentos.TextMatrix(iIndice, iGrid_OP_Col))

        Case 30881
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVTO_ESTOQUE_CADASTRADO", gErr, objMovEstoque.lCodigo)

        Case 55381
            Call Rotina_Erro(vbOKOnly, "ERRO_UM_NAO_PREENCHIDA", gErr, iIndice)

        Case 132079
            Call Rotina_Erro(vbOKOnly, "ERRO_CUSTO_GRID_NAO_PREENCHIDO", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162844)

    End Select

    Exit Function

End Function

Private Sub Limpa_Tela_MovEstoque()

Dim lErro As Long, lCodigo As Long
On Error GoTo Erro_Limpa_Tela_MovEstoque

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Função genérica para Limpar a Tela
    Call Limpa_Tela(Me)

    'Limpa o Label QuantDisponivel
    QuantDisponivel.Caption = ""

    'prepara o Grid para permitir inserir e excluir Linhas
    objGrid.iProibidoIncluir = 0
    objGrid.iProibidoExcluir = 0
    Call Grid_Inicializa(objGrid)
    
    'Limpa o Grid
    Call Grid_Limpa(objGrid)

    'Remove os ítens de colItensNumIntDoc
    Set colItensNumIntDoc = New Collection

    Codigo.PromptInclude = False
    Codigo.Text = ""
    Codigo.PromptInclude = True
        
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True

    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

    iAlterado = 0
    lCodigoAntigo = 0
    
    Set gobjAnotacao = Nothing

    Set gcolcolRastreamentoSerie = New Collection

    Exit Sub

Erro_Limpa_Tela_MovEstoque:

    Select Case gErr

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162845)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_GridMovimentos(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Tipo Movimento")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Custo")
    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("Ccl")
    objGridInt.colColuna.Add ("OS")
    objGridInt.colColuna.Add ("Serviço")
    objGridInt.colColuna.Add ("Lote\OP\Série")
    objGridInt.colColuna.Add ("Filial OP")
    objGridInt.colColuna.Add ("Conta Débito")
    objGridInt.colColuna.Add ("Conta Crédito")
    'objGridInt.colColuna.Add ("Estorno")
        
    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (Descricao.Name)
    objGridInt.colCampo.Add (UnidadeMed.Name)
    objGridInt.colCampo.Add (TipoMov.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (Custo.Name)
    objGridInt.colCampo.Add (Almoxarifado.Name)
    objGridInt.colCampo.Add (Ccl.Name)
    objGridInt.colCampo.Add (OPCodigo.Name)
    objGridInt.colCampo.Add (ProdutoOP.Name)
    objGridInt.colCampo.Add (Lote.Name)
    objGridInt.colCampo.Add (FilialOP.Name)
    objGridInt.colCampo.Add (ContaContabilDebito.Name)
    objGridInt.colCampo.Add (ContaContabilCredito.Name)
    'objGridInt.colCampo.Add (Estorno.Name)
    
    'Colunas do Grid
    iGrid_Sequencial_Col = 0
    iGrid_Produto_Col = 1
    iGrid_Descricao_Col = 2
    iGrid_UnidadeMed_Col = 3
    
    iGrid_TipoMov_Col = 4
    iGrid_Quantidade_Col = 5
    iGrid_Custo_Col = 6
    iGrid_Almoxarifado_Col = 7
    iGrid_Ccl_Col = 8
    iGrid_OP_Col = 9
    iGrid_ProdutoOP_Col = 10
    iGrid_Lote_Col = 11
    iGrid_FilialOP_Col = 12
    iGrid_ContaContabilDebito_Col = 13
    iGrid_ContaContabilCredito_Col = 14
    'iGrid_Estorno_Col = 15
    
    Estorno.left = POSICAO_FORA_TELA
    
    'Grid do GridInterno
    objGridInt.objGrid = GridMovimentos

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE + 1

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5

    'Largura da primeira coluna
    GridMovimentos.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridMovimentos = SUCESSO

    Exit Function

End Function

Private Sub GridMovimentos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Private Sub GridMovimentos_EnterCell()

    Call Grid_Entrada_Celula(objGrid, iAlterado)

End Sub

Private Sub GridMovimentos_GotFocus()

    Call Grid_Recebe_Foco(objGrid)

End Sub

Private Sub GridMovimentos_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer
Dim iLinhaAnterior As Integer, lNumIntDoc As Long

    If colItensNumIntDoc.Count >= GridMovimentos.Row Then
        lNumIntDoc = colItensNumIntDoc.Item(GridMovimentos.Row)
    Else
        lNumIntDoc = 0
    End If

    If lNumIntDoc = 0 Then

        'Verifica se a Tecla apertada foi Del
        If KeyCode = vbKeyDelete Then
        
            'Guarda iLinhasExistentes
            iLinhasExistentesAnterior = objGrid.iLinhasExistentes

            'Guarda o índice da Linha a ser Excluída
            iLinhaAnterior = GridMovimentos.Row

        End If

        Call Grid_Trata_Tecla1(KeyCode, objGrid)

        'Verifica se a Linha foi realmente excluída
        If objGrid.iLinhasExistentes < iLinhasExistentesAnterior Then
        
            'Exclui de colItensNumIntDoc o Item correspondente, se houver
            colItensNumIntDoc.Remove iLinhaAnterior
            gcolcolRastreamentoSerie.Remove iLinhaAnterior

        End If

    End If

End Sub

Private Sub GridMovimentos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If


End Sub

Private Sub GridMovimentos_LeaveCell()

    Call Saida_Celula(objGrid)

End Sub

Private Sub GridMovimentos_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGrid)
    
End Sub

Private Sub GridMovimentos_Scroll()

    Call Grid_Scroll(objGrid)

End Sub

Private Sub Almoxarifado_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Almoxarifado_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Almoxarifado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Almoxarifado
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Ccl_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Ccl_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Ccl_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Ccl
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Estorno_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Estorno_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Estorno_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Estorno
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub OPCodigo_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub OPCodigo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub OPCodigo_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGrid.objControle = OPCodigo
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Produto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ProdutoOP_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub ProdutoOP_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub ProdutoOP_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = ProdutoOP
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Quantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UnidadeMed_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub UnidadeMed_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub UnidadeMed_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = UnidadeMed
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub TipoMov_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub TipoMov_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub TipoMov_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = TipoMov
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_UnidadeMed(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iEntradaSaida As Integer
Dim dQuantidade As Double
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Saida_Celula_UnidadeMed

    Set objGridInt.objControle = UnidadeMed

    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col) = UnidadeMed.Text

    If Len(UnidadeMed.Text) > 0 Then

        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 141949
            
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 141950

        If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col))) = 0 Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then  'Alterado por Wagner 15/03/2006
            lErro = QuantDisponivel_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col))
            If lErro <> SUCESSO Then gError 55347
        Else
            lErro = QuantDisponivel_Lote_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col)))
            If lErro <> SUCESSO Then gError 78688
        End If

        If colItensNumIntDoc.Item(GridMovimentos.Row) = 0 Then
            
            'Testa se é movimento de entrada ou saída
            Call Testa_TipoMovimento(GridMovimentos.Row, iEntradaSaida)
            
            If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))) > 0 And iEntradaSaida = ESTOQUE_SAIDA Then
    
                dQuantidade = CDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))
    
                lErro = Testa_QuantRequisitada(dQuantidade)
                If lErro <> SUCESSO Then gError 55348
    
            End If
    
        End If

    Else
    
        QuantDisponivel.Caption = ""
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30353

    Saida_Celula_UnidadeMed = SUCESSO

    Exit Function

Erro_Saida_Celula_UnidadeMed:

    Saida_Celula_UnidadeMed = gErr

    Select Case gErr

        Case 30353, 55347, 55348, 78688, 141949, 141950
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162846)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Estorno(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Estorno

    Set objGridInt.objControle = Estorno

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30355

    Saida_Celula_Estorno = SUCESSO

    Exit Function

Erro_Saida_Celula_Estorno:

    Saida_Celula_Estorno = gErr

    Select Case gErr

        Case 30355
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162847)

    End Select

    Exit Function

End Function

Function Preenche_Tela(objMovEstoque As ClassMovEstoque) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Preenche_Tela

    'Limpa a tela sem Fechar o Comando de setas
    'Função genérica para Limpar a Tela
    Call Limpa_Tela(Me)

    'Limpa o Grid
    Call Grid_Limpa(objGrid)

    'Se o grid permite excluir e incluir Linhas
    If objGrid.iProibidoIncluir <> GRID_PROIBIDO_INCLUIR And objGrid.iProibidoExcluir <> GRID_PROIBIDO_EXCLUIR Then
        'prepara o Grid para não permitir inserir e excluir Linhas
        objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
        objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
        Call Grid_Inicializa(objGrid)
    End If
    
    'Limpa o Label QuantDisponivel
    QuantDisponivel.Caption = ""

    'Remove os ítens de colItensNumIntDoc
    Set colItensNumIntDoc = New Collection
    Set objMovEstoque.colItens = New ColItensMovEstoque

    'Lê os ítens do Movimento de Estoque
    lErro = CF("MovEstoqueInterno2Itens_Le", objMovEstoque)
    If lErro <> SUCESSO And lErro <> 30565 Then gError 30356

    If lErro = 30565 Then gError 55440

    'Passa as Informações de NumIntDoc de colItens para colItensNumIntDoc
    For iIndice = 1 To objMovEstoque.colItens.Count

        colItensNumIntDoc.Add objMovEstoque.colItens(iIndice).lNumIntDoc

    Next

    'Coloca os Dados na Tela
    Codigo.PromptInclude = False
    Codigo.Text = CStr(objMovEstoque.lCodigo)
    Codigo.PromptInclude = True

    If objMovEstoque.dtData <> DATA_NULA Then
        Data.PromptInclude = False
        Data.Text = Format(objMovEstoque.dtData, "dd/mm/yy")
        Data.PromptInclude = True
    Else
        Data.PromptInclude = False
        Data.Text = ""
        Data.PromptInclude = True
    End If
    
    Observacao.Text = objMovEstoque.sObservacao

'hora
    Hora.PromptInclude = False
    'este teste está correto
    If objMovEstoque.dtData <> DATA_NULA Then Hora.Text = Format(objMovEstoque.dtHora, "hh:mm:ss")
    Hora.PromptInclude = True

    lErro = Preenche_GridMovimentos(objMovEstoque.colItens)
    If lErro <> SUCESSO Then gError 30357

    'traz os dados contábeis para a tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela(objMovEstoque.colItens(1).lNumIntDoc)
    If lErro <> SUCESSO And lErro <> 36326 Then gError 39627

    iAlterado = 0
    lCodigoAntigo = objMovEstoque.lCodigo

    Preenche_Tela = SUCESSO

    Exit Function

Erro_Preenche_Tela:

    Preenche_Tela = gErr

    Select Case gErr

        Case 30356, 30357, 39627

        Case 55440
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVESTOQUE_NAO_CADASTRADO", gErr, giFilialEmpresa, objMovEstoque.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162848)

    End Select

    Exit Function

End Function

Private Function Preenche_GridMovimentos(colItens As ColItensMovEstoque) As Long

Dim iIndice As Integer
Dim sProdutoMascarado As String, sCclMascarado As String
Dim lErro As Long
Dim sProdutoOPMascarado As String
Dim objTipoMovEst As ClassTipoMovEst
Dim objItemMovEstoque As ClassItemMovEstoque
Dim sContaEnxutaEst As String
Dim sContaEnxutaAplic As String
Dim colRatreamentoMovto As New Collection
Dim objRatreamentoMovto As New ClassRastreamentoMovto
Dim objFilialOP As New AdmFiliais
Dim colRastreamentoSerie As Collection

On Error GoTo Erro_Preenche_GridMovimentos

    Set gcolcolRastreamentoSerie = New Collection

    'Preenche GridMovimentos
    For Each objItemMovEstoque In colItens

        iIndice = iIndice + 1

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        lErro = Mascara_RetornaProdutoTela(objItemMovEstoque.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 55482

        Produto.PromptInclude = False
        Produto.Text = sProdutoMascarado
        Produto.PromptInclude = True
        
        'preenche conta Est no grid
        If objItemMovEstoque.sContaContabilEst <> "" Then
            
            sContaEnxutaEst = String(STRING_CONTA, 0)
        
            lErro = Mascara_RetornaContaEnxuta(objItemMovEstoque.sContaContabilEst, sContaEnxutaEst)
            If lErro <> SUCESSO Then gError 17724
        
            ContaContabilCredito.PromptInclude = False
            ContaContabilCredito.Text = sContaEnxutaEst
            ContaContabilCredito.PromptInclude = True
            
            GridMovimentos.TextMatrix(iIndice, iGrid_ContaContabilCredito_Col) = ContaContabilCredito.Text
            
        End If
        
        'Preencher ContaAplic no grid
        If objItemMovEstoque.sContaContabilAplic <> "" Then
        
            sContaEnxutaAplic = String(STRING_CONTA, 0)
        
            lErro = Mascara_RetornaContaEnxuta(objItemMovEstoque.sContaContabilAplic, sContaEnxutaAplic)
            If lErro <> SUCESSO Then gError 17725
            
            ContaContabilDebito.PromptInclude = False
            ContaContabilDebito.Text = sContaEnxutaAplic
            ContaContabilDebito.PromptInclude = True
            
            GridMovimentos.TextMatrix(iIndice, iGrid_ContaContabilDebito_Col) = ContaContabilDebito.Text
        
        End If
        
        GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col) = sProdutoMascarado
        GridMovimentos.TextMatrix(iIndice, iGrid_Descricao_Col) = objItemMovEstoque.sProdutoDesc
        GridMovimentos.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItemMovEstoque.sSiglaUM
        GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemMovEstoque.dQuantidade)
        GridMovimentos.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = objItemMovEstoque.sAlmoxarifadoNomeRed

        If objItemMovEstoque.dCusto <> 0 Then GridMovimentos.TextMatrix(iIndice, iGrid_Custo_Col) = Format(objItemMovEstoque.dCusto, Custo.Format)

        If objItemMovEstoque.sCcl <> "" Then

            sCclMascarado = String(STRING_CCL, 0)

            lErro = Mascara_MascararCcl(objItemMovEstoque.sCcl, sCclMascarado)
            If lErro <> SUCESSO Then gError 22918

        Else

            sCclMascarado = ""

        End If

        GridMovimentos.TextMatrix(iIndice, iGrid_Ccl_Col) = sCclMascarado

        'Usa o colTipoMovEst para descobrir a descricao do TipoMov
        For Each objTipoMovEst In colTipoMovEst

            If objTipoMovEst.iCodigo = objItemMovEstoque.iTipoMov Then

                GridMovimentos.TextMatrix(iIndice, iGrid_TipoMov_Col) = CStr(objTipoMovEst.iCodigo) & "-" & objTipoMovEst.sDescricao
                Exit For

            End If

        Next

        If objItemMovEstoque.sProdutoOP <> "" Then
            sProdutoOPMascarado = String(STRING_PRODUTO, 0)

            lErro = Mascara_MascararProduto(objItemMovEstoque.sProdutoOP, sProdutoOPMascarado)
            If lErro <> SUCESSO Then gError 19473
        Else
            sProdutoOPMascarado = ""
        End If

        ProdutoOP.PromptInclude = False
        ProdutoOP.Text = sProdutoOPMascarado
        ProdutoOP.PromptInclude = True

        'Preenche as Colunas ligadas a Ordem de Produção
        GridMovimentos.TextMatrix(iIndice, iGrid_ProdutoOP_Col) = ProdutoOP.Text
        GridMovimentos.TextMatrix(iIndice, iGrid_OP_Col) = objItemMovEstoque.sOPCodigo

'        If objItemMovEstoque.lNumIntDocEst = 0 Then
'            GridMovimentos.TextMatrix(iIndice, iGrid_Estorno_Col) = "0"
'        Else
'            GridMovimentos.TextMatrix(iIndice, iGrid_Estorno_Col) = "1"
'        End If
        
        Set colRatreamentoMovto = New Collection
        
        'Le o Rastreamento e preenche o grid com o Número do Lote e o Numero da Filial OP
        lErro = CF("RastreamentoMovto_Le_DocOrigem", objItemMovEstoque.lNumIntDoc, TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE, colRatreamentoMovto)
        If lErro <> SUCESSO And lErro <> 78414 Then gError 78610
        
        'Se existe rastreamento
        If colRatreamentoMovto.Count > 0 Then
                        
            'Seta o primeiro Lote
            Set objRatreamentoMovto = colRatreamentoMovto(1)

            gcolcolRastreamentoSerie.Add objRatreamentoMovto.colRastreamentoSerie 'Inserido por Wagner 15/03/2006

            If Len(Trim(objRatreamentoMovto.sLote)) > 0 Then GridMovimentos.TextMatrix(iIndice, iGrid_Lote_Col) = objRatreamentoMovto.sLote
            
            If objRatreamentoMovto.iFilialOP > 0 Then
            
                objFilialOP.iCodFilial = objRatreamentoMovto.iFilialOP

                'Le a Filial Empresa da OP para pegar a descrição
                lErro = CF("FilialEmpresa_Le", objFilialOP)
                If lErro <> SUCESSO Then gError 78736

                GridMovimentos.TextMatrix(iIndice, iGrid_FilialOP_Col) = objFilialOP.iCodFilial & SEPARADOR & objFilialOP.sNome
            
            End If
            
        Else
            Set colRastreamentoSerie = New Collection
            gcolcolRastreamentoSerie.Add colRastreamentoSerie
            
        End If

    Next

    objGrid.iLinhasExistentes = colItens.Count

    lErro = Grid_Refresh_Checkbox(objGrid)
    If lErro <> SUCESSO Then gError 30358

    Preenche_GridMovimentos = SUCESSO

    Exit Function

Erro_Preenche_GridMovimentos:

    Preenche_GridMovimentos = gErr

    Select Case gErr

        Case 19473
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objItemMovEstoque.sProdutoOP)

        Case 30358, 22918, 17724, 17725, 78610, 78736

        Case 55482
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objItemMovEstoque.sProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162849)

    End Select

    Exit Function

End Function

Public Sub OP_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ProdutoOPGera_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub OP_Validate(bMantemFoco As Boolean)

Dim lErro As Long
Dim objOrdemProducao As New ClassOS
Dim objItemOP As New ClassItemOS
Dim vbMsg As VbMsgBoxResult
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_OP_Validate

    'Verifica se foi Preenchido
    If Len(Trim(OP.Text)) > 0 Then

        objOrdemProducao.sCodigo = OP.Text
        objOrdemProducao.iFilialEmpresa = giFilialEmpresa

        lErro = CF("OS_Le", objOrdemProducao)
        If lErro <> SUCESSO And lErro <> 195102 Then gError 34748

        If lErro = 195102 Then gError 206265
        
        lErro = CF("Produto_Formata", ProdutoOPGera.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 34750

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            objItemOP.iFilialEmpresa = giFilialEmpresa
            objItemOP.sCodigo = OP.Text
            objItemOP.sServico = sProdutoFormatado

            lErro = CF("ItemOS_Le", objItemOP)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 34751

            If lErro = ERRO_LEITURA_SEM_DADOS Then gError 206266

        End If

    End If

    Exit Sub

Erro_OP_Validate:

    bMantemFoco = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 206265
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_OSCODIGO_NAO_CADASTRADO", objOrdemProducao.sCodigo)

            If vbMsg = vbYes Then
                Call Chama_Tela("OrdemProducao", objOrdemProducao)
            End If

        Case 206266
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PARTICIPA_OS", gErr, objItemOP.sServico, objItemOP.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206267)

    End Select

    Exit Sub

End Sub

Public Sub ProdutoOPGera_Validate(bMantemFoco As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objItemOP As New ClassItemOS
Dim iProdutoPreenchido As Integer
Dim sProduto As String
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_ProdutoOPGera_Validate

    If Len(Trim(ProdutoOPGera.ClipText)) > 0 Then

        sProduto = ProdutoOPGera.Text

        lErro = CF("Produto_Critica", sProduto, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 Then gError ERRO_SEM_MENSAGEM
        
        If lErro = 25041 Then gError 206268
              
        If Len(Trim(OP.Text)) > 0 Then

            objItemOP.iFilialEmpresa = giFilialEmpresa
            objItemOP.sCodigo = OP.Text
            objItemOP.sServico = objProduto.sCodigo

            lErro = CF("ItemOS_Le", objItemOP)
            If lErro <> SUCESSO And lErro <> 34711 Then gError ERRO_SEM_MENSAGEM
            If lErro = 34711 Then gError 206270

        End If

    End If

    Exit Sub

Erro_ProdutoOPGera_Validate:

    bMantemFoco = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 206268
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", sProduto)

             If vbMsg = vbYes Then
                objProduto.sCodigo = sProduto
                Call Chama_Tela("Produto", objProduto)
            End If
    
        Case 206269
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PCP", gErr)

        Case 206270
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PARTICIPA_OS", gErr, objItemOP.sServico, objItemOP.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206271)

    End Select

    Exit Sub

End Sub

Private Function ProdutoLinha_Preenche(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iCclPreenchida As Integer
Dim sCclFormata As String
Dim sAlmoxarifadoPadrao As String
Dim sContaEnxutaAplic As String
Dim colRastreamentoSerie As New Collection

On Error GoTo Erro_ProdutoLinha_Preenche
    
    'Unidade de Medida
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMEstoque

    'Descricao
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Descricao_Col) = objProduto.sDescricao

    'Almoxarifado
    '(Utiliza Almoxarifado Padrão caso esteja preenchido) senão preenche com o Almoxarifado Padrão do produto ou da Filial
    If Len(Trim(AlmoxPadrao.Text)) <> 0 Then
        lErro = CF("EstoqueProduto_TestaAssociacao", Produto.Text, AlmoxPadrao)
        If lErro = SUCESSO Then
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col) = AlmoxPadrao.Text
        Else
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col) = ""
        End If
    Else
        
        'le o Nome reduzido do almoxarifado Padrão do Produto em Questão
        lErro = CF("AlmoxarifadoPadrao_Le_NomeReduzido", objProduto.sCodigo, sAlmoxarifadoPadrao)
        If lErro <> SUCESSO Then gError 52224
        
        'preenche o grid
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col) = sAlmoxarifadoPadrao
    
    End If
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_OP_Col) = OP.Text
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ProdutoOP_Col) = ProdutoOPGera.Text
        
    'Ccl
    lErro = CF("Ccl_Formata", CclPadrao.Text, sCclFormata, iCclPreenchida)
    If lErro <> SUCESSO Then gError 30168

    If iCclPreenchida = CCL_PREENCHIDA Then GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Ccl_Col) = CclPadrao.Text

    'Preenche Estorno com Valor 0 (Checked = False)
'    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Estorno_Col) = "0"

    'ALTERAÇÃO DE LINHAS EXISTENTES
    If (GridMovimentos.Row - GridMovimentos.FixedRows) = objGrid.iLinhasExistentes Then
        objGrid.iLinhasExistentes = objGrid.iLinhasExistentes + 1
        gcolcolRastreamentoSerie.Add colRastreamentoSerie 'Inserido por Wagner 15/03/2006
        colItensNumIntDoc.Add 0
    End If

    ProdutoLinha_Preenche = SUCESSO

    Exit Function

Erro_ProdutoLinha_Preenche:

    ProdutoLinha_Preenche = gErr

    Select Case gErr

        Case 30165, 30166, 30168, 52224
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162850)

    End Select

    Exit Function

End Function

Function Move_Tela_Memoria(objMovEstoque As ClassMovEstoque) As Long
'Preenche objMovEstoque (inclusive colItens)

Dim iIndice As Integer
Dim lCodigo As Long
Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    'Não preenche Tipo de Movimento a nível de Objeto
    If Len(Trim(Codigo.Text)) <> 0 Then objMovEstoque.lCodigo = CLng(Codigo.Text)

    If Len(Trim(Data.Text)) <> 0 Then
        objMovEstoque.dtData = CDate(Data.Text)

    Else
        objMovEstoque.dtData = DATA_NULA

    End If
    
    objMovEstoque.sObservacao = Observacao.Text

'hora
    If Len(Trim(Hora.ClipText)) > 0 Then
        objMovEstoque.dtHora = CDate(Hora.Text)
    Else
        objMovEstoque.dtHora = Time
    End If

    objMovEstoque.iFilialEmpresa = giFilialEmpresa

    For iIndice = 1 To objGrid.iLinhasExistentes

        lErro = Move_Itens_Memoria(iIndice, objMovEstoque)
        If lErro <> SUCESSO Then gError 30323

    Next

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 30323, 30324, 30913

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162851)

    End Select

    Exit Function

End Function

Function Move_Itens_Memoria(iIndice As Integer, objMovEstoque As ClassMovEstoque) As Long

Dim lErro As Long
Dim iTipoMov As Integer
Dim sProdutoFormatado As String
Dim sOPCodigo As String
Dim sProdutoFormatadoOP As String
Dim iProdutoPreenchido As Integer
Dim sCcl As String, sCclFormatada As String
Dim iCclPreenchida As Integer
Dim objAlmoxarifado As ClassAlmoxarifado
Dim iFilialEmpresa As Integer
Dim sContaFormatadaAplic As String
Dim sContaFormatadaEst As String
Dim iContaPreenchida As Integer
Dim objItemMovEstoque As ClassItemMovEstoque
Dim colRatreamentoMovto As New Collection
Dim objItemOP As ClassItemOS
Dim objOS As ClassOS
Dim objSolicSRV As ClassSolicSRV

On Error GoTo Erro_Move_Itens_Memoria

    With GridMovimentos

        'Pega o Código do Almoxarifado
        Set objAlmoxarifado = New ClassAlmoxarifado

        objAlmoxarifado.sNomeReduzido = .TextMatrix(iIndice, iGrid_Almoxarifado_Col)

        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25056 Then gError 30373

        If lErro = 25056 Then gError 30374

        sOPCodigo = ""
        sProdutoFormatadoOP = ""
        sProdutoFormatado = ""

        sCcl = .TextMatrix(iIndice, iGrid_Ccl_Col)

        If Len(Trim(sCcl)) <> 0 Then

            'Formata Ccl para BD
            lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
            If lErro <> SUCESSO Then gError 30375

        Else

            sCclFormatada = ""

        End If
        
        If .TextMatrix(iIndice, iGrid_ContaContabilCredito_Col) <> "" Then
        
            'Formata as Contas para o Bd
            lErro = CF("Conta_Formata", .TextMatrix(iIndice, iGrid_ContaContabilCredito_Col), sContaFormatadaEst, iContaPreenchida)
            If lErro <> SUCESSO Then gError 49657
        
        Else
            sContaFormatadaEst = ""
        End If
        
        If .TextMatrix(iIndice, iGrid_ContaContabilDebito_Col) <> "" Then
            
            lErro = CF("Conta_Formata", .TextMatrix(iIndice, iGrid_ContaContabilDebito_Col), sContaFormatadaAplic, iContaPreenchida)
            If lErro <> SUCESSO Then gError 49658
        
        Else
            sContaFormatadaAplic = ""
        End If
        
        'Formata o Produto para BD
        lErro = CF("Produto_Formata", .TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 30376

        'Preenche Tipo de Movimento para cada Item
        'Coloca o Tipo de Movimento em iTipoMov
        iTipoMov = Codigo_Extrai(.TextMatrix(iIndice, iGrid_TipoMov_Col))

        sOPCodigo = .TextMatrix(iIndice, iGrid_OP_Col)

        'Formata o ProdutoOP para o BD
        lErro = CF("Produto_Formata", .TextMatrix(iIndice, iGrid_ProdutoOP_Col), sProdutoFormatadoOP, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 30384

        iFilialEmpresa = objMovEstoque.iFilialEmpresa

        'Move os dados do rastreamento para a Memória
        lErro = Move_RastroEstoque_Memoria(iIndice, colRatreamentoMovto)
        If lErro <> SUCESSO Then gError 78315

        Set objItemMovEstoque = objMovEstoque.colItens.Add(colItensNumIntDoc(iIndice), iTipoMov, 0, 0, sProdutoFormatado, .TextMatrix(iIndice, iGrid_Descricao_Col), .TextMatrix(iIndice, iGrid_UnidadeMed_Col), StrParaDbl(.TextMatrix(iIndice, iGrid_Quantidade_Col)), objAlmoxarifado.iCodigo, .TextMatrix(iIndice, iGrid_Almoxarifado_Col), 0, sCclFormatada, 0, sOPCodigo, sProdutoFormatadoOP, sContaFormatadaAplic, sContaFormatadaEst, 0, colRatreamentoMovto, Nothing, DATA_NULA)
        
        objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMOS
        
        Set objItemOP = New ClassItemOS

        objItemOP.sCodigo = GridMovimentos.TextMatrix(iIndice, iGrid_OP_Col)
        objItemOP.iFilialEmpresa = giFilialEmpresa
        objItemOP.sServico = sProdutoFormatadoOP

        'Verifica se Produto faz parte da Ordem de Produção ==>>
        lErro = CF("ItemOS_Le", objItemOP)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 30395
        
        objItemMovEstoque.lNumIntDocOrigem = objItemOP.lNumIntDoc

        objItemMovEstoque.dCusto = StrParaDbl(.TextMatrix(iIndice, iGrid_Custo_Col))
        
        Set objOS = New ClassOS
        Set objSolicSRV = New ClassSolicSRV
        
        objOS.sCodigo = GridMovimentos.TextMatrix(iIndice, iGrid_OP_Col)
        objOS.iFilialEmpresa = giFilialEmpresa
        
        lErro = CF("OS_Le", objOS)
        If lErro <> SUCESSO And lErro <> 195102 Then gError 30395
        
        objSolicSRV.lCodigo = objOS.lCodSolSRV
        objSolicSRV.iFilialEmpresa = giFilialEmpresa
        
        lErro = CF("SolicitacaoSRV_Le", objSolicSRV)
        If lErro <> SUCESSO And lErro <> 195102 Then gError 30395
        
        objItemMovEstoque.lCliente = objSolicSRV.lCliente

    End With

    Move_Itens_Memoria = SUCESSO

    Exit Function

Erro_Move_Itens_Memoria:

    Move_Itens_Memoria = gErr

    Select Case gErr

        Case 30373, 30375, 30376, 30384, 49657, 49658, 78315

        Case 30374
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162852)

    End Select

    Exit Function

End Function

Function Move_RastroEstoque_Memoria(iLinha As Integer, colRastreamentoMovto As Collection) As Long
'Move o Rastro dos Itens de Movimento

Dim objProduto As New ClassProduto, lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objRastreamentoMovto As New ClassRastreamentoMovto

On Error GoTo Erro_Move_RastroEstoque_Memoria
    
    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 78311
    
    objProduto.sCodigo = sProdutoFormatado
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 78312

    If lErro = 28030 Then gError 78313
    
    If objProduto.iRastro <> PRODUTO_RASTRO_NENHUM Then
    
        If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
            
            'Se colocou o Número do Lote
            If Len(Trim(GridMovimentos.TextMatrix(iLinha, iGrid_Lote_Col))) <> 0 Then
            
                objRastreamentoMovto.sLote = GridMovimentos.TextMatrix(iLinha, iGrid_Lote_Col)
                
            End If
            
        ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
            
            'Se colocou o Número do lote e não colocou a filial ==> erro
            If Len(Trim(GridMovimentos.TextMatrix(iLinha, iGrid_Lote_Col))) <> 0 Then
            
                If Len(Trim(GridMovimentos.TextMatrix(iLinha, iGrid_FilialOP_Col))) = 0 Then gError 78734
                
                objRastreamentoMovto.sLote = GridMovimentos.TextMatrix(iLinha, iGrid_Lote_Col)
                objRastreamentoMovto.iFilialOP = Codigo_Extrai(GridMovimentos.TextMatrix(iLinha, iGrid_FilialOP_Col))
            End If
                                    
            'Se preencheu a filial e não colocou o Número do lote  ==> erro
            If Len(Trim(GridMovimentos.TextMatrix(iLinha, iGrid_FilialOP_Col))) <> 0 Then
                                    
                If Len(Trim(GridMovimentos.TextMatrix(iLinha, iGrid_Lote_Col))) = 0 Then gError 78314
            
            End If
            
        ElseIf objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
           
            For Each objRastreamentoMovto In gcolcolRastreamentoSerie.Item(iLinha)
                objRastreamentoMovto.iTipoDocOrigem = TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE
                colRastreamentoMovto.Add objRastreamentoMovto
            Next
            
        End If
        
        If Len(Trim(GridMovimentos.TextMatrix(iLinha, iGrid_Lote_Col))) <> 0 Then
        
            If Len(Trim(GridMovimentos.TextMatrix(iLinha, iGrid_Quantidade_Col))) > 0 Then objRastreamentoMovto.dQuantidade = CDbl(GridMovimentos.TextMatrix(iLinha, iGrid_Quantidade_Col))
            objRastreamentoMovto.iTipoDocOrigem = TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE
            objRastreamentoMovto.sProduto = sProdutoFormatado
                
            If objProduto.iRastro <> PRODUTO_RASTRO_NUM_SERIE Then
                colRastreamentoMovto.Add objRastreamentoMovto
            End If
        
        End If
        
    End If
    
    Move_RastroEstoque_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_RastroEstoque_Memoria:

    Move_RastroEstoque_Memoria = gErr
    
    Select Case gErr
        
        Case 78311, 78312
        
        Case 78313
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 78314
            Call Rotina_Erro(vbOKOnly, "ERRO_LOTE_RASTREAMENTO_NAO_PREENCHIDO", gErr, iLinha)
        
        Case 78733
            Call Rotina_Erro(vbOKOnly, "ERRO_OS_NAO_PREENCHIDO_GRID", gErr, iLinha)
        
        Case 78734
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_OP_NAO_PREENCHIDA", gErr, iLinha)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162853)
    
    End Select
    
    Exit Function
    
End Function

'inicio contabilidade
Private Sub CTBBotaoModeloPadrao_Click()

    Call objContabil.Contabil_BotaoModeloPadrao_Click

End Sub

Private Sub CTBModelo_Click()

    Call objContabil.Contabil_Modelo_Click

End Sub

Private Sub CTBGridContabil_Click()

    Call objContabil.Contabil_GridContabil_Click

    If giTipoVersao = VERSAO_LIGHT Then
        Call objContabil.Contabil_GridContabil_Consulta_Click
    End If

End Sub

Private Sub CTBGridContabil_EnterCell()

    Call objContabil.Contabil_GridContabil_EnterCell

End Sub

Private Sub CTBGridContabil_GotFocus()

    Call objContabil.Contabil_GridContabil_GotFocus

End Sub

Private Sub CTBGridContabil_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_GridContabil_KeyPress(KeyAscii)

End Sub

Private Sub CTBGridContabil_KeyDown(KeyCode As Integer, Shift As Integer)

    Call objContabil.Contabil_GridContabil_KeyDown(KeyCode)

End Sub


Private Sub CTBGridContabil_LeaveCell()

        Call objContabil.Contabil_GridContabil_LeaveCell

End Sub

Private Sub CTBGridContabil_Validate(Cancel As Boolean)

    Call objContabil.Contabil_GridContabil_Validate(Cancel)

End Sub

Private Sub CTBGridContabil_RowColChange()

    Call objContabil.Contabil_GridContabil_RowColChange

End Sub

Private Sub CTBGridContabil_Scroll()

    Call objContabil.Contabil_GridContabil_Scroll

End Sub

Private Sub CTBConta_Change()

    Call objContabil.Contabil_Conta_Change

End Sub

Private Sub CTBConta_GotFocus()

    Call objContabil.Contabil_Conta_GotFocus

End Sub

Private Sub CTBConta_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Conta_KeyPress(KeyAscii)

End Sub

Private Sub CTBConta_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Conta_Validate(Cancel)

End Sub

Private Sub CTBCcl_Change()

    Call objContabil.Contabil_Ccl_Change

End Sub

Private Sub CTBCcl_GotFocus()

    Call objContabil.Contabil_Ccl_GotFocus

End Sub

Private Sub CTBCcl_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Ccl_KeyPress(KeyAscii)

End Sub

Private Sub CTBCcl_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Ccl_Validate(Cancel)

End Sub

Private Sub CTBCredito_Change()

    Call objContabil.Contabil_Credito_Change

End Sub

Private Sub CTBCredito_GotFocus()

    Call objContabil.Contabil_Credito_GotFocus

End Sub

Private Sub CTBCredito_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Credito_KeyPress(KeyAscii)

End Sub

Private Sub CTBCredito_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Credito_Validate(Cancel)

End Sub

Private Sub CTBDebito_Change()

    Call objContabil.Contabil_Debito_Change

End Sub

Private Sub CTBDebito_GotFocus()

    Call objContabil.Contabil_Debito_GotFocus

End Sub

Private Sub CTBDebito_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Debito_KeyPress(KeyAscii)

End Sub

Private Sub CTBDebito_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Debito_Validate(Cancel)

End Sub

Private Sub CTBSeqContraPartida_Change()

    Call objContabil.Contabil_SeqContraPartida_Change

End Sub

Private Sub CTBSeqContraPartida_GotFocus()

    Call objContabil.Contabil_SeqContraPartida_GotFocus

End Sub

Private Sub CTBSeqContraPartida_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_SeqContraPartida_KeyPress(KeyAscii)

End Sub

Private Sub CTBSeqContraPartida_Validate(Cancel As Boolean)

    Call objContabil.Contabil_SeqContraPartida_Validate(Cancel)

End Sub

Private Sub CTBTvwContas_Expand(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwContas_Expand(Node, CTBTvwContas.Nodes)

End Sub

Private Sub CTBHistorico_Change()

    Call objContabil.Contabil_Historico_Change

End Sub

Private Sub CTBHistorico_GotFocus()

    Call objContabil.Contabil_Historico_GotFocus

End Sub

Private Sub CTBHistorico_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Historico_KeyPress(KeyAscii)

End Sub

Private Sub CTBHistorico_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Historico_Validate(Cancel)

End Sub

Private Sub CTBLancAutomatico_Click()

    Call objContabil.Contabil_LancAutomatico_Click

End Sub

Private Sub CTBAglutina_Click()
    
    Call objContabil.Contabil_Aglutina_Click

End Sub

Private Sub CTBAglutina_GotFocus()

    Call objContabil.Contabil_Aglutina_GotFocus

End Sub

Private Sub CTBAglutina_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Aglutina_KeyPress(KeyAscii)

End Sub

Private Sub CTBAglutina_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Aglutina_Validate(Cancel)

End Sub

Private Sub CTBTvwContas_NodeClick(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwContas_NodeClick(Node)

End Sub

Private Sub CTBTvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwCcls_NodeClick(Node)

End Sub

Private Sub CTBListHistoricos_DblClick()

    Call objContabil.Contabil_ListHistoricos_DblClick

End Sub

Private Sub CTBBotaoLimparGrid_Click()

    Call objContabil.Contabil_Limpa_GridContabil

End Sub

Private Sub CTBLote_Change()

    Call objContabil.Contabil_Lote_Change

End Sub

Private Sub CTBLote_GotFocus()

    Call objContabil.Contabil_Lote_GotFocus

End Sub

Private Sub CTBLote_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Lote_Validate(Cancel, Parent)

End Sub

Private Sub CTBDataContabil_Change()

    Call objContabil.Contabil_DataContabil_Change

End Sub

Private Sub CTBDataContabil_GotFocus()

    Call objContabil.Contabil_DataContabil_GotFocus

End Sub

Private Sub CTBDataContabil_Validate(Cancel As Boolean)

    Call objContabil.Contabil_DataContabil_Validate(Cancel, Parent)

End Sub

Private Sub objEventoLote_evSelecao(obj1 As Object)
'traz o lote selecionado para a tela

    Call objContabil.Contabil_objEventoLote_evSelecao(obj1)

End Sub

Private Sub objEventoDoc_evSelecao(obj1 As Object)

    Call objContabil.Contabil_objEventoDoc_evSelecao(obj1)

End Sub

Private Sub CTBDocumento_Change()

    Call objContabil.Contabil_Documento_Change

End Sub

Private Sub CTBDocumento_GotFocus()

    Call objContabil.Contabil_Documento_GotFocus

End Sub

Private Sub CTBBotaoImprimir_Click()

    Call objContabil.Contabil_BotaoImprimir_Click

End Sub

Private Sub CTBUpDown_DownClick()

    Call objContabil.Contabil_UpDown_DownClick

End Sub

Private Sub CTBUpDown_UpClick()

    Call objContabil.Contabil_UpDown_UpClick

End Sub

Private Sub CTBLabelDoc_Click()

    Call objContabil.Contabil_LabelDoc_Click

End Sub

Private Sub CTBLabelLote_Click()

    Call objContabil.Contabil_LabelLote_Click

End Sub

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim dQuantidadeConvertida As Double
Dim sProduto As String
Dim sUM As String
Dim dQuantidade As Double

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico

        Case QUANT_ESTOQUE
            For iLinha = 1 To objGrid.iLinhasExistentes
            
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_Quantidade_Col)) > 0 Then
                    
                    If Len(GridMovimentos.TextMatrix(iLinha, iGrid_UnidadeMed_Col)) > 0 Then
                    
                        lErro = CF("UMEstoque_Conversao", GridMovimentos.TextMatrix(iLinha, iGrid_Produto_Col), GridMovimentos.TextMatrix(iLinha, iGrid_UnidadeMed_Col), CDbl(GridMovimentos.TextMatrix(iLinha, iGrid_Quantidade_Col)), dQuantidadeConvertida)
                        If lErro <> SUCESSO Then gError 64208

                        objMnemonicoValor.colValor.Add dQuantidadeConvertida
                    Else
                        objMnemonicoValor.colValor.Add 0
                    End If
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next

        Case CODIGO1
            If Len(Codigo.Text) > 0 Then
                objMnemonicoValor.colValor.Add Codigo.Text
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case DATA1
            If Len(Data.ClipText) > 0 Then
                objMnemonicoValor.colValor.Add CDate(Data.FormattedText)
            Else
                objMnemonicoValor.colValor.Add DATA_NULA
            End If

        Case CCL1
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_Ccl_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_Ccl_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next

        Case ALMOXARIFADO1
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_Almoxarifado_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_Almoxarifado_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next

        Case ESTORNO1
            For iLinha = 1 To objGrid.iLinhasExistentes
                objMnemonicoValor.colValor.Add 0 'CInt(GridMovimentos.TextMatrix(iLinha, iGrid_Estorno_Col))
            Next

        Case Produto1
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_Produto_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next

        Case UNIDADE_MED
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_UnidadeMed_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_UnidadeMed_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next

        Case DESCRICAO1
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_Descricao_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_Descricao_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next

        Case OP_CODIGO
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_OP_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_OP_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next

        Case PRODUTO_OP
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_ProdutoOP_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_ProdutoOP_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next

        Case TIPO_MOV
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_TipoMov_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add Codigo_Extrai(GridMovimentos.TextMatrix(iLinha, iGrid_TipoMov_Col))
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next

        Case QUANTIDADE1
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_Quantidade_Col)) > 0 Then
                    
                    'Guarda os valores que serão passados como parâmetros em UMEstoque_Conversao
                    sProduto = GridMovimentos.TextMatrix(iLinha, iGrid_Produto_Col)
                    sUM = GridMovimentos.TextMatrix(iLinha, iGrid_UnidadeMed_Col)
                    dQuantidade = StrParaDbl(GridMovimentos.TextMatrix(iLinha, iGrid_Quantidade_Col))
                    
                    'Converte a quantidade para UM padrão estoque
                    lErro = CF("UMEstoque_Conversao", sProduto, sUM, dQuantidade, dQuantidadeConvertida)
                    If lErro <> SUCESSO Then gError 79943
                    
                    objMnemonicoValor.colValor.Add dQuantidadeConvertida
                
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
        
        Case CONTACONTABILCREDITO1
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_ContaContabilCredito_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_ContaContabilCredito_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
        
        Case CONTACONTABILDEBITO1
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_ContaContabilDebito_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_ContaContabilDebito_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next

        Case Else
            Error 39614

        End Select

        Calcula_Mnemonico = SUCESSO

        Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 39614
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
        
        Case 64208
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162854)

    End Select

    Exit Function

End Function

Private Function ContaContabilEst_Le(sContaEnxuta As String) As Long
'Conta contabil ---> vem como PADRAO da  tabela EstoqueProduto
'Caso nao encontre -----> não tratar erro

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_ContaContabilEst_Le

    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col))) > 0 And Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))) > 0 Then
        
        'preenche o objEstoqueProduto
        objAlmoxarifado.sNomeReduzido = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col)
        
        'pega o nome reduzido do Almoxarifado
        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25060 Then gError 49682
        
        If lErro = 25060 Then gError 52006
    
        'Formata o Produto para BD
        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 49683
    
        objEstoqueProduto.sProduto = sProdutoFormatado
        objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
        
        'le a conta de Contabil em Estoque Produto e senão encontrou le em Almoxarifado
        lErro = CF("EstoqueProdutoCC_Le", objEstoqueProduto)
        If lErro <> SUCESSO And lErro <> 49991 Then gError 49684
        
        If lErro = SUCESSO Then
            
            'devolve a conta no Formato de Tela
            lErro = Mascara_RetornaContaEnxuta(objEstoqueProduto.sContaContabil, sContaEnxuta)
            If lErro <> SUCESSO Then gError 49685
        
        End If
        
    End If
        
    ContaContabilEst_Le = SUCESSO
    
    Exit Function
    
Erro_ContaContabilEst_Le:

    ContaContabilEst_Le = gErr
    
        Select Case gErr
            
            Case 49682, 49683, 49684
            
            Case 49685
                Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objEstoqueProduto.sContaContabil)
            
            Case 52006
                Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)
           
            Case Else
                Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162855)
        
        End Select
        
        Exit Function
        
End Function

Function ContaContabilAplic_Le(sContaEnxuta As String) As String

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objTipoDeProduto As New ClassTipoDeProduto

On Error GoTo Erro_ContaContabilAplic_Le
    
    'Formata o Produto para o BD
    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 30384
    
    objProduto.sCodigo = sProdutoFormatado
    
    'Lê os demais atributos do Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 52083

    If lErro = 28030 Then gError 52084
    
    If giTipoVersao = VERSAO_FULL Then
        
        'verifica se a conta esta Preenchida, se nao le a conta em Estoque Produto
        If Trim(objProduto.sContaContabil) = "" Then
            
            objTipoDeProduto.iTipo = objProduto.iTipo
            
            'le a conta com o tipo passado como parametro
            lErro = CF("TipoDeProduto_Le", objTipoDeProduto)
            If lErro <> SUCESSO And lErro <> 22531 Then gError 52085
            
            'se não encontrou ----> Erro
            If lErro = 22531 Then gError 52086
            
            'preenche a conta no objProduto
            objProduto.sContaContabil = objTipoDeProduto.sContaContabil
                        
        End If
    
    End If
    
    If Trim(Len(objProduto.sContaContabil)) > 0 Then
    
        'devolve a conta no formato de Tela
        lErro = Mascara_RetornaContaEnxuta(objProduto.sContaContabil, sContaEnxuta)
        If lErro <> SUCESSO Then gError 52087
    
    End If
    
    ContaContabilAplic_Le = SUCESSO
    
    Exit Function

Erro_ContaContabilAplic_Le:

    ContaContabilAplic_Le = gErr

    Select Case gErr
            
            Case 52083, 52085
            
            Case 52084
                Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
            Case 52086
                Call Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_NAO_CADASTRADO", gErr, objTipoDeProduto.iTipo)
                        
            Case 52087
                Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objProduto.sContaContabil)
            
            Case Else
                Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162856)
        
        End Select
        
    Exit Function
    
End Function

Function ContaContabilProd_Le(sContaEnxuta As String) As String

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objTipoDeProduto As New ClassTipoDeProduto

On Error GoTo Erro_ContaContabilProd_Le
    
    'Formata o Produto para o BD
    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 55366
    
    objProduto.sCodigo = sProdutoFormatado
    
    'Lê os demais atributos do Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 55367

    If lErro = 28030 Then gError 55368
    
    If giTipoVersao = VERSAO_FULL Then
        
        'verifica se a conta esta Preenchida, se nao le a conta em Estoque Produto
        If Trim(objProduto.sContaContabilProducao) = "" Then
            
            objTipoDeProduto.iTipo = objProduto.iTipo
            
            'le a conta com o tipo passado como parametro
            lErro = CF("TipoDeProduto_Le", objTipoDeProduto)
            If lErro <> SUCESSO And lErro <> 22531 Then gError 55369
            
            'se não encontrou ----> Erro
            If lErro = 22531 Then gError 55370
            
            'preenche a conta no objProduto
            objProduto.sContaContabilProducao = objTipoDeProduto.sContaProducao
                        
        End If
    
    End If
    
    If Trim(Len(objProduto.sContaContabilProducao)) > 0 Then
    
        'devolve a conta no formato de Tela
        lErro = Mascara_RetornaContaEnxuta(objProduto.sContaContabilProducao, sContaEnxuta)
        If lErro <> SUCESSO Then gError 55371
    
    End If
    
    ContaContabilProd_Le = SUCESSO
    
    Exit Function

Erro_ContaContabilProd_Le:

    ContaContabilProd_Le = gErr

    Select Case gErr
            
            Case 55366, 55367, 55369
            
            Case 55368
                Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
            Case 55370
                Call Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_NAO_CADASTRADO", gErr, objTipoDeProduto.iTipo)
                        
            Case 55371
                Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objProduto.sContaContabil)
            
            Case Else
                Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162857)
        
        End Select
        
    Exit Function
    
End Function

Private Function QuantDisponivel_Calcula(sProduto As String, sAlmoxarifado As String, Optional objProduto As ClassProduto) As Long

Dim lErro As Long

On Error GoTo Erro_QuantDisponivel_Calcula

    If (objProduto Is Nothing) Then

        lErro = QuantDisponivel_Calcula1(sProduto, sAlmoxarifado)
        If lErro <> SUCESSO Then gError 55349
        
    Else
    
        lErro = QuantDisponivel_Calcula1(sProduto, sAlmoxarifado, objProduto)
        If lErro <> SUCESSO Then gError 55350

    End If

    lErro = Testa_Quantidade()
    If lErro <> SUCESSO Then gError 55351

    QuantDisponivel_Calcula = SUCESSO

    Exit Function

Erro_QuantDisponivel_Calcula:

    QuantDisponivel_Calcula = gErr

    Select Case gErr

        Case 55349, 55350, 55351

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162858)

    End Select

    Exit Function

End Function

Private Function QuantDisponivel_Lote_Calcula(sProduto As String, sAlmoxarifado As String, Optional objProduto As ClassProduto) As Long

Dim lErro As Long

On Error GoTo Erro_QuantDisponivel_Lote_Calcula

    If (objProduto Is Nothing) Then

        lErro = QuantDisponivel_Lote_Calcula1(sProduto, sAlmoxarifado, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col)))
        If lErro <> SUCESSO Then gError 78671
        
    Else
    
        lErro = QuantDisponivel_Lote_Calcula1(sProduto, sAlmoxarifado, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col)), objProduto)
        If lErro <> SUCESSO Then gError 78672

    End If

    lErro = Testa_Quantidade()
    If lErro <> SUCESSO Then gError 78673

    QuantDisponivel_Lote_Calcula = SUCESSO

    Exit Function

Erro_QuantDisponivel_Lote_Calcula:

    QuantDisponivel_Lote_Calcula = gErr

    Select Case gErr

        Case 78671, 78672, 78673 'Erros já tratados

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162859)

    End Select

    Exit Function

End Function

Private Function QuantDisponivel_Calcula1(sProduto As String, sAlmoxarifado As String, Optional objProduto As ClassProduto) As Long
'descobre a quantidade disponivel e coloca na tela

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sUnidadeMed As String
Dim dFator As Double
Dim dQuantTotal As Double
Dim dQuantidade As Double
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim iTipoMov As Integer

On Error GoTo Erro_QuantDisponivel_Calcula1

    QuantDisponivel.Caption = ""

    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col))) > 0 Then

        'Verifica se o produto está preenchido
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 55352
    
        If GridMovimentos.Row >= GridMovimentos.FixedRows And Len(Trim(sAlmoxarifado)) <> 0 And iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
            If (objProduto Is Nothing) Then
                Set objProduto = New ClassProduto
    
                objProduto.sCodigo = sProdutoFormatado
    
                'Lê o produto no BD para obter UM de estoque
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 55353
    
                If lErro = 28030 Then gError 55354
    
            End If
    
            objAlmoxarifado.sNomeReduzido = sAlmoxarifado
    
            lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then gError 55355
    
            If lErro = 25056 Then gError 55356
    
            objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
            objEstoqueProduto.sProduto = sProdutoFormatado
    
            'Lê o Estoque Produto correspondente ao Produto e ao Almoxarifado
            lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
            If lErro <> SUCESSO And lErro <> 21306 Then gError 55357
    
            'Se não encontrou EstoqueProduto no Banco de Dados
            If lErro = 21306 Then
            
                 QuantDisponivel.Caption = Formata_Estoque(0)
    
            Else
                sUnidadeMed = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col)
        
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProduto.sSiglaUMEstoque, sUnidadeMed, dFator)
                If lErro <> SUCESSO Then gError 55358
                
                iTipoMov = Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoMov_Col))
        
                If iTipoMov = MOV_EST_INT2_ENT_CONSERTO_TERC Or iTipoMov = MOV_EST_INT2_DEV_CONSERTO_TERC Then
                    QuantDisponivel.Caption = Formata_Estoque(objEstoqueProduto.dQuantConserto3 * dFator)
                ElseIf iTipoMov = MOV_EST_INT2_SAIDA_CONSERTO Then
                    QuantDisponivel.Caption = Formata_Estoque(objEstoqueProduto.dQuantDisponivel * dFator)
                ElseIf iTipoMov = MOV_EST_INT2_DEV_CONSERTO Then
                    QuantDisponivel.Caption = Formata_Estoque(objEstoqueProduto.dQuantConserto * dFator)
                End If
            End If
    
        Else
    
            'Limpa a Quantidade Disponível da Tela
            QuantDisponivel.Caption = ""
    
        End If

    End If
    
    QuantDisponivel_Calcula1 = SUCESSO

    Exit Function

Erro_QuantDisponivel_Calcula1:

    QuantDisponivel_Calcula1 = gErr

    Select Case gErr

        Case 55352, 55353, 55355, 55357, 55358

        Case 55354
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 55356
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162860)

    End Select

    Exit Function

End Function

Private Function QuantDisponivel_Lote_Calcula1(sProduto As String, sAlmoxarifado As String, sLote As String, iFilialOP As Integer, Optional objProduto As ClassProduto) As Long
'descobre a quantidade disponivel do Lote passado e coloca na tela

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sUnidadeMed As String
Dim dFator As Double
Dim dQuantTotal As Double
Dim dQuantidade As Double
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objRastreamentoLoteSaldo As New ClassRastreamentoLoteSaldo
Dim iTipoMov As Integer

On Error GoTo Erro_QuantDisponivel_Lote_Calcula1

    QuantDisponivel.Caption = ""

    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col))) > 0 Then

        'Verifica se o produto está preenchido
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 78674
        
        If GridMovimentos.Row >= GridMovimentos.FixedRows And Len(Trim(sAlmoxarifado)) <> 0 And iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
            If (objProduto Is Nothing) Then
                Set objProduto = New ClassProduto
    
                objProduto.sCodigo = sProdutoFormatado
    
                'Lê o produto no BD para obter UM de estoque
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 78675
    
                If lErro = 28030 Then gError 78676
    
            End If
    
            objAlmoxarifado.sNomeReduzido = sAlmoxarifado
    
            lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then gError 78677
    
            If lErro = 25056 Then gError 78678
    
            objRastreamentoLoteSaldo.iAlmoxarifado = objAlmoxarifado.iCodigo
            objRastreamentoLoteSaldo.sProduto = sProdutoFormatado
            objRastreamentoLoteSaldo.sLote = sLote
            objRastreamentoLoteSaldo.iFilialOP = iFilialOP
                       
            'Lê o Estoque Produto correspondente ao Produto e ao Almoxarifado
            lErro = CF("RastreamentoLoteSaldo_Le", objRastreamentoLoteSaldo)
            If lErro <> SUCESSO And lErro <> 78633 Then gError 78679
    
            'Se não encontrou EstoqueProduto no Banco de Dados
            If lErro = 78633 Then
            
                 QuantDisponivel.Caption = Formata_Estoque(0)
    
            Else
                sUnidadeMed = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col)
        
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProduto.sSiglaUMEstoque, sUnidadeMed, dFator)
                If lErro <> SUCESSO Then gError 78680
                
                iTipoMov = Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoMov_Col))
        
                If iTipoMov = MOV_EST_INT2_ENT_CONSERTO_TERC Or iTipoMov = MOV_EST_INT2_DEV_CONSERTO_TERC Then
                    QuantDisponivel.Caption = Formata_Estoque(objRastreamentoLoteSaldo.dQuantConserto3 * dFator)
                ElseIf iTipoMov = MOV_EST_INT2_SAIDA_CONSERTO Then
                    QuantDisponivel.Caption = Formata_Estoque(objRastreamentoLoteSaldo.dQuantDispNossa * dFator)
                ElseIf iTipoMov = MOV_EST_INT2_DEV_CONSERTO Then
                    QuantDisponivel.Caption = Formata_Estoque(objRastreamentoLoteSaldo.dQuantConserto * dFator)
                End If
            End If
    
        Else
    
            'Limpa a Quantidade Disponível da Tela
            QuantDisponivel.Caption = ""
    
        End If

    End If
    
    QuantDisponivel_Lote_Calcula1 = SUCESSO

    Exit Function

Erro_QuantDisponivel_Lote_Calcula1:

    QuantDisponivel_Lote_Calcula1 = gErr

    Select Case gErr

        Case 78674, 78675, 78677, 78679, 78680

        Case 78676
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 78678
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162861)

    End Select

    Exit Function

End Function

Private Function Testa_Quantidade() As Long

Dim dQuantidade As Double
Dim lErro As Long
Dim iEntradaSaida As Integer

On Error GoTo Erro_Testa_Quantidade

    If GridMovimentos.Row >= GridMovimentos.FixedRows Then

        If colItensNumIntDoc.Item(GridMovimentos.Row) = 0 Then
    
            'Testa se é movimento de entrada ou saída
            Call Testa_TipoMovimento(GridMovimentos.Row, iEntradaSaida)
    
            'Se a quantidade está preenchida e é uma saida de material do estoque
            If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))) <> 0 And iEntradaSaida = ESTOQUE_SAIDA Then
    
                dQuantidade = CDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))
    
                'Calcula a Quantidade requisitada
                lErro = Testa_QuantRequisitada(dQuantidade)
                If lErro <> SUCESSO Then gError 55359
    
            End If
    
        End If

    End If

    Testa_Quantidade = SUCESSO

    Exit Function

Erro_Testa_Quantidade:

    Testa_Quantidade = gErr

    Select Case gErr

        Case 55359

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162862)

    End Select
    
    Exit Function
    
End Function

Private Function Testa_QuantRequisitada(ByVal dQuantAtual As Double) As Long
'verifica se o saldo que está saindo do estoque é menor ou igual a quantidade disponivel

Dim lErro As Long
Dim iIndice As Integer
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sProdutoAtual As String
Dim sAlmoxarifado As String
Dim sAlmoxarifadoAtual As String
Dim sUnidadeAtual As String
Dim sUnidadeProd As String
Dim dQuantidadeProd As String
Dim dFator As Double
Dim objProduto As New ClassProduto
Dim dQuantTotal As Double
Dim iEntradaSaida As Integer

On Error GoTo Erro_Testa_QuantRequisitada

    If gobjMAT.iAceitaEstoqueNegativo = DESMARCADO Then

        sProdutoAtual = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)
        sAlmoxarifadoAtual = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col)
        sUnidadeAtual = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col)
    
        If Len(sProdutoAtual) > 0 And Len(sAlmoxarifadoAtual) > 0 And Len(sUnidadeAtual) > 0 Then
    
            lErro = CF("Produto_Formata", sProdutoAtual, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 55360
    
            objProduto.sCodigo = sProdutoFormatado
    
            'Lê o produto para saber qual é a sua ClasseUM
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 55361
        
            If lErro = 28030 Then gError 55362
        
            For iIndice = 1 To objGrid.iLinhasExistentes
        
                'Não pode somar a Linha atual
                If GridMovimentos.Row <> iIndice Then
        
                    sCodProduto = GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col)
                    sAlmoxarifado = GridMovimentos.TextMatrix(iIndice, iGrid_Almoxarifado_Col)
        
                    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
                    If lErro <> SUCESSO Then gError 55363
        
                    'Verifica se há outras movimentações do Produto no mesmo Almoxarifado
                    If UCase(sAlmoxarifado) = UCase(sAlmoxarifadoAtual) And UCase(objProduto.sCodigo) = UCase(sProdutoFormatado) Then
        
                        'Verifica se há alguma Quantidade informada
                        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))) <> 0 Then
        
                            sUnidadeProd = GridMovimentos.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
        
                            dQuantidadeProd = CDbl(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))
        
                            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUnidadeProd, sUnidadeAtual, dFator)
                            If lErro <> SUCESSO Then gError 55364
        
                            'Testa se é movimento de entrada ou saída
                            Call Testa_TipoMovimento(iIndice, iEntradaSaida)
        
                            If iEntradaSaida = ESTOQUE_ENTRADA Then
                                dQuantTotal = dQuantTotal - (dQuantidadeProd * dFator)
                            Else
                                dQuantTotal = dQuantTotal + (dQuantidadeProd * dFator)
                            End If
        
                        End If
        
                    End If
        
                End If
        
            Next
        
            dQuantTotal = dQuantTotal + dQuantAtual
    
            If dQuantTotal > StrParaDbl(QuantDisponivel.Caption) Then gError 55365
    
        End If

    End If

    Testa_QuantRequisitada = SUCESSO

    Exit Function

Erro_Testa_QuantRequisitada:

    Testa_QuantRequisitada = gErr

    Select Case gErr

        Case 55360, 55361, 55363, 55364

        Case 55362
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, sCodProduto)

        Case 55365
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_REQ_MAIOR", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162863)

    End Select

    Exit Function

End Function

Private Sub MovEstoque_Trata_Estorno_Versao_Light()

Dim iEstorno As Integer
Dim iLinha As Integer

'    If giTipoVersao = VERSAO_LIGHT Then
'
'        iEstorno = 0
'
'        For iLinha = 1 To objGrid.iLinhasExistentes
'
'            If GridMovimentos.TextMatrix(iLinha, iGrid_Estorno_Col) = MARCADO Then
'
'                iEstorno = 1
'                Exit For
'
'            End If
'
'        Next
'
'        Call objContabil.Contabil_Trata_Estorno_Versao_Light(iEstorno)
'
'    End If

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_MOVIMENTOS_ESTOQUE_MOVIMENTO
    Set Form_Load_Ocx = Me
    Caption = "Movimentos de Estoque"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "MovEstoqueSRV"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is AlmoxPadrao Then
            Call AlmoxPadraoLabel_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
        ElseIf Me.ActiveControl Is CclPadrao Then
            Call CclPadraoLabel_Click
        ElseIf Me.ActiveControl Is Codigo Then
            Call CodigoLabel_Click
        ElseIf Me.ActiveControl Is ProdutoOP Then
            Call BotaoServicos_Click
        ElseIf Me.ActiveControl Is Almoxarifado Then
            Call BotaoEstoque_Click
        ElseIf Me.ActiveControl Is OPCodigo Then
            Call BotaoOP_Click
        ElseIf Me.ActiveControl Is Ccl Then
            Call BotaoCcls_Click
        ElseIf Me.ActiveControl Is ContaContabilCredito Or Me.ActiveControl Is ContaContabilDebito Then
            Call BotaoPlanoConta_Click
        ElseIf Me.ActiveControl Is Lote Then
            Call BotaoLote_Click
        ElseIf Me.ActiveControl Is OP Then
            Call OPLabel_Click
        ElseIf Me.ActiveControl Is ProdutoOPGera Then
            Call ProdutoOPLabel_Click
        End If
    End If

End Sub


Private Sub objEventoProdutoOP_OP_evSelecao(obj1 As Object)

Dim objItemOP As ClassOS
Dim lErro As Long
Dim sProdutoMascarado As String

On Error GoTo Erro_objEventoProdutoOP_OP_evSelecao

    Set objItemOP = obj1

    sProdutoMascarado = String(STRING_PRODUTO, 0)

    lErro = Mascara_RetornaProdutoTela(objItemOP.sServico, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 206231

    ProdutoOPGera.PromptInclude = False
    ProdutoOPGera.Text = sProdutoMascarado
    ProdutoOPGera.PromptInclude = True
    
    Call ProdutoOPGera_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

Erro_objEventoProdutoOP_OP_evSelecao:

    Select Case gErr
        
        Case 206231
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOTELA", gErr, objItemOP.sServico)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206232)

    End Select
    
    Exit Sub

End Sub


Public Sub OPLabel_Click()

Dim objOrdemProducao As New ClassOS
Dim colSelecao As New Collection
Dim objItemOP As New ClassOS
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_OPLabel_Click
    
    lErro = CF("Produto_Formata", ProdutoOPGera.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Se o Produto estiver preenchido
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        colSelecao.Add sProdutoFormatado
        
        If Len(Trim(OP.Text)) <> 0 Then objItemOP.sCodigo = Trim(OP.Text)
                        
        Call Chama_Tela("ItemOSLista", colSelecao, objItemOP, objEventoOP_OP)
        
    Else
    
        If Len(Trim(OP.Text)) <> 0 Then objOrdemProducao.sCodigo = Trim(OP.Text)

        Call Chama_Tela("OSAbertaLista", colSelecao, objOrdemProducao, objEventoOP_OP_Todos)
    
    End If
    
    Exit Sub
    
Erro_OPLabel_Click:
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206235)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoOP_OP_evSelecao(obj1 As Object)

Dim objItemOP As ClassItemOS

    Set objItemOP = obj1

    'Preeche OP com o codigo de objOrdemProducao
    OP.Text = objItemOP.sCodigo

    Me.Show

End Sub

Private Sub objEventoOP_OP_Todos_evSelecao(obj1 As Object)

Dim objOP As ClassOS

    Set objOP = obj1

    'Preeche OP com o codigo de objOrdemProducao
    OP.Text = objOP.sCodigo

    Me.Show

End Sub

Public Sub ProdutoOPLabel_Click()

Dim objProduto As New ClassProduto
Dim colSelecao As New Collection
Dim objItemOP As New ClassOS
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_ProdutoOPLabel_Click
    
    If Len(Trim(OP.Text)) = 0 Then gError 206237
    
    lErro = CF("Produto_Formata", ProdutoOPGera.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then objItemOP.sServico = sProdutoFormatado
        
    colSelecao.Add Trim(OP.Text)
        
    Call Chama_Tela("ItemOSLista", colSelecao, objItemOP, objEventoProdutoOP_OP, "Codigo = ?")
    
    Exit Sub
    
Erro_ProdutoOPLabel_Click:
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 206237
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGOOS_NAO_PREENCHIDO", gErr)
            OP.SetFocus
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206238)

    End Select
    
    Exit Sub

End Sub

Private Sub objEventoProdutoOP_evSelecao(obj1 As Object)

Dim objProduto As ClassProduto
Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String

On Error GoTo Erro_objEventoProdutoOP_evSelecao

    Set objProduto = obj1

    lErro = CF("Produto_Formata", objProduto.sCodigo, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 206239

        ProdutoOPGera.PromptInclude = False
        ProdutoOPGera.Text = sProdutoMascarado
        ProdutoOPGera.PromptInclude = True
            
    End If

    Me.Show

    Exit Sub

Erro_objEventoProdutoOP_evSelecao:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 206239
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOTELA", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206240)

    End Select
    
    Exit Sub

End Sub

Private Sub CodigoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CodigoLabel, Source, X, Y)
End Sub

Private Sub CodigoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CodigoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
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

Private Sub CclPadraoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CclPadraoLabel, Source, X, Y)
End Sub

Private Sub CclPadraoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CclPadraoLabel, Button, Shift, X, Y)
End Sub

Private Sub AlmoxPadraoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(AlmoxPadraoLabel, Source, X, Y)
End Sub

Private Sub AlmoxPadraoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(AlmoxPadraoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub QuantDisponivel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantDisponivel, Source, X, Y)
End Sub

Private Sub QuantDisponivel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantDisponivel, Button, Shift, X, Y)
End Sub

Private Sub CTBCclLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclLabel, Source, X, Y)
End Sub

Private Sub CTBCclLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclLabel, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel7, Source, X, Y)
End Sub

Private Sub CTBLabel7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel7, Button, Shift, X, Y)
End Sub

Private Sub CTBContaDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBContaDescricao, Source, X, Y)
End Sub

Private Sub CTBContaDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBContaDescricao, Button, Shift, X, Y)
End Sub

Private Sub CTBCclDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclDescricao, Source, X, Y)
End Sub

Private Sub CTBCclDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclDescricao, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel21, Source, X, Y)
End Sub

Private Sub CTBLabel21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel21, Button, Shift, X, Y)
End Sub

Private Sub CTBOrigem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBOrigem, Source, X, Y)
End Sub

Private Sub CTBOrigem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBOrigem, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel14, Source, X, Y)
End Sub

Private Sub CTBLabel14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel14, Button, Shift, X, Y)
End Sub

Private Sub CTBPeriodo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBPeriodo, Source, X, Y)
End Sub

Private Sub CTBPeriodo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBPeriodo, Button, Shift, X, Y)
End Sub

Private Sub CTBExercicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBExercicio, Source, X, Y)
End Sub

Private Sub CTBExercicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBExercicio, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel13, Source, X, Y)
End Sub

Private Sub CTBLabel13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel13, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel5, Source, X, Y)
End Sub

Private Sub CTBLabel5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel5, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelHistoricos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelHistoricos, Source, X, Y)
End Sub

Private Sub CTBLabelHistoricos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelHistoricos, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelContas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelContas, Source, X, Y)
End Sub

Private Sub CTBLabelContas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelContas, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelCcl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelCcl, Source, X, Y)
End Sub

Private Sub CTBLabelCcl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelCcl, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel1, Source, X, Y)
End Sub

Private Sub CTBLabel1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel1, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelTotais, Source, X, Y)
End Sub

Private Sub CTBLabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelTotais, Button, Shift, X, Y)
End Sub

Private Sub CTBTotalDebito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBTotalDebito, Source, X, Y)
End Sub

Private Sub CTBTotalDebito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBTotalDebito, Button, Shift, X, Y)
End Sub

Private Sub CTBTotalCredito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBTotalCredito, Source, X, Y)
End Sub

Private Sub CTBTotalCredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBTotalCredito, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel8, Source, X, Y)
End Sub

Private Sub CTBLabel8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel8, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelDoc_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelDoc, Source, X, Y)
End Sub

Private Sub CTBLabelDoc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelDoc, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelLote_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelLote, Source, X, Y)
End Sub

Private Sub CTBLabelLote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelLote, Button, Shift, X, Y)
End Sub


Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

Private Sub objEventoOP_evSelecao(obj1 As Object)

Dim objOrdemProducao As ClassOS
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long
    
On Error GoTo Erro_objEventoOP_evSelecao

    Set objOrdemProducao = obj1
    
    If (GridMovimentos.Row > 0) Then

        sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)

        lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 52201

        'Verifica se o produto está preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
            'Preenche a OP
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_OP_Col) = objOrdemProducao.sCodigo
            OPCodigo.Text = objOrdemProducao.sCodigo
        
        End If
        
    End If

    Me.Show

    Exit Sub
    
Erro_objEventoOP_evSelecao:

    Select Case gErr
    
        Case 52201
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162867)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoLote_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer
Dim objRastroLoteSaldo As New ClassRastroLoteSaldo
Dim sLote As String
Dim objAlmoxarifado As ClassAlmoxarifado

On Error GoTo Erro_BotaoLote_Click

    If (GridMovimentos.Row = 0) Then gError 140223

    sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)
    sLote = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 140224

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 140225
    
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col))) = 0 Then gError 177296
    
    Set objAlmoxarifado = New ClassAlmoxarifado
    
    objAlmoxarifado.sNomeReduzido = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col)

    lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
    If lErro <> SUCESSO And lErro <> 25056 Then gError 177297
       
    If Len(Trim(sLote)) > 0 Then
        objRastroLoteSaldo.sLote = sLote
    End If

    colSelecao.Add sProdutoFormatado
    colSelecao.Add objAlmoxarifado.iCodigo

    Call Chama_Tela("RastroLoteSaldoLista", colSelecao, objRastroLoteSaldo, objEventoRastroLote, "Produto = ? AND Almoxarifado = ?")

    Exit Sub

Erro_BotaoLote_Click:

    Select Case gErr

        Case 140224, 177297
        
        Case 140223
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
                    
        Case 140225
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
            
        Case 177296
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_PREENCHIDO", gErr, GridMovimentos.Row)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165459)

    End Select

    Exit Sub

End Sub

Private Sub objEventoRastroLote_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRastroLoteSaldo As New ClassRastroLoteSaldo
Dim objProduto As New ClassProduto

On Error GoTo Erro_objEventoRastroLote_evSelecao

    Set objRastroLoteSaldo = obj1

    If (GridMovimentos.Row > 0) Then
        Lote.Text = objRastroLoteSaldo.sLote
        
        'Carrega as séries na coleção global
        lErro = Carrega_Series(gcolcolRastreamentoSerie.Item(GridMovimentos.Row), StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)), Lote.Text, StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), GridMovimentos.Row)
        If lErro <> SUCESSO Then gError 141913
        
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col) = objRastroLoteSaldo.sLote
    End If

    objProduto.sCodigo = objRastroLoteSaldo.sProduto
            
    'Lê os demais atributos do Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 140226
    
    If objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
            'Preenche a Quantidade
            lErro = QuantDisponivel_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col))
            If lErro <> SUCESSO Then gError 140228
    Else
        'Preenche a Quantidade do Lote
        lErro = QuantDisponivel_Lote_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col), Lote.Text, Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialOP_Col)))
        If lErro <> SUCESSO Then gError 140228
    End If
    
    Me.Show

    Exit Sub

Erro_objEventoRastroLote_evSelecao:

    Select Case gErr
    
        Case 140226 To 140228, 141913
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165460)

    End Select

    Exit Sub

End Sub

Public Sub BotaoSerie_Click()
'Chama a tela de Lote de Rastreamento

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objItemMovEstoque As New ClassItemMovEstoque
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim bPodeAlterarQtd As Boolean

On Error GoTo Erro_BotaoSerie_Click
    
    'Verifica se tem alguma linha selecionada no Grid
    If GridMovimentos.Row = 0 Then gError 141914
    
    'Se o produto não foi preenchido, erro
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))) = 0 Then gError 141915
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col))) = 0 Then gError 177303
    If StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)) = 0 Then gError 177304
        
    'Formata o produto
    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 141916
    
    'Lê o produto
    objProduto.sCodigo = sProdutoFormatado
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 141917
       
    objItemMovEstoque.dQuantidade = StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))
    objItemMovEstoque.iItemNF = GridMovimentos.Row
    objItemMovEstoque.sAlmoxarifadoNomeRed = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col)
    objItemMovEstoque.sProduto = sProdutoFormatado
    objItemMovEstoque.sSiglaUM = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col)
    objItemMovEstoque.iTipoMov = Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_TipoMov_Col))
       
    bPodeAlterarQtd = True
    If colItensNumIntDoc.Count >= GridMovimentos.Row Then
        If colItensNumIntDoc.Item(GridMovimentos.Row) <> 0 Then
            bPodeAlterarQtd = False
        End If
    End If
        
    'Chama a tela de browse RastroLoteLista passando como parâmetro a seleção do Filtro (sSelecao)
    Call Chama_Tela_Modal("RastreamentoSerie", gcolcolRastreamentoSerie.Item(GridMovimentos.Row), objItemMovEstoque, Me.Name, bPodeAlterarQtd)
                    
    lErro = Acerta_Quantidade_Rastreada(GridMovimentos.Row)
    If lErro <> SUCESSO Then gError 141918
                    
    Exit Sub

Erro_BotaoSerie_Click:

    Select Case gErr
    
        Case 141914
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 141915 To 141918
        
        Case 177303
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_LOTE_NAO_PREENCHIDO", gErr, GridMovimentos.Row)
        
        Case 177304
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_QUANTLOTE_NAO_PREENCHIDA", gErr, GridMovimentos.Row)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141919)
    
    End Select
    
    Exit Sub

End Sub

Public Function Carrega_Series(colRastreamentoMovto As Collection, ByVal dQuantidade As Double, ByVal sLoteIni As String, ByVal dQuantidadeAnterior As Double, ByVal sLoteIniAnterior As String, ByVal iLinha As Integer)
'Gera as séries a partir da série inicial e quantidade

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iIndice As Integer
Dim objRastreamentoMovto As ClassRastreamentoMovto
Dim objRastreamentoSerie As ClassRastreamentoLote
Dim objRastreamentoSerieIni As ClassRastreamentoLote
Dim objItemMovEstoque As ClassItemMovEstoque
Dim objAlmoxarifado As ClassAlmoxarifado
Dim vbResult As VbMsgBoxResult
Dim colRastreamentoMovtoAux As New Collection
Dim iTipoMovto As Integer

On Error GoTo Erro_Carrega_Series

    'Formata o produto
    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 141921
    
    'Lê o produto
    objProduto.sCodigo = sProdutoFormatado

    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 141922
    
    'Produto não cadastrado
    If lErro = 28030 Then gError 141923

    If objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
    
        iTipoMovto = Codigo_Extrai(GridMovimentos.TextMatrix(iLinha, iGrid_TipoMov_Col))
    
        If dQuantidadeAnterior <> 0 And Len(Trim(sLoteIniAnterior)) <> 0 And iTipoMovtoAnt = iTipoMovto Then
            
            If Abs(dQuantidade - dQuantidadeAnterior) > QTDE_ESTOQUE_DELTA Or sLoteIni <> sLoteIniAnterior Then
            
                vbResult = Rotina_Aviso(vbYesNo, "AVISO_MODIFICACAO_SERIES")
                If vbResult = vbNo Then gError 141920
            Else
                vbResult = vbNo
        
            End If
            
        Else
            vbResult = vbYes
            iTipoMovtoAnt = iTipoMovto
        End If
                    
        If vbResult = vbYes Then
                           
            If Len(Trim(sLoteIni)) <> 0 Then
            
                If Not IsNumeric(right(sLoteIni, objProduto.iSerieParteNum)) Then gError 141924
                
                Set objRastreamentoSerieIni = New ClassRastreamentoLote
                
                objRastreamentoSerieIni.sProduto = objProduto.sCodigo
                objRastreamentoSerieIni.iFilialOP = Codigo_Extrai(GridMovimentos.TextMatrix(iLinha, iGrid_FilialOP_Col))
                objRastreamentoSerieIni.sCodigo = sLoteIni
                
                lErro = CF("RastreamentoLote_Le", objRastreamentoSerieIni)
                If lErro <> SUCESSO And lErro <> 75710 Then gError 141925
                
                If lErro <> SUCESSO Then gError 141926
            
                Set objItemMovEstoque = New ClassItemMovEstoque
                
                objItemMovEstoque.dQuantidade = Fix(dQuantidade)
                objItemMovEstoque.iItemNF = GridMovimentos.Row
                objItemMovEstoque.sAlmoxarifadoNomeRed = GridMovimentos.TextMatrix(iLinha, iGrid_Almoxarifado_Col)
                objItemMovEstoque.sProduto = sProdutoFormatado
                objItemMovEstoque.sSiglaUM = GridMovimentos.TextMatrix(iLinha, iGrid_UnidadeMed_Col)
                objItemMovEstoque.iTipoMov = iTipoMovto
                
                Set objAlmoxarifado = New ClassAlmoxarifado
                
                objAlmoxarifado.sNomeReduzido = objItemMovEstoque.sAlmoxarifadoNomeRed
        
                lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25056 Then gError 177237
                
                objItemMovEstoque.iAlmoxarifado = objAlmoxarifado.iCodigo
                                              
                lErro = CF("Rastreamento_Serie_Gera", objItemMovEstoque, objProduto, sLoteIni, colRastreamentoMovtoAux)
                If lErro <> SUCESSO Then gError 177240

            End If
            
            'Remove os dados anteriores
            For iIndice = colRastreamentoMovto.Count To 1 Step -1
                colRastreamentoMovto.Remove iIndice
            Next
            
            'Coloca os novos dados
            For Each objRastreamentoMovto In colRastreamentoMovtoAux
                colRastreamentoMovto.Add objRastreamentoMovto
            Next
            
        End If
        
    End If

    Carrega_Series = SUCESSO
    
    Exit Function

Erro_Carrega_Series:

    Carrega_Series = gErr

    Select Case gErr
    
        Case 141921, 141922, 141925, 141927, 141920, 141929, 177240
        
        Case 141923
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case 141924
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIEPROX_PARTENUMERICA_NAO_NUMERICA", gErr, right(sLoteIni, objProduto.iSerieParteNum))
        
        Case 141926
            Call Rotina_Erro(vbOKOnly, "ERRO_RASTREAMENTOLOTE_NAO_CADASTRADO", gErr, objRastreamentoSerieIni.sProduto, objRastreamentoSerieIni.sCodigo, objRastreamentoSerieIni.iFilialOP)
        
        Case 141928
            Call Rotina_Erro(vbOKOnly, "ERRO_RASTREAMENTOLOTE_NAO_CADASTRADO", gErr, objRastreamentoSerie.sProduto, objRastreamentoSerie.sCodigo, objRastreamentoSerie.iFilialOP)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141930)

    End Select

    Exit Function

End Function

Public Function Acerta_Quantidade_Rastreada(ByVal iLinha As Integer)
'Acerta a quantidade do grid com base na quantidadse da coleção global de movimentos de séries

Dim lErro As Long
Dim dQuantidade As Double
Dim objRastreamentoSerie As ClassRastreamentoMovto

On Error GoTo Erro_Acerta_Quantidade_Rastreada

    For Each objRastreamentoSerie In gcolcolRastreamentoSerie.Item(iLinha)
    
        dQuantidade = dQuantidade + objRastreamentoSerie.dQuantidade
    
    Next
                
    If colItensNumIntDoc.Item(iLinha) = 0 Then

        If Len(Trim(QuantDisponivel.Caption)) <> 0 Then 'And GridMovimentos.TextMatrix(iLinha, iGrid_Estorno_Col) <> "1" Then

            lErro = Testa_QuantRequisitada(dQuantidade)
            If lErro <> SUCESSO Then gError 141932

        End If

    End If
        
    GridMovimentos.TextMatrix(iLinha, iGrid_Quantidade_Col) = Formata_Estoque(dQuantidade)

    Acerta_Quantidade_Rastreada = SUCESSO
    
    Exit Function

Erro_Acerta_Quantidade_Rastreada:

    Acerta_Quantidade_Rastreada = gErr

    Select Case gErr
    
        Case 141932

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141931)

    End Select

    Exit Function

End Function

Private Sub CTBGerencial_Click()
    
    Call objContabil.Contabil_Gerencial_Click

End Sub

Private Sub CTBGerencial_GotFocus()

    Call objContabil.Contabil_Gerencial_GotFocus

End Sub

Private Sub CTBGerencial_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Gerencial_KeyPress(KeyAscii)

End Sub

Private Sub CTBGerencial_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Gerencial_Validate(Cancel)

End Sub

Public Sub Anotacao_Extrai(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long

On Error GoTo Erro_Anotacao_Extrai

    objAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_MOVESTOQUE
    If Len(Trim(Codigo.Text)) > 0 Then
        objAnotacao.sID = CStr(giFilialEmpresa) & "," & Codigo.Text
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158093)
     
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158094)
     
    End Select
     
    Exit Sub

End Sub

Private Sub Custo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub
Private Sub Custo_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Custo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Custo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Custo
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_Custo(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iEntradaSaida As Integer

On Error GoTo Erro_Saida_Celula_Custo

    Set objGridInt.objControle = Custo

    If Len(Trim(Custo.Text)) > 0 Then

        lErro = Valor_Positivo_Critica(Custo.Text)
        If lErro <> SUCESSO Then gError 132076

        Custo.Text = Format(Custo.Text, Custo.Format)

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 132077
    
    Saida_Celula_Custo = SUCESSO

    Exit Function

Erro_Saida_Celula_Custo:

    Saida_Celula_Custo = gErr

    Select Case gErr

        Case 132076 To 132077
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)

    End Select

    Exit Function

End Function
