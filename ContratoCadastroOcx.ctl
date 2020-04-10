VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl ContratoCadastroOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   5235
      Index           =   2
      Left            =   210
      TabIndex        =   43
      Top             =   660
      Visible         =   0   'False
      Width           =   9210
      Begin VB.CommandButton BotaoServicos 
         Caption         =   "Serviços/Peças"
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
         Left            =   2550
         TabIndex        =   62
         Top             =   4785
         Width           =   1695
      End
      Begin VB.CommandButton BotaoMedicoes 
         Caption         =   "Medições"
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
         Left            =   5595
         TabIndex        =   59
         Top             =   4785
         Width           =   1695
      End
      Begin VB.CommandButton BotaoFaturamentos 
         Caption         =   "Faturamentos"
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
         Left            =   7410
         TabIndex        =   58
         Top             =   4785
         Width           =   1695
      End
      Begin VB.Frame Frame3 
         Caption         =   "Itens"
         Height          =   4740
         Left            =   30
         TabIndex        =   44
         Top             =   0
         Width           =   9135
         Begin VB.TextBox ItemDescricao 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   5235
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   80
            Top             =   1710
            Width           =   3000
         End
         Begin MSMask.MaskEdBox UltParcCobrada 
            Height          =   270
            Left            =   4065
            TabIndex        =   79
            Top             =   1740
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   3
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QtdeParcelas 
            Height          =   270
            Left            =   6015
            TabIndex        =   78
            Top             =   1695
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   3
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataRefFim 
            Height          =   315
            Left            =   6270
            TabIndex        =   61
            Top             =   945
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            MaxLength       =   8
            Format          =   "dd/mm/yy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataRefIni 
            Height          =   315
            Left            =   5265
            TabIndex        =   60
            Top             =   525
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            MaxLength       =   8
            Format          =   "dd/mm/yy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataProx 
            Height          =   315
            Left            =   1545
            TabIndex        =   51
            Top             =   960
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            MaxLength       =   8
            Format          =   "dd/mm/yy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Item 
            Height          =   315
            Left            =   3270
            TabIndex        =   50
            Top             =   675
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.TextBox Obs 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   3075
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   48
            Top             =   1020
            Width           =   3000
         End
         Begin MSMask.MaskEdBox DataInicio 
            Height          =   315
            Left            =   3435
            TabIndex        =   36
            Top             =   750
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            MaxLength       =   8
            Format          =   "dd/mm/yy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Produto 
            Height          =   315
            Left            =   5280
            TabIndex        =   31
            Top             =   1185
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.ComboBox UM 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2070
            TabIndex        =   33
            Top             =   1380
            Width           =   720
         End
         Begin VB.ComboBox CondPagto 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   465
            TabIndex        =   38
            Top             =   1305
            Width           =   1440
         End
         Begin VB.ComboBox Periodicidade 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   540
            TabIndex        =   37
            Top             =   975
            Width           =   1680
         End
         Begin VB.CheckBox Medicao 
            Caption         =   "Medicao"
            Height          =   315
            Left            =   2865
            TabIndex        =   35
            Top             =   1590
            Width           =   1035
         End
         Begin VB.CheckBox Cobrar 
            Caption         =   "Cobrar"
            Height          =   315
            Left            =   4125
            TabIndex        =   30
            Tag             =   "1"
            Top             =   1575
            Width           =   870
         End
         Begin MSMask.MaskEdBox Valor 
            Height          =   315
            Left            =   3945
            TabIndex        =   34
            Top             =   405
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   315
            Left            =   1920
            TabIndex        =   32
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridItens 
            Height          =   1725
            Left            =   75
            TabIndex        =   28
            Top             =   225
            Width           =   8910
            _ExtentX        =   15716
            _ExtentY        =   3043
            _Version        =   393216
         End
         Begin VB.Label Total 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7605
            TabIndex        =   45
            Top             =   4305
            Width           =   1305
         End
         Begin VB.Label TotalLabel 
            Caption         =   "Total:"
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
            TabIndex        =   46
            Top             =   4350
            Width           =   540
         End
      End
      Begin VB.CommandButton BotaoProduto 
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
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   4785
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Identificação"
      Height          =   5280
      Index           =   1
      Left            =   240
      TabIndex        =   39
      Top             =   630
      Width           =   9165
      Begin VB.Frame Frame7 
         Caption         =   "Geração da Nota Fiscal"
         Height          =   495
         Left            =   15
         TabIndex        =   76
         Top             =   4770
         Width           =   9105
         Begin VB.CheckBox optNFe 
            Caption         =   "NF Eletrônica Federal"
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
            Left            =   3855
            TabIndex        =   22
            Top             =   225
            Width           =   2400
         End
         Begin VB.ComboBox Serie 
            Height          =   315
            Left            =   7200
            TabIndex        =   23
            Top             =   135
            Width           =   765
         End
         Begin VB.CheckBox optRecibo 
            Caption         =   "Gerar Recibo provisório de Serviço"
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
            TabIndex        =   21
            Top             =   225
            Width           =   3465
         End
         Begin VB.Label SerieLabel 
            AutoSize        =   -1  'True
            Caption         =   "Série:"
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
            Left            =   6600
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   77
            Top             =   195
            Width           =   510
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Identificação"
         Height          =   2325
         Left            =   0
         TabIndex        =   67
         Top             =   -15
         Width           =   9105
         Begin VB.TextBox Observacao 
            Height          =   735
            Left            =   1125
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   1545
            Width           =   7920
         End
         Begin VB.ComboBox FilialCliente 
            Height          =   315
            Left            =   7140
            TabIndex        =   4
            Top             =   855
            Width           =   1920
         End
         Begin VB.TextBox Descricao 
            Height          =   315
            Left            =   1125
            MaxLength       =   50
            TabIndex        =   2
            Top             =   510
            Width           =   7935
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2475
            TabIndex        =   1
            Top             =   195
            Value           =   1  'Checked
            Width           =   795
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   315
            Left            =   1125
            TabIndex        =   0
            Top             =   165
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Cliente 
            Height          =   315
            Left            =   1125
            TabIndex        =   3
            Top             =   855
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   "_"
         End
         Begin MSComCtl2.UpDown UpDownIniContrato 
            Height          =   300
            Left            =   2235
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   1200
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataIniContrato 
            Height          =   315
            Left            =   1125
            TabIndex        =   5
            Top             =   1200
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownFimContrato 
            Height          =   300
            Left            =   4950
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1200
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataFimContrato 
            Height          =   315
            Left            =   3840
            TabIndex        =   7
            Top             =   1200
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownRenovContrato 
            Height          =   300
            Left            =   8250
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1200
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataRenovContrato 
            Height          =   315
            Left            =   7140
            TabIndex        =   9
            Top             =   1200
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label CodigoLabel 
            Caption         =   "Contrato:"
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
            Left            =   285
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   75
            Top             =   210
            Width           =   795
         End
         Begin VB.Label ObservacaoLabel 
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
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   15
            TabIndex        =   74
            Top             =   1575
            Width           =   1080
         End
         Begin VB.Label ClienteLabel 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   420
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   73
            Top             =   930
            Width           =   660
         End
         Begin VB.Label FilialClienteLabel 
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
            Left            =   6570
            TabIndex        =   72
            Top             =   930
            Width           =   465
         End
         Begin VB.Label DescricaoLabel 
            Caption         =   "Descrição:"
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
            Left            =   120
            TabIndex        =   71
            Top             =   555
            Width           =   960
         End
         Begin VB.Label DataRenovContratoLabel 
            Caption         =   "Renovação:"
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
            Left            =   6075
            TabIndex        =   70
            Top             =   1245
            Width           =   1080
         End
         Begin VB.Label DataIniContratoLabel 
            Caption         =   "Início:"
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
            Left            =   525
            TabIndex        =   69
            Top             =   1245
            Width           =   555
         End
         Begin VB.Label DataFimContratoLabel 
            Caption         =   "Término:"
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
            Left            =   3060
            TabIndex        =   68
            Top             =   1245
            Width           =   750
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Padrão para Cobrança"
         Height          =   570
         Left            =   0
         TabIndex        =   63
         Top             =   2310
         Width           =   9120
         Begin VB.ComboBox PeriodicidadeP 
            Height          =   315
            Left            =   3840
            TabIndex        =   14
            Top             =   225
            Width           =   1860
         End
         Begin VB.ComboBox CondPagtoP 
            Height          =   315
            Left            =   7200
            TabIndex        =   15
            Top             =   225
            Width           =   1875
         End
         Begin MSComCtl2.UpDown UpDownInicioP 
            Height          =   300
            Left            =   2220
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   210
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataInicioP 
            Height          =   300
            Left            =   1110
            TabIndex        =   12
            Top             =   225
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label Label7 
            Caption         =   "Início:"
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
            Left            =   495
            TabIndex        =   66
            Top             =   270
            Width           =   555
         End
         Begin VB.Label Label1 
            Caption         =   "Periodicidade:"
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
            Left            =   2565
            TabIndex        =   65
            Top             =   285
            Width           =   1230
         End
         Begin VB.Label CondPagtoLabel 
            Caption         =   "Condição Pagto:"
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
            Left            =   5760
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   64
            Top             =   285
            Width           =   1770
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Complemento para Tributação\Faturamento"
         Height          =   930
         Left            =   0
         TabIndex        =   53
         Top             =   3840
         Width           =   9120
         Begin MSMask.MaskEdBox NaturezaOp 
            Height          =   315
            Left            =   1875
            TabIndex        =   19
            Top             =   225
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TipoTributacao 
            Height          =   315
            Left            =   1875
            TabIndex        =   20
            Top             =   555
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin VB.Label LblTipoTrib 
            Caption         =   "Tipo de Tributação:"
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
            Left            =   120
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   57
            Top             =   570
            Width           =   1785
         End
         Begin VB.Label DescNatOp 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2475
            TabIndex        =   56
            Top             =   225
            Width           =   6585
         End
         Begin VB.Label DescTipoTrib 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2475
            TabIndex        =   55
            Top             =   555
            Width           =   6585
         End
         Begin VB.Label NaturezaLabel 
            AutoSize        =   -1  'True
            Caption         =   "Natureza Operação:"
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
            Left            =   90
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   54
            Top             =   240
            Width           =   1725
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Complemento para Contabilização"
         Height          =   945
         Left            =   0
         TabIndex        =   40
         Top             =   2880
         Width           =   9120
         Begin VB.ComboBox Historico 
            Height          =   315
            Left            =   1110
            TabIndex        =   16
            Top             =   225
            Width           =   7980
         End
         Begin MSMask.MaskEdBox ContaContabil 
            Height          =   315
            Left            =   7215
            TabIndex        =   18
            Top             =   570
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Ccl 
            Height          =   315
            Left            =   1110
            TabIndex        =   17
            Top             =   570
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   556
            _Version        =   393216
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
         Begin VB.Label CclLabel 
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
            Left            =   750
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   47
            Top             =   645
            Width           =   345
         End
         Begin VB.Label HistoricoLabel 
            AutoSize        =   -1  'True
            Caption         =   "Histórico:"
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
            TabIndex        =   42
            Top             =   270
            Width           =   825
         End
         Begin VB.Label ContaContabilLabel 
            AutoSize        =   -1  'True
            Caption         =   "Conta Contabil Crédito:"
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
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   41
            ToolTipText     =   "Conta Contábil de Aplicação"
            Top             =   645
            Width           =   1980
         End
      End
   End
   Begin VB.PictureBox Picture 
      Height          =   555
      Left            =   7335
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   30
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ContratoCadastroOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "ContratoCadastroOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "ContratoCadastroOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "ContratoCadastroOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5640
      Left            =   150
      TabIndex        =   49
      Top             =   300
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   9948
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Itens"
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
Attribute VB_Name = "ContratoCadastroOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim iAlterado As Integer
Dim iClienteAlterado As Integer
Dim iContratoAlterado As Integer

Private iFrameAtual As Integer

'HElp
Const IDH_RASTROCONTRATO = 0

'Property Variables:
Dim m_Caption As String
Event Unload()

'Evento de browser
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoCondPagto As AdmEvento
Attribute objEventoCondPagto.VB_VarHelpID = -1
Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1
Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1
Private WithEvents objEventoMedicao As AdmEvento
Attribute objEventoMedicao.VB_VarHelpID = -1
Private WithEvents objEventoNF As AdmEvento
Attribute objEventoNF.VB_VarHelpID = -1
Private WithEvents objEventoNaturezaOp As AdmEvento
Attribute objEventoNaturezaOp.VB_VarHelpID = -1
Private WithEvents objEventoTiposDeTributacao As AdmEvento
Attribute objEventoTiposDeTributacao.VB_VarHelpID = -1

Dim objGridItens As AdmGrid
Dim iGrid_Produto_Col As Integer
Dim iGrid_Descricao_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_Cobrar_Col As Integer
Dim iGrid_Medicao_Col As Integer
Dim iGrid_DataInicio_Col As Integer
Dim iGrid_Periodicidade_Col As Integer
Dim iGrid_CondPagto_Col As Integer
Dim iGrid_Item_Col As Integer
Dim iGrid_DataProx_Col As Integer
Dim iGrid_DataRefIni_Col As Integer
Dim iGrid_DataRefFim_Col As Integer
Dim iGrid_Obs_Col As Integer
Dim iGrid_UM_Col As Integer
Dim iGrid_QtdeParcelas_Col As Integer
Dim iGrid_UltParcCobrada_Col As Integer

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_Itens = 2

Private Sub Ativo_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub BotaoProduto_Click()

Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim lErro As Long
Dim colSelecao As Collection
Dim sProduto1 As String

On Error GoTo Erro_BotaoProdutos_Click

    If Me.ActiveControl Is Produto Then
    
        sProduto1 = Produto.Text
        
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridItens.Row = 0 Then gError 129406

        sProduto1 = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
        
    End If
    
    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 129407
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    'preenche o codigo do produto
    objProduto.sCodigo = sProduto
    
    'Chama a tela de browse ProdutoVendaLista
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoProduto)

    Exit Sub
        
Erro_BotaoProdutos_Click:
    
    Select Case gErr
        
        Case 129406
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 129407 'Tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154973)

    End Select

    Exit Sub

End Sub

Private Sub BotaoServicos_Click()

Dim objItensDeContratoSrv As New ClassItensDeContratoSrv
Dim colSelecao As New Collection
    
    colSelecao.Add Codigo.Text

    Call Chama_Tela("ItensDeContratoSRVLista", colSelecao, objItensDeContratoSrv, objEventoCodigo, "Contrato = ?")
    
End Sub

Private Sub CodigoLabel_Click()

Dim objContrato As New ClassContrato
Dim colSelecao As New Collection

    If Len(Trim(Codigo.Text)) > 0 Then
        objContrato.sCodigo = Codigo.Text
        objContrato.iFilialEmpresa = giFilialEmpresa
    End If
    
    Call Chama_Tela("ContratosLista", colSelecao, objContrato, objEventoCodigo)

End Sub

Private Sub CondPagto_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CondPagtoLabel_Click()

Dim objCondicaoPagto As New ClassCondicaoPagto
Dim colSelecao As New Collection

    'Se Condição de Pagto estiver preenchida, extrai o código
    objCondicaoPagto.iCodigo = CondPagto_Extrai(CondPagto)

    'Chama a Tela CondicoesPagamentoCRLista
    Call Chama_Tela("CondicaoPagtoCRLista", colSelecao, objCondicaoPagto, objEventoCondPagto)

End Sub

Private Function Carrega_Historico(objCombo As ComboBox) As Long
'Carrega a combo de historicos com os historicos da tabela "HistPadraMovConta"

Dim lErro As Long
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome

On Error GoTo Erro_Carrega_Historico

    'le o Codigo e a descricao de todos os historicos
    lErro = CF("Cod_Nomes_Le", "HistPadraoMovConta", "Codigo", "Descricao", STRING_NOME, colCodigoNome)
    If lErro <> SUCESSO Then gError 129503

    For Each objCodigoNome In colCodigoNome

        'Insere na Combo de historicos
        objCombo.AddItem objCodigoNome.sNome
        objCombo.ItemData(objCombo.NewIndex) = objCodigoNome.iCodigo

    Next

    Carrega_Historico = SUCESSO

    Exit Function

Erro_Carrega_Historico:

    Carrega_Historico = gErr

    Select Case gErr

        Case 129503

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154974)

    End Select
    
    Exit Function

End Function

Private Function Carrega_ComboPeriodicidade(objCombo As ComboBox) As Long

Dim lErro As Long
   
On Error GoTo Erro_Carrega_ComboPeriodicidade
    
    objCombo.AddItem PERIODICIDADE_LIVRE & SEPARADOR & STRING_PERIODICIDADE_LIVRE
    objCombo.ItemData(objCombo.NewIndex) = PERIODICIDADE_LIVRE
    
    objCombo.AddItem PERIODICIDADE_SEMANAL & SEPARADOR & STRING_PERIODICIDADE_SEMANAL
    objCombo.ItemData(objCombo.NewIndex) = PERIODICIDADE_SEMANAL
     
    objCombo.AddItem PERIODICIDADE_DECENDIAL & SEPARADOR & STRING_PERIODICIDADE_DECENDIAL
    objCombo.ItemData(objCombo.NewIndex) = PERIODICIDADE_DECENDIAL
     
    objCombo.AddItem PERIODICIDADE_QUINZENAL & SEPARADOR & STRING_PERIODICIDADE_QUINZENAL
    objCombo.ItemData(objCombo.NewIndex) = PERIODICIDADE_QUINZENAL
    
    objCombo.AddItem PERIODICIDADE_MENSAL & SEPARADOR & STRING_PERIODICIDADE_MENSAL
    objCombo.ItemData(objCombo.NewIndex) = PERIODICIDADE_MENSAL
    
    objCombo.AddItem PERIODICIDADE_BIMESTRAL & SEPARADOR & STRING_PERIODICIDADE_BIMESTRAL
    objCombo.ItemData(objCombo.NewIndex) = PERIODICIDADE_BIMESTRAL
    
    objCombo.AddItem PERIODICIDADE_TRIMESTRAL & SEPARADOR & STRING_PERIODICIDADE_TRIMESTRAL
    objCombo.ItemData(objCombo.NewIndex) = PERIODICIDADE_TRIMESTRAL
    
    objCombo.AddItem PERIODICIDADE_QUADRIMESTRAL & SEPARADOR & STRING_PERIODICIDADE_QUADRIMESTRAL
    objCombo.ItemData(objCombo.NewIndex) = PERIODICIDADE_QUADRIMESTRAL
    
    objCombo.AddItem PERIODICIDADE_SEMESTRAL & SEPARADOR & STRING_PERIODICIDADE_SEMESTRAL
    objCombo.ItemData(objCombo.NewIndex) = PERIODICIDADE_SEMESTRAL
    
    objCombo.AddItem PERIODICIDADE_ANUAL & SEPARADOR & STRING_PERIODICIDADE_ANUAL
    objCombo.ItemData(objCombo.NewIndex) = PERIODICIDADE_ANUAL
    
    Carrega_ComboPeriodicidade = SUCESSO
    
    Exit Function
    
Erro_Carrega_ComboPeriodicidade:

    Carrega_ComboPeriodicidade = gErr
    
    Select Case gErr
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154975)

    End Select

    Exit Function
    
End Function
    
Private Function Move_Tela_Memoria(objContrato As ClassContrato) As Long

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim sCclFormatada As String
Dim iCclPreenchida As Integer
Dim objSerie As New ClassSerie

On Error GoTo Erro_Move_Tela_Memoria

    With objContrato
    
        'Verifica se o Cliente foi preenchido
        If Len(Trim(Cliente.ClipText)) > 0 Then
    
            objcliente.sNomeReduzido = Cliente.Text
    
            'Lê o Cliente através do Nome Reduzido
            lErro = CF("Cliente_Le_NomeReduzido", objcliente)
            If lErro <> SUCESSO And lErro <> 12348 Then gError 129491
    
            If lErro = SUCESSO Then .lCliente = objcliente.lCodigo
                                
        End If
    
        .dtDataFimContrato = StrParaDate(DataFimContrato.Text)
        .dtDataIniCobrancaPadrao = StrParaDate(DataInicioP.Text)
        .dtDataIniContrato = StrParaDate(DataIniContrato.Text)
        .dtDataRenovContrato = StrParaDate(DataRenovContrato.Text)
        
        .iCondPagtoPadrao = CondPagto_Extrai(CondPagtoP)
        .iFilCli = Codigo_Extrai(FilialCliente.Text)
        .iPeriodicidadePadrao = Codigo_Extrai(PeriodicidadeP.Text)
        
        .iAtivo = StrParaInt(Ativo.Value)
        
        .iFilialEmpresa = giFilialEmpresa
        .sCodigo = Codigo.Text
        .sContaContabil = ContaContabil.ClipText
        .sDescricao = Descricao.Text
        .sObservacao = Observacao.Text
        .sHistoricoContabil = Historico.Text
        If Len(Trim(Serie.Text)) > 0 Then .sSerie = Converte_Serie_Eletronica(Serie.Text, optNFe.Value)
        If optNFe.Value = vbChecked Then
            .iNFe = MARCADO
        Else
            .iNFe = DESMARCADO
        End If
        
        lErro = CF("Ccl_Formata", Ccl.Text, sCclFormatada, iCclPreenchida)
        If lErro <> SUCESSO Then gError 129560
        
        .sCcl = sCclFormatada
        
        .sNaturezaOp = Trim(NaturezaOp.Text)
        .iTipoTributacao = StrParaInt(TipoTributacao.Text)
        
        .iTipo = CONTRATOS_RECEBER
        
        If optRecibo.Value = vbChecked Then
            .iRecibo = MARCADO
        Else
            .iRecibo = DESMARCADO
        End If
        
    End With
    
    If Len(Trim(objContrato.sSerie)) > 0 Then
        objSerie.sSerie = objContrato.sSerie
        'Busca a série no BD
        lErro = CF("Serie_Le", objSerie)
        If lErro <> SUCESSO And lErro <> 22202 Then gError 129560
        If lErro <> SUCESSO Then gError 129561 'Se não encontrou
    End If
    
    Call Move_GridItens_Memoria(objContrato)

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
        
        Case 129491, 129560
        
        Case 129561
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, objContrato.sSerie)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154976)

    End Select

    Exit Function

End Function

Private Function Soma_Valor() As Double

Dim iIndice As Integer
Dim dSoma As Double

    dSoma = 0

    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        If StrParaInt(GridItens.TextMatrix(iIndice, iGrid_Cobrar_Col)) = MARCADO Then
            dSoma = dSoma + (StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Valor_Col)) * StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col)))
        End If
    
    Next

    Total.Caption = Format(dSoma, "Standard")

End Function

Private Function Move_GridItens_Memoria(objContrato As ClassContrato) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objItensDeContrato As ClassItensDeContrato
Dim sProduto As String
Dim iPreenchido As Integer

On Error GoTo Erro_Move_GridItens_Memoria

    For iIndice = 1 To objGridItens.iLinhasExistentes

        Set objItensDeContrato = New ClassItensDeContrato

        With objItensDeContrato
        
            .iQtdeParcelas = StrParaInt(GridItens.TextMatrix(iIndice, iGrid_QtdeParcelas_Col))
            .iUltParcCobrada = StrParaInt(GridItens.TextMatrix(iIndice, iGrid_UltParcCobrada_Col))
            .dtDataRefIni = StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataRefIni_Col))
            .dtDataRefFim = StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataRefFim_Col))
            .dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
            .dtDataIniCobranca = StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataInicio_Col))
            .dtDataProxCobranca = StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataProx_Col))
            .dValor = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Valor_Col))
            .iCobrar = StrParaInt(GridItens.TextMatrix(iIndice, iGrid_Cobrar_Col))
            .iMedicao = StrParaInt(GridItens.TextMatrix(iIndice, iGrid_Medicao_Col))
            .iSeq = StrParaInt(GridItens.TextMatrix(iIndice, iGrid_Item_Col))
            
            CondPagto.Text = GridItens.TextMatrix(iIndice, iGrid_CondPagto_Col)

            .iCondPagto = CondPagto_Extrai(CondPagto)
            .iPeriodicidade = Codigo_Extrai(GridItens.TextMatrix(iIndice, iGrid_Periodicidade_Col))
                       
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Produto_Col))) > 0 Then
                lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProduto, iPreenchido)
                If lErro <> SUCESSO Then gError 129521
                .sProduto = sProduto
            End If
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Produto_Col))) = 0 Or iPreenchido = PRODUTO_VAZIO Then gError 129520
            
            .sDescProd = GridItens.TextMatrix(iIndice, iGrid_Descricao_Col)
            .sObservacao = GridItens.TextMatrix(iIndice, iGrid_Obs_Col)
            .sUM = GridItens.TextMatrix(iIndice, iGrid_UM_Col)
                                              
            '??? acho que está sobrando
            .dtDataRefIni = StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataRefIni_Col))
            .dtDataRefFim = StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataRefFim_Col))
                                              
        End With
            
        objContrato.colItens.Add objItensDeContrato
    
    Next

    Move_GridItens_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItens_Memoria:

    Move_GridItens_Memoria = gErr

    Select Case gErr
    
        Case 129520
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 129521

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154977)

    End Select

    Exit Function

End Function

Private Function Valida_Grid_Itens() As Long

Dim iIndice As Integer
Dim lErro As Long
Dim dQuantidade As Double

On Error GoTo Erro_Valida_Grid_Itens

    'Verifica se há itens no grid
    If objGridItens.iLinhasExistentes = 0 Then gError 129570

    'para cada item do grid
    For iIndice = 1 To objGridItens.iLinhasExistentes

        'Valida Quantidade
        If StrParaInt(GridItens.TextMatrix(iIndice, iGrid_Medicao_Col)) <> MARCADO Then
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 129571
        
            lErro = Valor_NaoNegativo_Critica(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
            If lErro <> SUCESSO Then gError 129572

        End If

        'Valida Valor
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Valor_Col))) = 0 Then gError 129573

        lErro = Valor_Positivo_Critica(GridItens.TextMatrix(iIndice, iGrid_Valor_Col))
        If lErro <> SUCESSO Then gError 129574

        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_CondPagto_Col))) = 0 Then gError 133580

        'Valida Data Inicial
        If StrParaDate(DataIniContrato.Text) > StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataInicio_Col)) Then gError 129832

        If StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataProx_Col)) = DATA_NULA Then gError 136042
        
        If StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataRefIni_Col)) = DATA_NULA Then gError 136043
        If StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataRefFim_Col)) = DATA_NULA Then gError 136044
        If StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataRefFim_Col)) < StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataRefIni_Col)) Then gError 136055
        If StrParaInt(GridItens.TextMatrix(iIndice, iGrid_QtdeParcelas_Col)) < StrParaInt(GridItens.TextMatrix(iIndice, iGrid_UltParcCobrada_Col)) Then gError 201204
    
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Descricao_Col))) = 0 Then gError 201581
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_UM_Col))) = 0 Then gError 201582
        
    Next

    Valida_Grid_Itens = SUCESSO

    Exit Function

Erro_Valida_Grid_Itens:

    Valida_Grid_Itens = gErr

    Select Case gErr

        Case 129570
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ITENS_CONTRATOS", gErr)

        Case 129571
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_ITEM_NAO_PREENCHIDA", gErr, iIndice)

        Case 129572, 129574

        Case 129573
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORUNITARIO_ITEM_NAO_PREENCHIDO", gErr, iIndice)

        Case 129832
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICOBRANCA_MENOR_DATAINICONTRATO", gErr, iIndice)

        Case 133580
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDPAGTO_ITEM_NAO_PREENCHIDO", gErr, iIndice)

        Case 136042
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAPROXCOBRANCA_NAO_PREENCHIDO", gErr, iIndice)
        
        Case 136043
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAREFINI_NAO_PREENCHIDO", gErr, iIndice)
        
        Case 136044
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAREFFIM_NAO_PREENCHIDO", gErr, iIndice)
        
        Case 136055
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAREFINI_MAIOR_DATAREFFIM", gErr, iIndice)
            
        Case 201204
            Call Rotina_Erro(vbOKOnly, "ERRO_QTDEPARC_MENOR_ULTPARC", gErr, iIndice)
        
        Case 201581
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_ITEMCONTRATO_NAO_PREENCHIDA", gErr, iIndice)

        Case 201582
            Call Rotina_Erro(vbOKOnly, "ERRO_UM_ITEMCONTRATO_NAO_PREENCHIDA", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154978)

    End Select

    Exit Function

End Function

Private Function ProdutoLinha_Preenche(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim iSeq As Integer
Dim iIndice As Integer

On Error GoTo Erro_ProdutoLinha_Preenche

    For iIndice = 1 To objGridItens.iLinhasExistentes
        If iIndice <> GridItens.Row Then
            If GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text Then gError 129540
        End If
    Next

    GridItens.TextMatrix(GridItens.Row, iGrid_Descricao_Col) = objProduto.sDescricao
    GridItens.TextMatrix(GridItens.Row, iGrid_UM_Col) = objProduto.sSiglaUMEstoque
                     
    If GridItens.Row > 1 Then
        iSeq = StrParaInt(GridItens.TextMatrix(GridItens.Row - 1, iGrid_Item_Col)) + 1
    Else
        iSeq = 1
    End If
        
    GridItens.TextMatrix(GridItens.Row, iGrid_Item_Col) = iSeq
                    
    GridItens.TextMatrix(GridItens.Row, iGrid_Cobrar_Col) = MARCADO
                    
    If Len(Trim(DataInicioP.Text)) > 0 Then
        GridItens.TextMatrix(GridItens.Row, iGrid_DataInicio_Col) = DataInicioP.Text
    End If

    If Len(Trim(PeriodicidadeP.Text)) > 0 Then
        GridItens.TextMatrix(GridItens.Row, iGrid_Periodicidade_Col) = PeriodicidadeP.Text
    End If

    If Len(Trim(CondPagtoP.Text)) > 0 Then
        GridItens.TextMatrix(GridItens.Row, iGrid_CondPagto_Col) = CondPagtoP.Text
    End If

    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_DataInicio_Col))) > 0 And Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_Periodicidade_Col)) <> 0 Then
        Call Atualiza_DataProxCobranca(StrParaDate(GridItens.TextMatrix(GridItens.Row, iGrid_DataInicio_Col)), Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_Periodicidade_Col)))
    End If
    
    GridItens.TextMatrix(GridItens.Row, iGrid_DataRefIni_Col) = Format(DataIniContrato.Text, "dd/mm/yyyy")
    GridItens.TextMatrix(GridItens.Row, iGrid_DataRefFim_Col) = ""
    GridItens.TextMatrix(GridItens.Row, iGrid_QtdeParcelas_Col) = ""
    GridItens.TextMatrix(GridItens.Row, iGrid_UltParcCobrada_Col) = ""
       
    Call Grid_Refresh_Checkbox(objGridItens)
       
    'ALTERAÇÃO DE LINHAS EXISTENTES
    If (GridItens.Row - GridItens.FixedRows) = objGridItens.iLinhasExistentes Then
        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
    End If
    
    ProdutoLinha_Preenche = SUCESSO

    Exit Function

Erro_ProdutoLinha_Preenche:

    ProdutoLinha_Preenche = gErr

    Select Case gErr

        Case 129540
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_REPETIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 154979)

    End Select

    Exit Function

End Function

Private Function Carrega_CondicaoPagamento(objCombo As ComboBox) As Long

Dim lErro As Long
Dim colCod_DescReduzida As New AdmColCodigoNome
Dim objCod_DescReduzida As AdmCodigoNome

On Error GoTo Erro_Carrega_CondicaoPagamento

    'Lê o código e a descrição reduzida de todas as Condições de Pagamento
    lErro = CF("CondicoesPagto_Le_Recebimento", colCod_DescReduzida)
    If lErro <> SUCESSO Then gError 129461

   For Each objCod_DescReduzida In colCod_DescReduzida

        'Adiciona novo item na List da Combo CondicaoPagamento
        objCombo.AddItem CInt(objCod_DescReduzida.iCodigo) & SEPARADOR & objCod_DescReduzida.sNome
        objCombo.ItemData(objCombo.NewIndex) = objCod_DescReduzida.iCodigo

    Next

    Carrega_CondicaoPagamento = SUCESSO

    Exit Function

Erro_Carrega_CondicaoPagamento:

    Carrega_CondicaoPagamento = gErr

    Select Case gErr

        Case 129461

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154980)

    End Select

    Exit Function

End Function

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iFrameAtual = 1

    Set objGridItens = New AdmGrid
    
    Set objEventoCliente = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoCondPagto = New AdmEvento
    Set objEventoCodigo = New AdmEvento
    Set objEventoContaContabil = New AdmEvento
    Set objEventoCcl = New AdmEvento
    Set objEventoNaturezaOp = New AdmEvento
    Set objEventoTiposDeTributacao = New AdmEvento
    Set objEventoMedicao = New AdmEvento
    Set objEventoNF = New AdmEvento
    
    If Len(Trim(gobjFAT.sFormatoPrecoUnitario)) <> 0 Then
        Valor.Format = gobjFAT.sFormatoPrecoUnitario
    End If
    
    Total.Caption = Format("0", "Standard")
    
    DataIniContrato.PromptInclude = False
    DataIniContrato.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataIniContrato.PromptInclude = True
   
    'Preenche Lista da Combobox
    lErro = Carrega_ComboPeriodicidade(Periodicidade)
    If lErro <> SUCESSO Then gError 129458
    
    lErro = Carrega_ComboPeriodicidade(PeriodicidadeP)
    If lErro <> SUCESSO Then gError 129459
    
    lErro = CF("Carrega_CondicaoPagamento", CondPagto, MODULO_CONTASARECEBER)
    If lErro <> SUCESSO Then gError 129460
    
    lErro = CF("Carrega_CondicaoPagamento", CondPagtoP, MODULO_CONTASARECEBER)
    If lErro <> SUCESSO Then gError 129462
    
    lErro = Carrega_Serie
    If lErro <> SUCESSO Then gError 129462
    
    lErro = Carrega_Historico(Historico)
    If lErro <> SUCESSO Then gError 136075
    
    lErro = Inicializa_GridItens(objGridItens)
    If lErro <> SUCESSO Then gError 131046

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 129527
   
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabil)
    If lErro <> SUCESSO Then gError 131051
   
    lErro = CF("Inicializa_Mascara_Ccl_MaskEd", Ccl)
    If lErro <> SUCESSO Then gError 131052

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 131046, 129458, 129459, 129460, 129462, 129527, 131051, 131052, 136075

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154981)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, UnloadMode, Cancel, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objGridItens = Nothing
    Set objEventoCliente = Nothing
    Set objEventoCondPagto = Nothing
    Set objEventoProduto = Nothing
    Set objEventoCodigo = Nothing
    Set objEventoContaContabil = Nothing
    Set objEventoCcl = Nothing
    Set objEventoNaturezaOp = Nothing
    Set objEventoTiposDeTributacao = Nothing
    Set objEventoMedicao = Nothing
    Set objEventoNF = Nothing
    
    'Fecha o Comando de Setas
    Call ComandoSeta_Liberar(Me.Name)


End Sub

Private Function Inicializa_GridItens(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Alocação

Dim iIndice As Integer

    Set objGridItens.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("UM")
    objGridInt.colColuna.Add ("Valor Unitário")
    objGridInt.colColuna.Add ("Cobrar")
    objGridInt.colColuna.Add ("Com Medição")
    objGridInt.colColuna.Add ("Dt Início Cobrança")
    objGridInt.colColuna.Add ("Periodicidade")
    objGridInt.colColuna.Add ("Dt Próxima Cobrança")
    objGridInt.colColuna.Add ("Cond. Pagto")
    objGridInt.colColuna.Add ("Dt Ref Ini")
    objGridInt.colColuna.Add ("Dt Ref Fim")
    objGridInt.colColuna.Add ("Observação")
    objGridInt.colColuna.Add ("Qtd Parcelas")
    objGridInt.colColuna.Add ("Ult. Cobrada")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Item.Name)
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (ItemDescricao.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (UM.Name)
    objGridInt.colCampo.Add (Valor.Name)
    objGridInt.colCampo.Add (Cobrar.Name)
    objGridInt.colCampo.Add (Medicao.Name)
    objGridInt.colCampo.Add (DataInicio.Name)
    objGridInt.colCampo.Add (Periodicidade.Name)
    objGridInt.colCampo.Add (DataProx.Name)
    objGridInt.colCampo.Add (CondPagto.Name)
    objGridInt.colCampo.Add (DataRefIni.Name)
    objGridInt.colCampo.Add (DataRefFim.Name)
    objGridInt.colCampo.Add (Obs.Name)
    objGridInt.colCampo.Add (QtdeParcelas.Name)
    objGridInt.colCampo.Add (UltParcCobrada.Name)

    'Colunas da Grid
    iGrid_Item_Col = 1
    iGrid_Produto_Col = 2
    iGrid_Descricao_Col = 3
    iGrid_Quantidade_Col = 4
    iGrid_UM_Col = 5
    iGrid_Valor_Col = 6
    iGrid_Cobrar_Col = 7
    iGrid_Medicao_Col = 8
    iGrid_DataInicio_Col = 9
    iGrid_Periodicidade_Col = 10
    iGrid_DataProx_Col = 11
    iGrid_CondPagto_Col = 12
    iGrid_DataRefIni_Col = 13
    iGrid_DataRefFim_Col = 14
    iGrid_Obs_Col = 15
    iGrid_QtdeParcelas_Col = 16
    iGrid_UltParcCobrada_Col = 17
 
    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 500

    'Habilita a execução da Rotina_Grid_Enable
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = 6000

    objGridInt.objGrid.Rows = 500

    objGridInt.iLinhasVisiveis = 10
       
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridItens)

    Inicializa_GridItens = SUCESSO

    Exit Function

End Function

Public Function Trata_Parametros(Optional objContrato As ClassContrato) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objContrato Is Nothing) Then
        
        lErro = Traz_Contrato_Tela(objContrato)
        If lErro <> SUCESSO Then gError 129518

    End If

    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
    
        Case 129518
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154982)

    End Select

    Exit Function

End Function

Private Sub Limpa_Tela_Contrato()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Contrato

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Call Limpa_Tela(Me)
    
    Historico.Text = ""
    FilialCliente.Clear
    CondPagtoP.Text = ""
    PeriodicidadeP.Text = ""
    Serie.Text = ""
    
    optRecibo.Value = vbUnchecked
    optNFe.Value = vbUnchecked
    
    Total.Caption = Format("0", "Standard")
    
    DataIniContrato.PromptInclude = False
    DataIniContrato.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataIniContrato.PromptInclude = True
    
    DescNatOp.Caption = ""
    DescTipoTrib.Caption = ""
        
    Call Grid_Limpa(objGridItens)
    
    iAlterado = 0
    iClienteAlterado = 0
    iContratoAlterado = 0
    
    Cliente.Enabled = True
    FilialCliente.Enabled = True
    ClienteLabel.Enabled = True
    FilialClienteLabel.Enabled = True
    
    Exit Sub

Erro_Limpa_Tela_Contrato:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154983)

    End Select

    Exit Sub

End Sub

Private Function Carrega_Grid_Itens(objContrato As ClassContrato) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objItensDeContrato As ClassItensDeContrato
Dim sProduto As String
Dim sProdutoAux As String

On Error GoTo Erro_Carrega_Grid_Itens

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridItens)

    For Each objItensDeContrato In objContrato.colItens
           
        With objItensDeContrato
            
            iIndice = iIndice + 1
           
            sProdutoAux = objItensDeContrato.sProduto
           
            lErro = Mascara_RetornaProdutoTela(sProdutoAux, sProduto)
            If lErro <> SUCESSO Then gError 129523
    
            'Mascara o produto enxuto
            Produto.PromptInclude = False
            Produto.Text = sProduto
            Produto.PromptInclude = True

            GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
            GridItens.TextMatrix(iIndice, iGrid_Descricao_Col) = .sDescProd
                   
            GridItens.TextMatrix(iIndice, iGrid_Cobrar_Col) = .iCobrar
            GridItens.TextMatrix(iIndice, iGrid_Medicao_Col) = .iMedicao
            
            If .iCondPagto <> 0 Then
                CondPagto.Text = .iCondPagto
                lErro = Combo_Seleciona_Grid(CondPagto, .iCondPagto)
                If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 129564
                GridItens.TextMatrix(iIndice, iGrid_CondPagto_Col) = CondPagto.Text
            End If
            
            If .iPeriodicidade <> 0 Then
                Periodicidade.Text = .iPeriodicidade
                lErro = Combo_Seleciona_Grid(Periodicidade, .iPeriodicidade)
                If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 129565
                GridItens.TextMatrix(iIndice, iGrid_Periodicidade_Col) = Periodicidade.Text
            End If
            If .dtDataIniCobranca <> DATA_NULA Then GridItens.TextMatrix(iIndice, iGrid_DataInicio_Col) = Format(.dtDataIniCobranca, "dd/mm/yyyy")
            If .dtDataProxCobranca <> DATA_NULA Then GridItens.TextMatrix(iIndice, iGrid_DataProx_Col) = Format(.dtDataProxCobranca, "dd/mm/yyyy")
            
            GridItens.TextMatrix(iIndice, iGrid_Obs_Col) = .sObservacao
            If .dQuantidade <> 0 Then GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(.dQuantidade)
            GridItens.TextMatrix(iIndice, iGrid_UM_Col) = .sUM
            GridItens.TextMatrix(iIndice, iGrid_Valor_Col) = Format(.dValor, Valor.Format)
            GridItens.TextMatrix(iIndice, iGrid_Item_Col) = .iSeq
            
            If .dtDataRefIni <> DATA_NULA Then GridItens.TextMatrix(iIndice, iGrid_DataRefIni_Col) = Format(.dtDataRefIni, "dd/mm/yyyy")
            If .dtDataRefFim <> DATA_NULA Then GridItens.TextMatrix(iIndice, iGrid_DataRefFim_Col) = Format(.dtDataRefFim, "dd/mm/yyyy")
            
            If .iQtdeParcelas > 0 Then
                GridItens.TextMatrix(iIndice, iGrid_QtdeParcelas_Col) = CStr(.iQtdeParcelas)
                GridItens.TextMatrix(iIndice, iGrid_UltParcCobrada_Col) = CStr(.iUltParcCobrada)
            End If
        
        End With
            
    Next
       
    Call Grid_Refresh_Checkbox(objGridItens)

    'Atualiza o número de linhas existentes
    objGridItens.iLinhasExistentes = iIndice
    
    Call Soma_Valor
    
    Carrega_Grid_Itens = SUCESSO
        
    Exit Function

Erro_Carrega_Grid_Itens:

    Call Grid_Limpa(objGridItens)
    
    Carrega_Grid_Itens = gErr

    Select Case gErr
    
        Case 129523, 129564, 129565
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154984)

    End Select

    Exit Function

End Function

Public Sub Cliente_Formata(lCliente As Long)

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Formata

    Cliente.Text = lCliente

    'Busca o Cliente no BD
    lErro = TP_Cliente_Le(Cliente, objcliente, iCodFilial)
    If lErro <> SUCESSO Then gError 129500

    lErro = CF("FiliaisClientes_Le_Cliente", objcliente, colCodigoNome)
    If lErro <> SUCESSO Then gError 129501

    'Preenche ComboBox de Filiais
    Call CF("Filial_Preenche", FilialCliente, colCodigoNome)

    iClienteAlterado = 0
    
    Exit Sub

Erro_Cliente_Formata:

    Select Case gErr

        Case 129500, 129501

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154985)

    End Select

    Exit Sub

End Sub

Public Sub Filial_Formata(objFilial As Object, iFilial As Integer)

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Formata

    objFilial.Text = CStr(iFilial)
    sCliente = Cliente.Text
    objFilialCliente.iCodFilial = iFilial

    'Pesquisa se existe Filial com o código extraído
    lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
    If lErro <> SUCESSO And lErro <> 17660 Then gError 129498

    If lErro = 17660 Then gError 129499

    'Coloca na tela a Filial lida
    objFilial.Text = iFilial & SEPARADOR & objFilialCliente.sNome

    Exit Sub

Erro_Filial_Formata:

    Select Case gErr

        Case 129498

        Case 129499
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, objFilial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154986)

    End Select

    Exit Sub

End Sub

Private Function Traz_Contrato_Tela(objContrato As ClassContrato) As Long

Dim lErro As Long
Dim sCclMascarado As String

On Error GoTo Erro_Traz_Contrato_Tela

    If objContrato.lNumIntDoc = 0 Then
    
        lErro = CF("Contrato_Le", objContrato)
        If lErro <> SUCESSO And lErro <> 129332 Then gError 129404
        
    Else
    
        lErro = CF("Contrato_Le2", objContrato)
        If lErro <> SUCESSO And lErro <> 129261 Then gError 136187
    
    End If
        
    'Se o contrato é novo não tenta trazer os dados
    If lErro <> SUCESSO Then
        Cliente.Enabled = True
        FilialCliente.Enabled = True
        ClienteLabel.Enabled = True
        FilialClienteLabel.Enabled = True
        Exit Function
    Else
        
        If objContrato.iTipo <> CONTRATOS_RECEBER Then gError 132900
        
        Call Limpa_Tela_Contrato
        
        Cliente.Enabled = False
        FilialCliente.Enabled = False
        ClienteLabel.Enabled = False
        FilialClienteLabel.Enabled = False
    End If
           
    With objContrato
    
        If .dtDataFimContrato <> DATA_NULA Then DataFimContrato.Text = Format(.dtDataFimContrato, "dd/mm/yy")
        If .dtDataIniCobrancaPadrao <> DATA_NULA Then DataInicioP.Text = Format(.dtDataIniCobrancaPadrao, "dd/mm/yy")
        If .dtDataIniContrato <> DATA_NULA Then DataIniContrato.Text = Format(.dtDataIniContrato, "dd/mm/yy")
        If .dtDataRenovContrato <> DATA_NULA Then DataRenovContrato.Text = Format(.dtDataRenovContrato, "dd/mm/yy")
        
        If .lCliente <> 0 Then
            Call Cliente_Formata(.lCliente)
            Call Filial_Formata(FilialCliente, .iFilCli)
        End If
        
        If .iCondPagtoPadrao > 0 Then
            CondPagtoP.Text = .iCondPagtoPadrao
            Call CondPagtoP_Validate(bSGECancelDummy)
        Else
            CondPagtoP.Text = ""
        End If
        
        If .iPeriodicidadePadrao > 0 Then
            PeriodicidadeP.Text = .iPeriodicidadePadrao
            Call PeriodicidadeP_Validate(bSGECancelDummy)
        Else
            PeriodicidadeP.Text = ""
        End If
        
        Ativo.Value = .iAtivo
        Codigo.Text = .sCodigo
        
        ContaContabil.PromptInclude = False
        ContaContabil.Text = .sContaContabil
        ContaContabil.PromptInclude = True
        
        Descricao.Text = .sDescricao
        Observacao.Text = .sObservacao
        
        
        If Len(Trim(.sCcl)) > 0 Then
        
            'Preenche a CCL
            sCclMascarado = String(STRING_CCL, 0)
        
            lErro = Mascara_MascararCcl(.sCcl, sCclMascarado)
            If lErro <> SUCESSO Then gError 129561
        
            Ccl.PromptInclude = False
            Ccl.Text = sCclMascarado
            Ccl.PromptInclude = True
        
        End If
        
        Historico.Text = .sHistoricoContabil
            
    End With

    NaturezaOp.Text = objContrato.sNaturezaOp
    Call NaturezaOp_Validate(bSGECancelDummy)
    
    If objContrato.iTipoTributacao <> 0 Then TipoTributacao.Text = objContrato.iTipoTributacao
    Call TipoTributacao_Validate(bSGECancelDummy)
    
    If objContrato.iRecibo = MARCADO Then
        optRecibo.Value = vbChecked
    Else
        optRecibo.Value = vbUnchecked
    End If
    
    If Len(Trim(objContrato.sSerie)) > 0 Then
        
        Serie.Text = Desconverte_Serie_Eletronica(objContrato.sSerie)
        If ISSerieEletronica(objContrato.sSerie) Then
            optNFe.Value = vbChecked
        Else
            optNFe.Value = vbUnchecked
        End If
    Else
        Serie.Text = ""
        optNFe.Value = vbUnchecked
    End If

    lErro = Carrega_Grid_Itens(objContrato)
    If lErro <> SUCESSO Then gError 129566
     
    iAlterado = 0
    iClienteAlterado = 0
    iContratoAlterado = 0
    
    Traz_Contrato_Tela = SUCESSO

    Exit Function

Erro_Traz_Contrato_Tela:
    
    Traz_Contrato_Tela = gErr

    Select Case gErr
    
        Case 129404, 129561, 129566, 136187
        
        Case 132900
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_NAO_RECEBER", gErr, objContrato.sCodigo, objContrato.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154987)

    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RASTROCONTRATO
    Set Form_Load_Ocx = Me
    Caption = "Cadastro de Contrato a Receber"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ContratoCadastro"

End Function

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

Public Sub Ccl_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CclLabel_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim sCclFormatada As String
Dim iCclPreenchida As Integer
Dim objCcl As New ClassCcl

On Error GoTo Erro_LabelCcl_Click

    'Critica o formato do centro de custo
    lErro = CF("Ccl_Formata", Ccl.Text, sCclFormatada, iCclPreenchida)
    If lErro <> SUCESSO Then gError 129562

    objCcl.sCcl = sCclFormatada

    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCcl)

    Exit Sub

Erro_LabelCcl_Click:

    Select Case gErr

        Case 129562

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154988)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCcl_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclMascarado As String

On Error GoTo Erro_objEventoCcl_evSelecao

    Set objCcl = obj1

    sCclMascarado = String(STRING_CCL, 0)

    'Coloca a conta no formato conta enxuta
    lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then gError 129563

    Ccl.PromptInclude = False
    Ccl.Text = sCclMascarado
    Ccl.PromptInclude = True

    Exit Sub

Erro_objEventoCcl_evSelecao:

    Select Case gErr

        Case 129563

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154989)

        End Select

    Exit Sub

End Sub

Public Sub Ccl_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sCclFormatada As String
Dim objCcl As New ClassCcl
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Ccl_Validate

    If Len(Trim(Ccl.ClipText)) <> 0 Then

        lErro = CF("Ccl_Critica", Ccl.Text, sCclFormatada, objCcl)
        If lErro <> SUCESSO And lErro <> 5703 Then gError 129438

        If lErro = 5703 Then gError 129439

    End If

    Exit Sub

Erro_Ccl_Validate:

    Cancel = True

    Select Case gErr

        Case 129438

        Case 129439
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", Ccl.Text)

            If vbMsg = vbYes Then
                objCcl.sCcl = sCclFormatada
                Call Chama_Tela("CclTela", objCcl)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 154990)

    End Select

    Exit Sub
End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO
    iContratoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objContrato As New ClassContrato

On Error GoTo Erro_Codigo_Validate

    If Len(Trim(Codigo.Text)) = 0 Then Exit Sub

    objContrato.sCodigo = Codigo.Text
    objContrato.iFilialEmpresa = giFilialEmpresa
    
    If iContratoAlterado = REGISTRO_ALTERADO Then

        lErro = Traz_Contrato_Tela(objContrato)
        If lErro <> SUCESSO Then gError 129405
        
    End If
        
    Call ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 129405

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154991)

    End Select

End Sub

Private Sub CondPagtoP_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CondPagtoP_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CondPagtoP_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim iCodigo As Integer

On Error GoTo Erro_CondPagtoP_Validate

    'Verifica se a Condicaopagamento foi preenchida
    If Len(Trim(CondPagtoP.Text)) = 0 Then Exit Sub

    'Verifica se é uma Condicaopagamento selecionada
    If CondPagtoP.Text = CondPagtoP.List(CondPagtoP.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(CondPagtoP, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 129472

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objCondicaoPagto.iCodigo = iCodigo

        'Tenta ler CondicaoPagto com esse código no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 129720
        
        'Não encontrou CondicaoPagto no BD
        If lErro <> SUCESSO Then gError 129721

        'Encontrou CondicaoPagto no BD e não é de Recebimento
        If objCondicaoPagto.iEmRecebimento = 0 Then gError 129722

        'Coloca no Text da Combo
        CondPagtoP.Text = CondPagto_Traz(objCondicaoPagto)

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 129723

    Exit Sub

Erro_CondPagtoP_Validate:

    Cancel = True

    Select Case gErr
    
        Case 129472, 129720
        
        Case 129721

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAO_PAGAMENTO")

            If vbMsgRes = vbYes Then
                'Chama a tela de CondicaoPagto
                Call Chama_Tela("CondicoesPagto", objCondicaoPagto)

            End If

        Case 129722
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_RECEBIMENTO", gErr, iCodigo)

        Case 129723
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", gErr, CondPagto.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154992)

    End Select

    Exit Sub

End Sub

Private Sub ContaContabil_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataFimContrato_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataIniContrato_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataInicioP_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataRenovContrato_GotFocus()

     Call MaskEdBox_TrataGotFocus(DataRenovContrato, iAlterado)

End Sub

Private Sub DataRenovContrato_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataRenovContrato_Validate

    'Verifica se a Data de Emissao foi digitada
    If Len(Trim(DataRenovContrato.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataRenovContrato.Text)
    If lErro <> SUCESSO Then gError 129580

    Exit Sub

Erro_DataRenovContrato_Validate:

    Cancel = True

    Select Case gErr

        'se houve erro de crítica, segura o foco
        Case 129580

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154993)

    End Select

    Exit Sub

End Sub

Private Sub DataIniContrato_GotFocus()

     Call MaskEdBox_TrataGotFocus(DataIniContrato, iAlterado)

End Sub

Private Sub DataIniContrato_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataIniContrato_Validate

    'Verifica se a Data de Emissao foi digitada
    If Len(Trim(DataIniContrato.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataIniContrato.Text)
    If lErro <> SUCESSO Then gError 129581

    Exit Sub

Erro_DataIniContrato_Validate:

    Cancel = True

    Select Case gErr

        'se houve erro de crítica, segura o foco
        Case 129581

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154994)

    End Select

    Exit Sub

End Sub

Private Sub DataFimContrato_GotFocus()

     Call MaskEdBox_TrataGotFocus(DataFimContrato, iAlterado)

End Sub

Private Sub DataFimContrato_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataFimContrato_Validate

    'Verifica se a Data de Emissao foi digitada
    If Len(Trim(DataFimContrato.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataFimContrato.Text)
    If lErro <> SUCESSO Then gError 129582

    Exit Sub

Erro_DataFimContrato_Validate:

    Cancel = True

    Select Case gErr

        'se houve erro de crítica, segura o foco
        Case 129582

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154995)

    End Select

    Exit Sub

End Sub

Private Sub DataInicioP_GotFocus()

     Call MaskEdBox_TrataGotFocus(DataInicioP, iAlterado)

End Sub

Private Sub DataInicioP_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataInicioP_Validate

    'Verifica se a Data de Emissao foi digitada
    If Len(Trim(DataInicioP.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataInicioP.Text)
    If lErro <> SUCESSO Then gError 131045

    Exit Sub

Erro_DataInicioP_Validate:

    Cancel = True

    Select Case gErr

        'se houve erro de crítica, segura o foco
        Case 131045

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154996)

    End Select

    Exit Sub

End Sub

Private Sub DataRenovContrato_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialCliente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialCliente_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Historico_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Historico_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iTamanho As Integer
Dim iCodigo As Integer
Dim iIndice As Integer
Dim sDescricao As Long
Dim objHistMovCta As New ClassHistMovCta

On Error GoTo Erro_Historico_Validate

    'Verifica o tamanho do texto em historico
    iTamanho = Len(Trim(Historico.Text))

    If iTamanho = 0 Then Exit Sub
    
    'Verifica se é maior que o tamanho maximo
    If iTamanho > 50 Then gError 129512
    
    'Verifica se o que foi digitado é numerico
    If IsNumeric(Trim(Historico.Text)) Then
    
        lErro = Valor_Inteiro_Critica(Trim(Historico.Text))
        If lErro <> SUCESSO Then gError 129513
           
        objHistMovCta.iCodigo = StrParaInt(Trim(Historico.Text))
                
        lErro = CF("HistMovCta_Le", objHistMovCta)
        If lErro <> SUCESSO And lErro <> 15011 Then gError 129514
    
        If lErro = 15011 Then gError 129515
                
        Historico.Text = objHistMovCta.sDescricao
            
    End If

    Exit Sub

Erro_Historico_Validate:

    Cancel = True

    Select Case gErr
    
        Case 129512
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_HISTORICOMOVCONTA", gErr)
         
        Case 129513
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INTEIRO", gErr, Historico.Text)
        
        Case 129514
        
        Case 129515
            Call Rotina_Erro(vbOKOnly, "ERRO_HISTMOVCTA_NAO_CADASTRADO", gErr, objHistMovCta.iCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154997)

    End Select

    Exit Sub

End Sub

Private Sub Atualiza_DataProxCobranca(ByVal dtDataInicio As Date, ByVal iPeriodicidade As Integer)
'Calcula a data da próxima cobrança em cima de:
    'Se o contrato foi faturado => data da última cobrança + a Periodicidade
    'Senão => data Inicial de cobrança

Dim dtDataProxima As Date
Dim dtDataUltCobranca As Date
Dim objContrato As New ClassContrato
Dim objItensDeContrato As New ClassItensDeContrato
Dim objNF As New ClassNFiscal
Dim objItemNF As New ClassItemNF
Dim bAchou As Boolean
Dim lErro As Long
Dim dtDataRefFim As Date
Dim dtDataRefFimAnt As Date
Dim dtDataRefIni As Date

On Error GoTo Erro_Atualiza_DataProxCobranca

    objContrato.sCodigo = Codigo.Text
    objContrato.iFilialEmpresa = giFilialEmpresa
    
    bAchou = True
    
    dtDataRefIni = StrParaDate(GridItens.TextMatrix(GridItens.Row, iGrid_DataRefIni_Col))
    dtDataRefFimAnt = StrParaDate(GridItens.TextMatrix(GridItens.Row, iGrid_DataRefFim_Col))

    'Busca o Contrato No Banco
    lErro = CF("Contrato_Le", objContrato)
    If lErro <> SUCESSO And lErro <> 129332 Then gError 129919
    If lErro = 129332 Then bAchou = False
    
    If objContrato.iTipo <> CONTRATOS_RECEBER And bAchou Then gError 132901
    
    'Se Achou o contrato busca o item de contrato
    If bAchou Then
        objItensDeContrato.iSeq = GridItens.TextMatrix(GridItens.Row, iGrid_Item_Col)
        objItensDeContrato.lNumIntContrato = objContrato.lNumIntDoc
    
        lErro = CF("ItensDeContrato_Le2", objItensDeContrato)
        If lErro <> SUCESSO And lErro <> 129266 Then gError 129920
        If lErro = 129266 Then bAchou = False
    
    End If

    'Se Achou o Itens de Contrato verifica se ele já foi faturado
    If bAchou Then
    
        objItemNF.objCobrItensContrato.lNumIntItensContrato = objItensDeContrato.lNumIntDoc
        objItemNF.objCobrItensContrato.lMedicao = 0
    
        'Obtém os itens fiscais gerados por último (Com maior data da última cobrança)
        lErro = CF("ItensDeContrato_Le_DadosFatura", objNF, objItemNF)
        If lErro <> SUCESSO And lErro <> 129904 And lErro <> 129907 And lErro <> 129908 Then gError 129921
        If lErro <> SUCESSO Then bAchou = False
        
    End If

    'Se o contrato é faturado a data da próxima cobranca é
    'A data da última cobrança + a Periodicidade
    'Senão é a data Inicial
    dtDataUltCobranca = objItemNF.objCobrItensContrato.dtDataUltCobranca
    
    Select Case iPeriodicidade
    
        Case PERIODICIDADE_LIVRE
            dtDataProxima = dtDataUltCobranca
            dtDataRefFim = dtDataRefFimAnt
        
        Case PERIODICIDADE_SEMANAL
            dtDataProxima = DateAdd("d", 7, dtDataUltCobranca)
            dtDataRefFim = DateAdd("d", 7, dtDataRefIni)
            dtDataRefFim = DateAdd("d", -1, dtDataRefFim)
        
        Case PERIODICIDADE_DECENDIAL
            dtDataProxima = DateAdd("d", 10, dtDataUltCobranca)
            dtDataRefFim = DateAdd("d", 10, dtDataRefIni)
            dtDataRefFim = DateAdd("d", -1, dtDataRefFim)
        
        Case PERIODICIDADE_QUINZENAL
            dtDataProxima = DateAdd("d", 15, dtDataUltCobranca)
            dtDataRefFim = DateAdd("d", 15, dtDataRefIni)
            dtDataRefFim = DateAdd("d", -1, dtDataRefFim)
        
        Case PERIODICIDADE_MENSAL
            dtDataProxima = DateAdd("m", 1, dtDataUltCobranca)
            dtDataRefFim = DateAdd("m", 1, dtDataRefIni)
            dtDataRefFim = DateAdd("d", -1, dtDataRefFim)
        
        Case PERIODICIDADE_BIMESTRAL
            dtDataProxima = DateAdd("m", 2, dtDataUltCobranca)
            dtDataRefFim = DateAdd("m", 2, dtDataRefIni)
            dtDataRefFim = DateAdd("d", -1, dtDataRefFim)
        
        Case PERIODICIDADE_TRIMESTRAL
            dtDataProxima = DateAdd("m", 3, dtDataUltCobranca)
            dtDataRefFim = DateAdd("m", 3, dtDataRefIni)
            dtDataRefFim = DateAdd("d", -1, dtDataRefFim)
        
        Case PERIODICIDADE_QUADRIMESTRAL
            dtDataProxima = DateAdd("m", 4, dtDataUltCobranca)
            dtDataRefFim = DateAdd("m", 4, dtDataRefIni)
            dtDataRefFim = DateAdd("d", -1, dtDataRefFim)
        
        Case PERIODICIDADE_SEMESTRAL
            dtDataProxima = DateAdd("m", 6, dtDataUltCobranca)
            dtDataRefFim = DateAdd("m", 6, dtDataRefIni)
            dtDataRefFim = DateAdd("d", -1, dtDataRefFim)
        
        Case PERIODICIDADE_ANUAL
            dtDataProxima = DateAdd("yyyy", 1, dtDataUltCobranca)
            dtDataRefFim = DateAdd("yyyy", 1, dtDataRefIni)
            dtDataRefFim = DateAdd("d", -1, dtDataRefFim)
        
        Case Else
            dtDataProxima = dtDataUltCobranca
            dtDataRefFim = dtDataRefFimAnt
        
    End Select
        
    If Not bAchou Then
        dtDataProxima = dtDataInicio
        dtDataRefFim = dtDataRefFimAnt
    End If
    
    GridItens.TextMatrix(GridItens.Row, iGrid_DataProx_Col) = Format(dtDataProxima, "dd/mm/yyyy")
    
    If dtDataRefFimAnt <> DATA_NULA Then
        GridItens.TextMatrix(GridItens.Row, iGrid_DataRefFim_Col) = Format(dtDataRefFim, "dd/mm/yyyy")
    End If
    
    Exit Sub

Erro_Atualiza_DataProxCobranca:

    Select Case gErr
    
        Case 129919 To 129921
            
        Case 132901
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_NAO_RECEBER", gErr, objContrato.sCodigo, objContrato.iFilialEmpresa)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154998)

    End Select
    
End Sub

Private Sub ContaContabil_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaFormatada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim sContaMascarada As String

On Error GoTo Erro_ContaContabil_Validate

    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica_Modulo", ContaContabil.Text, ContaContabil.ClipText, objPlanoConta, MODULO_TESOURARIA)
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 129508

    If lErro = SUCESSO Then

        sContaFormatada = objPlanoConta.sConta

        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then gError 129509

        ContaContabil.PromptInclude = False
        ContaContabil.Text = sContaMascarada
        ContaContabil.PromptInclude = True

    'Se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then

        'Critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
        lErro = CF("Conta_Critica", ContaContabil.Text, sContaFormatada, objPlanoConta, MODULO_TESOURARIA)
        If lErro <> SUCESSO And lErro <> 5700 Then gError 129510

        'Conta não cadastrada
        If lErro = 5700 Then gError 129511

    End If

    Exit Sub

Erro_ContaContabil_Validate:

    Cancel = True

    Select Case gErr

        Case 129508, 129510
    
        Case 129509
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            
        Case 129511
            'Pergunta se o usuário deseja cadastrar a conta
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabil.Text)
            'Se sim
            If vbMsgRes = vbYes Then
                'Preenche o parâmetro que será passado para a tela
                objPlanoConta.sConta = sContaFormatada
                'Chama a tela de plano de contas
                Call Chama_Tela("PlanoConta", objPlanoConta)
            End If
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154999)
    
    End Select

    Exit Sub
    
End Sub

Public Sub ContaContabilLabel_Click()
'BROWSE PLANO_CONTA :

Dim lErro As Long
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim iContaPreenchida As Integer
Dim sConta As String

On Error GoTo Erro_ContaContabilLabel_Click

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaContabil.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then gError 129507

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    Call Chama_Tela("PlanoContaTESLista", colSelecao, objPlanoConta, objEventoContaContabil)

    Exit Sub

Erro_ContaContabilLabel_Click:

    Select Case gErr

        Case 129507
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155000)

    End Select

    Exit Sub
    
End Sub

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
        If lErro <> SUCESSO Then gError 129506

        ContaContabil.Text = sContaEnxuta
        ContaContabil.PromptInclude = True
    End If

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case gErr

        Case 129506
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155001)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim objContrato As ClassContrato
Dim bCancel As Boolean

    Set objContrato = obj1

    Codigo.Text = objContrato.sCodigo

    Call Codigo_Validate(bCancel)

    Exit Sub
    
End Sub

Private Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Opcao_Click()

    'Se frame selecionado não for o atual
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        'Esconde o frame atual, mostra o novo
        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index

        Select Case iFrameAtual
        
            Case TAB_Identificacao
                Parent.HelpContextID = IDH_NOTA_FISCAL_FATURA_PAGAR_ID
                
            Case TAB_Itens
                Parent.HelpContextID = IDH_NOTA_FISCAL_FATURA_COBRANCA
                                   
        End Select

    End If
    
End Sub

Private Sub Periodicidade_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PeriodicidadeP_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PeriodicidadeP_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PeriodicidadeP_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_PeriodicidadeP_Validate

    'Verifica se a Condicaopagamento foi preenchida
    If Len(Trim(PeriodicidadeP.Text)) = 0 Then Exit Sub

    'Verifica se é uma Condicaopagamento selecionada
    If PeriodicidadeP.Text = PeriodicidadeP.List(PeriodicidadeP.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(PeriodicidadeP, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 129504

    If lErro = 6730 Then gError 129728
    If lErro = 6731 Then gError 129729

    Exit Sub

Erro_PeriodicidadeP_Validate:

    Cancel = True

    Select Case gErr
    
        Case 129728 To 129729
            Call Rotina_Erro(vbOKOnly, "ERRO_PERIODICIDADE_INVALIDA", gErr)
    
        Case 129504

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155002)

    End Select

    Exit Sub

End Sub

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


Private Sub ClienteLabel_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As New Collection
Dim sOrdenacao As String

On Error GoTo Erro_ClienteLabel_Click

    'Se é possível extrair o código do cliente do conteúdo do controle
    If LCodigo_Extrai(Cliente.Text) <> 0 Then

        'Guarda o código para ser passado para o browser
        objcliente.lCodigo = LCodigo_Extrai(Cliente.Text)

        sOrdenacao = "Codigo"

    'Senão, ou seja, se está digitado o nome do cliente
    Else
        
        'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
        objcliente.sNomeReduzido = Cliente.Text
        
        sOrdenacao = "Nome Reduzido + Código"
    
    End If
    
    'Chama a tela de consulta de cliente
    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoCliente, "", sOrdenacao)

    Exit Sub
    
Erro_ClienteLabel_Click:

    Select Case gErr
    
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155003)
    
    End Select
    
End Sub

Private Sub Cliente_Change()

    iAlterado = REGISTRO_ALTERADO
    iClienteAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim iCodFilial As Integer
Dim objTipoCliente As New ClassTipoCliente
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate

    If iClienteAlterado <> REGISTRO_ALTERADO Then Exit Sub

    'Se Cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código ou CPF ou CGC)
        lErro = TP_Cliente_Le(Cliente, objcliente, iCodFilial)
        If lErro <> SUCESSO Then gError 129422

        'Lê coleção de códigos, nomes de Filiais do Cliente
        lErro = CF("FiliaisClientes_Le_Cliente", objcliente, colCodigoNome)
        If lErro <> SUCESSO Then gError 129423

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", FilialCliente, colCodigoNome)

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", FilialCliente, iCodFilial)

        'Lê o tipo do cliente
        If objcliente.iTipo > 0 Then
            objTipoCliente.iCodigo = objcliente.iTipo
            lErro = CF("TipoCliente_Le", objTipoCliente)
            If lErro <> SUCESSO And lErro <> 19062 Then gError 42144
        End If

        If objcliente.iCondicaoPagto > 0 Then
            CondPagtoP.Text = objcliente.iCondicaoPagto
        ElseIf objTipoCliente.iCondicaoPagto > 0 Then
            CondPagtoP.Text = objTipoCliente.iCondicaoPagto
        End If
        
        Call CondPagtoP_Validate(bSGECancelDummy)
       
    'Se Cliente não está preenchido
    ElseIf Len(Trim(Cliente.Text)) = 0 Then

        'Limpa a Combo de Filiais
        FilialCliente.Clear

    End If
    
    iClienteAlterado = 0
    
    Exit Sub

Erro_Cliente_Validate:
        
    Cancel = True

    Select Case gErr
    
        Case 129422, 129423
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155004)

    End Select

    Exit Sub

End Sub

Private Sub FilialCliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objFilialCliente As New ClassFilialCliente
Dim iCodigo As Integer
Dim sNomeRed As String
Dim objClienteContatos As New ClassClienteContatos

On Error GoTo Erro_FilialCliente_Validate
        
    If Len(Trim(FilialCliente.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o item selecionado na ComboBox Filial
    If FilialCliente.Text = FilialCliente.List(FilialCliente.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(FilialCliente, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 129418

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        'Verifica se foi preenchido o Cliente
        If Len(Trim(Cliente.Text)) = 0 Then gError 129419

        'Lê o Cliente que está na tela
        sNomeRed = Trim(Cliente.Text)

        'Passa o Código da Filial que está na tela para o Obj
        objFilialCliente.iCodFilial = iCodigo

        'Lê Filial no BD a partir do NomeReduzido do Cliente e Código da Filial
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sNomeRed, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 129420

        'Se não existe a Filial
        If lErro = 17660 Then gError 129421

        'Encontrou Filial no BD, coloca no Text da Combo
        FilialCliente.Text = CStr(objFilialCliente.iCodFilial) & SEPARADOR & objFilialCliente.sNome

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 129505
    
    Exit Sub
    
Erro_FilialCliente_Validate:

    Select Case gErr

        Case 129418, 129420

        Case 129419
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 129421
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE1", FilialCliente.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If

        Case 129505
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA", gErr, FilialCliente.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155005)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objContrato As New ClassContrato

On Error GoTo Erro_Gravar_Registro

    If Len(Trim(Codigo.Text)) = 0 Then gError 129410
    If Len(Trim(Cliente.Text)) = 0 Then gError 129411
    If Codigo_Extrai(FilialCliente.Text) = 0 Then gError 129412
    If Len(Trim(Descricao.Text)) = 0 Then gError 129413
    If Len(Trim(DataIniContrato.ClipText)) = 0 Then gError 129414
    If Len(Trim(DataInicioP.ClipText)) = 0 Then gError 129415
    If Len(Trim(PeriodicidadeP.Text)) = 0 Then gError 129525
    If optNFe.Value = vbChecked And Len(Trim(Serie.Text)) = 0 Then gError 202764
    
    If Len(Trim(DataFimContrato.ClipText)) <> 0 Then
        If StrParaDate(DataIniContrato.Text) > StrParaDate(DataFimContrato.Text) Then gError 129535
    End If
    
    If Len(Trim(DataRenovContrato.ClipText)) <> 0 Then
        If StrParaDate(DataIniContrato.Text) > StrParaDate(DataRenovContrato.Text) Then gError 131053
        If Len(Trim(DataFimContrato.ClipText)) <> 0 And StrParaDate(DataRenovContrato.Text) > StrParaDate(DataFimContrato.Text) Then gError 131053
    End If
    
    If StrParaDate(DataIniContrato.Text) > StrParaDate(DataInicioP.Text) Then gError 131054
    
    lErro = Valida_Grid_Itens()
    If lErro <> SUCESSO Then gError 129575
       
    lErro = Move_Tela_Memoria(objContrato)
    If lErro <> SUCESSO Then gError 129416
    
    GL_objMDIForm.MousePointer = vbHourglass
       
    lErro = CF("Contrato_Grava", objContrato)
    If lErro <> SUCESSO Then gError 129417
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
        
        Case 129410
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CONTRATO_PREENCHIDO", gErr)

        Case 129411
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 129412
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_INFORMADA", gErr)

        Case 129413
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAOCONTRATO_NAO_INFORMADO", gErr)
        
        Case 129414
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICONTRATO_NAO_PREENCHIDA", gErr)

        Case 129415
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINIPADRAO_NAO_PREENCHIDA", gErr)

        Case 129525
            Call Rotina_Erro(vbOKOnly, "ERRO_PERIODICIDADEP_NAO_PREENCHIDA", gErr)

        Case 129416, 129417
        
        Case 129535
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MAIOR_DATAFINAL", gErr)
        
        Case 129575
        
        Case 131053
            Call Rotina_Erro(vbOKOnly, "ERRO_DATARENOVACAO_MAIOR", gErr)
        
        Case 131054
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICOBRANCAP_MENOR_DATAINICONTRATO", gErr)
        
        Case 132835
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_NAO_PREENCHIDA", gErr)
        
        Case 132836
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_NAO_PREENCHIDO", gErr)
            
        Case 202764
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155006)

    End Select

    Exit Function

End Function

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call CodigoLabel_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoProduto_Click
        ElseIf Me.ActiveControl Is Cliente Then
            Call ClienteLabel_Click
        ElseIf Me.ActiveControl Is CondPagtoP Then
            Call CondPagtoLabel_Click
        ElseIf Me.ActiveControl Is ContaContabil Then
            Call ContaContabilLabel_Click
        ElseIf Me.ActiveControl Is Ccl Then
            Call CclLabel_Click
        ElseIf Me.ActiveControl Is NaturezaOp Then
            Call NaturezaLabel_Click
        ElseIf Me.ActiveControl Is TipoTributacao Then
            Call LblTipoTrib_Click
        End If
          
    End If

End Sub

'#######################################################################
'INÍCIO DOS UPDOWN
'#######################################################################
Private Sub UpDownIniContrato_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownIniContrato_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataIniContrato, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 129483

    Exit Sub

Erro_UpDownIniContrato_DownClick:

    Select Case gErr

        Case 129483

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155007)

    End Select

    Exit Sub

End Sub

Private Sub UpDownIniContrato_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownIniContrato_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataIniContrato, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 129484

    Exit Sub

Erro_UpDownIniContrato_UpClick:

    Select Case gErr

        Case 129484

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155008)

    End Select

    Exit Sub

End Sub

Private Sub UpDownInicioP_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownInicioP_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataInicioP, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 129485

    Exit Sub

Erro_UpDownInicioP_DownClick:

    Select Case gErr

        Case 129485

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155009)

    End Select

    Exit Sub

End Sub

Private Sub UpDownInicioP_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownInicioP_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataInicioP, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 129486

    Exit Sub

Erro_UpDownInicioP_UpClick:

    Select Case gErr

        Case 129486

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155010)

    End Select

    Exit Sub

End Sub

Private Sub UpDownRenovContrato_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownRenovContrato_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataRenovContrato, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 129487

    Exit Sub

Erro_UpDownRenovContrato_DownClick:

    Select Case gErr

        Case 129487

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155011)

    End Select

    Exit Sub

End Sub

Private Sub UpDownRenovContrato_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownRenovContrato_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataRenovContrato, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 129488

    Exit Sub

Erro_UpDownRenovContrato_UpClick:

    Select Case gErr

        Case 129488

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155012)

    End Select

    Exit Sub

End Sub

Private Sub UpDownFimContrato_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownFimContrato_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataFimContrato, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 129489

    Exit Sub

Erro_UpDownFimContrato_DownClick:

    Select Case gErr

        Case 129489

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155013)

    End Select

    Exit Sub

End Sub

Private Sub UpDownFimContrato_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownFimContrato_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataFimContrato, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 131050

    Exit Sub

Erro_UpDownFimContrato_UpClick:

    Select Case gErr

        Case 131050

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155014)

    End Select

    Exit Sub

End Sub
'#######################################################################
'FIM DOS UPDOWN
'#######################################################################

'#######################################################################
'INÍCIO DAS FUNÇÕES DE SAÍDA DE CÉLULA
'#######################################################################
Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
                
        If objGridInt.objGrid Is GridItens Then
        
            Select Case GridItens.Col
    
                Case iGrid_Cobrar_Col
    
                    lErro = Saida_Celula_Cobrar(objGridInt)
                    If lErro <> SUCESSO Then gError 129424
                
                Case iGrid_CondPagto_Col
    
                    lErro = Saida_Celula_CondPagto(objGridInt)
                    If lErro <> SUCESSO Then gError 129425
        
                Case iGrid_DataInicio_Col
    
                    lErro = Saida_Celula_DataInicio(objGridInt)
                    If lErro <> SUCESSO Then gError 129426
        
                Case iGrid_DataProx_Col
    
                    lErro = Saida_Celula_DataProx(objGridInt)
                    If lErro <> SUCESSO Then gError 131055
        
                Case iGrid_Descricao_Col
    
                    lErro = Saida_Celula_Descricao(objGridInt)
                    If lErro <> SUCESSO Then gError 129427
            
                Case iGrid_Medicao_Col
    
                    lErro = Saida_Celula_Medicao(objGridInt)
                    If lErro <> SUCESSO Then gError 129428
                
                Case iGrid_Item_Col
    
                    lErro = Saida_Celula_Item(objGridInt)
                    If lErro <> SUCESSO Then gError 129840
            
                Case iGrid_Obs_Col
    
                    lErro = Saida_Celula_Obs(objGridInt)
                    If lErro <> SUCESSO Then gError 129429
             
                Case iGrid_Periodicidade_Col
    
                    lErro = Saida_Celula_Periodicidade(objGridInt)
                    If lErro <> SUCESSO Then gError 129430
             
                Case iGrid_Produto_Col
    
                    lErro = Saida_Celula_Produto(objGridInt)
                    If lErro <> SUCESSO Then gError 129431
                         
                Case iGrid_Quantidade_Col
    
                    lErro = Saida_Celula_Quantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 129433
             
                Case iGrid_UM_Col
    
                    lErro = Saida_Celula_UM(objGridInt)
                    If lErro <> SUCESSO Then gError 129434
                
                Case iGrid_Valor_Col
    
                    lErro = Saida_Celula_Valor(objGridInt)
                    If lErro <> SUCESSO Then gError 129435
                
                Case iGrid_DataRefIni_Col
    
                    lErro = Saida_Celula_DataRefIni(objGridInt)
                    If lErro <> SUCESSO Then gError 136065
                
                Case iGrid_DataRefFim_Col
    
                    lErro = Saida_Celula_DataRefFim(objGridInt)
                    If lErro <> SUCESSO Then gError 136066
      
                Case iGrid_QtdeParcelas_Col
    
                    lErro = Saida_Celula_QtdeParcelas(objGridInt)
                    If lErro <> SUCESSO Then gError 129433
             
                Case iGrid_UltParcCobrada_Col
    
                    lErro = Saida_Celula_UltParcCobrada(objGridInt)
                    If lErro <> SUCESSO Then gError 129433
             
             End Select
                
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 129437
        
        iAlterado = REGISTRO_ALTERADO

    End If
       
    Saida_Celula = SUCESSO
    
    Exit Function

Erro_Saida_Celula:
    
    Saida_Celula = gErr
    
    Select Case gErr

        Case 129424 To 129436
        
        Case 129840, 131055, 136065, 136066

        Case 129437
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155015)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    If Len(Trim(Quantidade.ClipText)) <> 0 Then

        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 129440

        Quantidade.Text = Formata_Estoque(Quantidade.Text)

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129444

    Call Soma_Valor

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 129440, 129444
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 155016)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Valor(objGridInt As AdmGrid) As Long
'Faz a crítica da celula Valor do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridInt.objControle = Valor

    'Verifica se valor está preenchido
    If Len(Trim(Valor.Text)) > 0 Then
    
        'Critica se valor é positivo
        lErro = Valor_Positivo_Critica(Valor.Text)
        If lErro <> SUCESSO Then gError 129452

        Valor.Text = Format(Valor.Text, Valor.Format)
        
        'Acrescenta uma linha no Grid se for o caso
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
              
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129453
    
    Call Soma_Valor
    
    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = gErr

    Select Case gErr

        Case 129452, 129453
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155017)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Medicao(objGridInt As AdmGrid) As Long
'Faz a crítica da celula Medicao do grid que está deixando de ser a corrente
'Inserido por Wagner

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Medicao

    Set objGridInt.objControle = Medicao

    'Verifica se valor está preenchido
    If Medicao.Value <> 0 Then
           
        'Acrescenta uma linha no Grid se for o caso
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 131049

    Saida_Celula_Medicao = SUCESSO

    Exit Function

Erro_Saida_Celula_Medicao:

    Saida_Celula_Medicao = gErr

    Select Case gErr

        Case 131049
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155018)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Cobrar(objGridInt As AdmGrid) As Long
'Faz a crítica da celula Cobrar do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Cobrar

    Set objGridInt.objControle = Cobrar

    'Verifica se valor está preenchido
    If Cobrar.Value <> 0 Then
           
        'Acrescenta uma linha no Grid se for o caso
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129454

    Saida_Celula_Cobrar = SUCESSO

    Exit Function

Erro_Saida_Celula_Cobrar:

    Saida_Celula_Cobrar = gErr

    Select Case gErr

        Case 129454
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155019)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataInicio(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DataInicio

    Set objGridInt.objControle = DataInicio

    'verifica se a data está preenchida
    If Len(Trim(DataInicio.ClipText)) > 0 Then

        'verifica se a data é válida
        lErro = Data_Critica(DataInicio.Text)
        If lErro <> SUCESSO Then gError 129450
                
        Call Atualiza_DataProxCobranca(StrParaDate(DataInicio.Text), Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_Periodicidade_Col)))
                
        'ALTERAÇÃO DE LINHAS EXISTENTES
        If (GridItens.Row - GridItens.FixedRows) = objGridItens.iLinhasExistentes Then
            objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129451

    Saida_Celula_DataInicio = SUCESSO

    Exit Function

Erro_Saida_Celula_DataInicio:

    Saida_Celula_DataInicio = gErr

    Select Case gErr

        Case 129450, 129451
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155020)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataProx(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DataProx

    Set objGridInt.objControle = DataProx

    'verifica se a data está preenchida
    If Len(Trim(DataProx.ClipText)) > 0 Then

        'verifica se a data é válida
        lErro = Data_Critica(DataProx.Text)
        If lErro <> SUCESSO Then gError 131048
        
        'ALTERAÇÃO DE LINHAS EXISTENTES
        If (GridItens.Row - GridItens.FixedRows) = objGridItens.iLinhasExistentes Then
            objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129449

    Saida_Celula_DataProx = SUCESSO

    Exit Function

Erro_Saida_Celula_DataProx:

    Saida_Celula_DataProx = gErr

    Select Case gErr

        Case 131048, 129449
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155021)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Obs(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Obs

    Set objGridInt.objControle = Obs

    GridItens.TextMatrix(GridItens.Row, iGrid_Obs_Col) = Obs.Text

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129447

    Saida_Celula_Obs = SUCESSO

    Exit Function

Erro_Saida_Celula_Obs:

    Saida_Celula_Obs = gErr

    Select Case gErr

        Case 129447
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 155022)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Descricao(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Descricao

    Set objGridInt.objControle = ItemDescricao

    GridItens.TextMatrix(GridItens.Row, iGrid_Descricao_Col) = ItemDescricao.Text

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129446

    Saida_Celula_Descricao = SUCESSO

    Exit Function

Erro_Saida_Celula_Descricao:

    Saida_Celula_Descricao = gErr

    Select Case gErr

        Case 129446
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 155023)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UM(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_UM

    Set objGridInt.objControle = UM

    GridItens.TextMatrix(GridItens.Row, iGrid_UM_Col) = UM.Text

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129445

    Saida_Celula_UM = SUCESSO

    Exit Function

Erro_Saida_Celula_UM:

    Saida_Celula_UM = gErr

    Select Case gErr

        Case 129445
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 155024)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_CondPagto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim iCodigo As Integer

On Error GoTo Erro_Saida_Celula_CondPagto

    Set objGridInt.objControle = CondPagto

    'Verifica se a Condicaopagamento foi preenchida
    If Len(Trim(CondPagto.Text)) <> 0 Then

        'Verifica se é uma Condicaopagamento selecionada
        If CondPagto.Text <> CondPagto.List(CondPagto.ListIndex) Then
    
            'Tenta selecionar na combo
            lErro = Combo_Seleciona_Grid(CondPagto, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 129473
            
            'Nao existe o ítem com o CÓDIGO na List da ComboBox
            If lErro = 6730 Then
        
                objCondicaoPagto.iCodigo = iCodigo
        
                'Tenta ler CondicaoPagto com esse código no BD
                lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
                If lErro <> SUCESSO And lErro <> 19205 Then gError 129724
                
                'Não encontrou CondicaoPagto no BD
                If lErro <> SUCESSO Then gError 129725
        
                'Encontrou CondicaoPagto no BD e não é de Recebimento
                If objCondicaoPagto.iEmRecebimento = 0 Then gError 129726
        
                'Coloca no Text da Combo
                CondPagto.Text = CondPagto_Traz(objCondicaoPagto)
        
            End If
        
            'Não existe o ítem com a STRING na List da ComboBox
            If lErro = 6731 Then gError 129727
            
            GridItens.TextMatrix(GridItens.Row, iGrid_CondPagto_Col) = CondPagto.Text
        
        End If

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129474

    Saida_Celula_CondPagto = SUCESSO

    Exit Function

Erro_Saida_Celula_CondPagto:

    Saida_Celula_CondPagto = gErr

    Select Case gErr
        
        Case 129726
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_RECEBIMENTO", gErr, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 129725

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAO_PAGAMENTO")

            If vbMsgRes = vbYes Then
                'Chama a tela de CondicaoPagto
                Call Chama_Tela("CondicoesPagto", objCondicaoPagto)

            End If
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 129727
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", gErr, CondPagto.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 129473, 129474, 129724
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 155025)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Periodicidade(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Saida_Celula_Periodicidade

    Set objGridInt.objControle = Periodicidade

    'Verifica se a Condicaopagamento foi preenchida
    If Len(Trim(Periodicidade.Text)) <> 0 Then
    
        'Verifica se é uma Condicaopagamento selecionada
        If Periodicidade.Text <> Periodicidade.List(Periodicidade.ListIndex) Then
    
            'Tenta selecionar na combo
            lErro = Combo_Seleciona(Periodicidade, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 129475
            
            If lErro = 6730 Then gError 129730
            If lErro = 6731 Then gError 129731
            
            GridItens.TextMatrix(GridItens.Row, iGrid_Periodicidade_Col) = Periodicidade.Text

        End If
        
        Call Atualiza_DataProxCobranca(StrParaDate(GridItens.TextMatrix(GridItens.Row, iGrid_DataInicio_Col)), Codigo_Extrai(Periodicidade.Text))
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129476
    
    Saida_Celula_Periodicidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Periodicidade:

    Saida_Celula_Periodicidade = gErr

    Select Case gErr

        Case 129475, 129476
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 129730 To 129731
            Call Rotina_Erro(vbOKOnly, "ERRO_PERIODICIDADE_INVALIDA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 155026)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Item(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Item

    Set objGridInt.objControle = Item
    
    For iIndice = 1 To objGridItens.iLinhasExistentes
        If iIndice <> GridItens.Row Then
            If GridItens.TextMatrix(iIndice, iGrid_Item_Col) = Item.Text Then gError 129839
        End If
    Next
        
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 131047

    Saida_Celula_Item = SUCESSO

    Exit Function

Erro_Saida_Celula_Item:

    Saida_Celula_Item = gErr

    Select Case gErr
        
        Case 129839
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_JA_EXISTENTE", gErr, Item.Text, iIndice)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 131047
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 155027)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult
Dim iIndice As Integer
Dim sProduto As String
Dim iSeq As Integer

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    'Verifica se o produto existe e foi preenchido
    lErro = CF("Produto_Critica", Produto.Text, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 25041 Then gError 129441

    'se o produto não estiver cadastrado
    If lErro = 25041 Then gError 129442
            
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
        lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then gError 129550
    
        Produto.PromptInclude = False
        Produto.Text = sProduto
        Produto.PromptInclude = True
   
        lErro = ProdutoLinha_Preenche(objProduto)
        If lErro <> SUCESSO Then gError 132500
            
    Else
        
        GridItens.TextMatrix(GridItens.Row, iGrid_UM_Col) = ""
        GridItens.TextMatrix(GridItens.Row, iGrid_Descricao_Col) = ""
        GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = ""
    
    End If
           
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129443

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 129441, 129443, 132500
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 129442
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)

            If vbMsg = vbYes Then
                objProduto.sCodigo = Produto.Text
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
                 
        Case 129550
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 155028)

    End Select

    Exit Function

End Function
'#######################################################################
'FIM DAS FUNÇÕES DE SAÍDA DE CÉLULA
'#######################################################################

'#######################################################################
'INÍCIO DO SCRIPT DO GRID
'#######################################################################
Public Sub GridItens_GotFocus()

    Call Grid_Recebe_Foco(objGridItens)

End Sub

Public Sub GridItens_EnterCell()

    Call Grid_Entrada_Celula(objGridItens, iAlterado)

End Sub

Public Sub GridItens_LeaveCell()

    Call Saida_Celula(objGridItens)

End Sub

Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Public Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinha As Integer
Dim iLinhaAtual As Integer
Dim iLinhasExistentesAnterior As Integer

    Call Grid_Trata_Tecla1(KeyCode, objGridItens)

    iLinhaAtual = GridItens.Row
    
    iLinhasExistentesAnterior = objGridItens.iLinhasExistentes
    If KeyCode = vbKeyDelete Then

        Call Soma_Valor
        
    End If

End Sub

Public Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Public Sub GridItens_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridItens)

End Sub

Public Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGridItens)

End Sub

Public Sub GridItens_Scroll()

    Call Grid_Scroll(objGridItens)

End Sub
'#######################################################################
'FIM DO SCRIPT DO GRID
'#######################################################################

'#######################################################################
'INÍCIO DO SCRIPT PARA CAMPOS DO GRID
'#######################################################################
Public Sub Obs_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Obs_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Obs_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Obs_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Obs
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Quantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Quantidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub UM_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UM_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub UM_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub UM_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = UM
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Item_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Item_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Item_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Item_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Item
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Valor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Valor_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Valor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Valor
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub CondPagto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CondPagto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub CondPagto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub CondPagto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = CondPagto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Periodicidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Periodicidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Periodicidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Periodicidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Periodicidade
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Medicao_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Medicao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Medicao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Medicao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Medicao
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Cobrar_Click()

    iAlterado = REGISTRO_ALTERADO
           
    Call Soma_Valor
    
End Sub

Public Sub Cobrar_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Cobrar_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Cobrar_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Cobrar
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True


End Sub

Public Sub ItemDescricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ItemDescricao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub ItemDescricao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub ItemDescricao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = ItemDescricao
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DataInicio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataInicio_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub DataInicio_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub DataInicio_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DataInicio
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DataProx_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataProx_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub DataProx_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub DataProx_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DataProx
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Produto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Produto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub
'#######################################################################
'FIM DO SCRIPT PARA CAMPOS DO GRID
'#######################################################################

'#######################################################################
'INÍCIO DO SCRIPT DO BROWSE
'#######################################################################
Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente
Dim bCancel As Boolean

    Set objcliente = obj1

    'Preenche o Cliente com o Cliente selecionado
    Cliente.Text = objcliente.sNomeReduzido

    'Dispara o Validate de Cliente
    Call Cliente_Validate(bCancel)

    Exit Sub

End Sub

Private Sub objEventoCondPagto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCondicaoPagto As ClassCondicaoPagto
Dim vbMsgRes As VbMsgBoxResult
Dim dPercAcresFin As Double

On Error GoTo Erro_objEventoCondPagto_evSelecao

    Set objCondicaoPagto = obj1

    'Preenche campo CondicaoPagamento
    CondPagtoP.Text = CondPagto_Traz(objCondicaoPagto)

    Exit Sub

Erro_objEventoCondPagto_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155029)

     End Select

     Exit Sub
End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String
Dim objTipoDeProduto As New ClassTipoDeProduto
Dim objCTBConfig As New ClassCTBConfig
Dim objItemMovEst As New ClassItemMovEstoque
    
On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1
    
    If GridItens.Row = 0 Then gError 129465
    
    lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 129466

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then gError 129467

    sProdutoMascarado = String(STRING_PRODUTO, 0)

    lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 129468

    'Lê os demais atributos do Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 129469

    If lErro = 28030 Then gError 129470
    
    Produto.PromptInclude = False
    Produto.Text = sProdutoMascarado
    Produto.PromptInclude = True

    If Not (Me.ActiveControl Is Produto) Then
       
        'Preenche a Linha do Grid
        lErro = ProdutoLinha_Preenche(objProduto)
        If lErro <> SUCESSO Then gError 129471
        
        GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text
    
    End If
    
    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 129466, 129469, 129471

        Case 129468
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
    
        Case 129470
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
                
        Case 129465
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
   
        Case 129467
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_PREENCHIDO_LINHA_GRID", gErr, GridItens.Row)
   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 155030)

    End Select

    Exit Sub

End Sub
'#######################################################################
'FIM DO SCRIPT DO BROWSE
'#######################################################################

'#######################################################################
'INÍCIO SCRIPT DE BOTÕES SUPERIORES
'#######################################################################
Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    
    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 129408

    Call Limpa_Tela_Contrato

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 129408

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 155031)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 129409

    Call Limpa_Tela_Contrato
    
    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 129409

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155032)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim objContrato As New ClassContrato
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'Verifica se o código foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 129477

    objContrato.sCodigo = Codigo.Text
    objContrato.iFilialEmpresa = giFilialEmpresa
    
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_CONTRATO", objContrato.sCodigo)

    If vbMsgRes = vbYes Then

        'Exclui a producao
        lErro = CF("Contrato_Exclui", objContrato)
        If lErro <> SUCESSO Then gError 129478

        Call Limpa_Tela_Contrato
    End If
    
    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 129477
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 129478
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155033)
    
    End Select
    
    Exit Sub
End Sub
'#######################################################################
'FIM SCRIPT DE BOTÕES SUPERIORES
'#######################################################################

'#######################################################################
'INÍCIO DAS ROTINAS RELACIONADAS AS SETAS DO SISTEMA
'#######################################################################
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objContrato As New ClassContrato

On Error GoTo Erro_Tela_Preenche

    objContrato.sCodigo = colCampoValor.Item("Codigo").vValor
    objContrato.iFilialEmpresa = giFilialEmpresa

    If Len(Trim(objContrato.sCodigo)) > 0 Then

        lErro = Traz_Contrato_Tela(objContrato)
        If lErro <> SUCESSO Then gError 129480

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 129480

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155034)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objContrato As New ClassContrato

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Contratos"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objContrato)
    If lErro <> SUCESSO Then gError 129481

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objContrato.sCodigo, STRING_CONTRATO_CODIGO, "Codigo"
      
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "Tipo", OP_IGUAL, CONTRATOS_RECEBER
      
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 129481

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155035)

    End Select

    Exit Sub

End Sub
'#######################################################################
'FIM DAS ROTINAS RELACIONADAS AS SETAS DO SISTEMA
'#######################################################################

'#######################################################################
'INÍCIO DO SCRIPT PARA MODO DE EDICAO
'#######################################################################
Private Sub CodigoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CodigoLabel, Source, X, Y)
End Sub

Private Sub CodigoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CodigoLabel, Button, Shift, X, Y)
End Sub

Private Sub Codigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Codigo, Source, X, Y)
End Sub

Private Sub Ativo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Ativo, Source, X, Y)
End Sub

Private Sub Ativo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Ativo, Button, Shift, X, Y)
End Sub

Private Sub DescricaoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescricaoLabel, Source, X, Y)
End Sub

Private Sub DescricaoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescricaoLabel, Button, Shift, X, Y)
End Sub

Private Sub Descricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Descricao, Source, X, Y)
End Sub

Private Sub Descricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Descricao, Button, Shift, X, Y)
End Sub

Private Sub ClienteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ClienteLabel, Source, X, Y)
End Sub

Private Sub ClienteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ClienteLabel, Button, Shift, X, Y)
End Sub

Private Sub Cliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Cliente, Source, X, Y)
End Sub

Private Sub FilialClienteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialClienteLabel, Source, X, Y)
End Sub

Private Sub FilialClienteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialClienteLabel, Button, Shift, X, Y)
End Sub

Private Sub FilialCliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialCliente, Source, X, Y)
End Sub

Private Sub ObservacaoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ObservacaoLabel, Source, X, Y)
End Sub

Private Sub ObservacaoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ObservacaoLabel, Button, Shift, X, Y)
End Sub

Private Sub Observacao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Observacao, Source, X, Y)
End Sub

Private Sub Observacao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Observacao, Button, Shift, X, Y)
End Sub

Private Sub DataIniContratoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataIniContratoLabel, Source, X, Y)
End Sub

Private Sub DataIniContratoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataIniContratoLabel, Button, Shift, X, Y)
End Sub

Private Sub DataIniContrato_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataIniContrato, Source, X, Y)
End Sub

Private Sub DataFimContratoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataFimContratoLabel, Source, X, Y)
End Sub

Private Sub DataFimContratoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataFimContratoLabel, Button, Shift, X, Y)
End Sub

Private Sub DataFimContrato_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataFimContrato, Source, X, Y)
End Sub


Private Sub DataRenovContratoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataRenovContratoLabel, Source, X, Y)
End Sub

Private Sub DataRenovContratoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataRenovContratoLabel, Button, Shift, X, Y)
End Sub

Private Sub DataRenovContrato_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataRenovContrato, Source, X, Y)
End Sub

Private Sub HistoricoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(HistoricoLabel, Source, X, Y)
End Sub

Private Sub HistoricoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(HistoricoLabel, Button, Shift, X, Y)
End Sub

Private Sub Historico_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Historico, Source, X, Y)
End Sub

Private Sub ContaContabilLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContaContabilLabel, Source, X, Y)
End Sub

Private Sub ContaContabilLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContaContabilLabel, Button, Shift, X, Y)
End Sub

Private Sub ContaContabil_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContaContabil, Source, X, Y)
End Sub

Private Sub CclLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CclLabel, Source, X, Y)
End Sub

Private Sub CclLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CclLabel, Button, Shift, X, Y)
End Sub

Private Sub Ccl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Ccl, Source, X, Y)
End Sub
'#######################################################################
'FIM DO SCRIPT PARA MODO DE EDICAO
'#######################################################################

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim lErro As Long
Dim objClasseUM As New ClassClasseUM
Dim colSiglas As New Collection
Dim objUM As ClassUnidadeDeMedida
Dim sUM As String
Dim iTipo As Integer
Dim sUnidadeMed As String
Dim iIndice As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    'Formata o produto do grid de itens
    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 129528

    Select Case objControl.Name

        Case Produto.Name
            'Se o produto estiver preenchido desabilita
            If iProdutoPreenchido <> PRODUTO_VAZIO Then
                Produto.Enabled = False
            Else
                Produto.Enabled = True
            End If

        Case UM.Name
                    
            'guarda a um go grid nessa coluna
            sUM = GridItens.TextMatrix(iLinha, iGrid_UM_Col)

            UM.Enabled = True

            'Guardo o valor da Unidade de Medida da Linha
            sUnidadeMed = UM.Text

            UM.Clear

            If iProdutoPreenchido <> PRODUTO_VAZIO Then

                objProduto.sCodigo = sProdutoFormatado
                'Lê o produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 129529

                If lErro = 28030 Then gError 129530

                objClasseUM.iClasse = objProduto.iClasseUM
                'Lê as UMs do produto
                lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                If lErro <> SUCESSO Then gError 129531
                'Carrega a combo de UMs
                For Each objUM In colSiglas
                    UM.AddItem objUM.sSigla
                Next

                If UM.ListCount <> 0 Then

                    For iIndice = 0 To UM.ListCount - 1

                        If UM.List(iIndice) = sUnidadeMed Then
                            UM.ListIndex = iIndice
                            Exit For
                        End If
                    Next
                End If

            Else
                UM.Enabled = False
            End If

        Case Quantidade.Name
                
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Or left(GridItens.TextMatrix(iLinha, 0), 1) = "#" Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
                     
        Case ItemDescricao.Name
                
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case Medicao.Name
                
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case Cobrar.Name
                
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
                        
            End If

        Case Valor.Name
                
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case Item.Name
        
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Or left(GridItens.TextMatrix(iLinha, 0), 1) = "#" Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case DataInicio.Name
                
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case DataProx.Name
                
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case CondPagto.Name
                
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case Obs.Name
                
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
           
        Case Periodicidade.Name
                
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
    
        Case Else
            
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 129528, 129529, 129531

        Case 129530
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155036)

    End Select

    Exit Sub

End Sub

'#########################################
'Inserido por Wagner
Private Sub NaturezaLabel_Click()

Dim objNaturezaOp As New ClassNaturezaOp
Dim colSelecao As New Collection
Dim dtDataRef As Date

    'Se NaturezaOP estiver preenchida coloca no Obj
    objNaturezaOp.sCodigo = NaturezaOp.Text
       
    'selecao p/obter apenas as nat de saida
    colSelecao.Add NATUREZA_SAIDA_COD_INICIAL
    colSelecao.Add NATUREZA_SAIDA_COD_FINAL

    'Chama a Tela de browse de NaturezaOp
    Call Chama_Tela("NaturezaOpLista", colSelecao, objNaturezaOp, objEventoNaturezaOp, "{fn LENGTH(Codigo) } = 4")

End Sub

Private Sub NaturezaOp_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NaturezaOp_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objNaturezaOp As New ClassNaturezaOp
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_NaturezaOp_Validate
    
    'Verifica se a NaturezaOP foi informada
    If Len(Trim(NaturezaOp.ClipText)) <> 0 Then

        objNaturezaOp.sCodigo = Trim(NaturezaOp.Text)
    
        If objNaturezaOp.sCodigo < NATUREZA_SAIDA_COD_INICIAL Or objNaturezaOp.sCodigo > NATUREZA_SAIDA_COD_FINAL Then gError 94495
        
        'Lê a NaturezaOp
        lErro = CF("NaturezaOperacao_Le", objNaturezaOp)
        If lErro <> SUCESSO And lErro <> 17958 Then gError 94493
    
        'Se não existir --> Erro
        If lErro = 17958 Then gError 94494
        
        DescNatOp.Caption = objNaturezaOp.sDescricao
        
    Else
        DescNatOp.Caption = ""
        
    End If
        
    Exit Sub

Erro_NaturezaOp_Validate:

    Cancel = True

    Select Case gErr

        Case 94493

        Case 94494

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_NATUREZA_OPERACAO", NaturezaOp.Text)
            If vbMsgRes = vbYes Then
                Call Chama_Tela("NaturezaOperacao", objNaturezaOp)
            Else
            End If

        Case 94495
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_SAIDA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155037)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNaturezaOp_evSelecao(obj1 As Object)

Dim objNaturezaOp As New ClassNaturezaOp

    Set objNaturezaOp = obj1
        
    'Preenche a natureza de Opereração do frame principal
    NaturezaOp.Text = objNaturezaOp.sCodigo
    Call NaturezaOp_Validate(bSGECancelDummy)

    Me.Show

End Sub

Public Sub LblTipoTrib_Click()

Dim colSelecao As New Collection
Dim objTipoTrib As New ClassTipoDeTributacaoMovto

    'apenas tipos de saida
    colSelecao.Add "0"
    colSelecao.Add "0"
    
    Call Chama_Tela("TiposDeTribMovtoLista", colSelecao, objTipoTrib, objEventoTiposDeTributacao)

End Sub

Private Sub objEventoTiposDeTributacao_evSelecao(obj1 As Object)

Dim objTipoTrib As ClassTipoDeTributacaoMovto

    Set objTipoTrib = obj1

    TipoTributacao.Text = objTipoTrib.iTipo
    Call TipoTributacao_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Public Sub TipoTributacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoTributacao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTipoDeTributacao As New ClassTipoDeTributacaoMovto
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_TipoTributacao_Validate

    If Len(Trim(TipoTributacao.Text)) = 0 Then
        'Limpa o campo da descrição
        DescTipoTrib.Caption = ""
    End If

    objTipoDeTributacao.iTipo = StrParaInt(TipoTributacao.Text)

    If objTipoDeTributacao.iTipo <> 0 Then
        lErro = CF("TipoTributacao_Le", objTipoDeTributacao)
        If lErro <> SUCESSO And lErro <> 27259 Then gError 27663

        'Se não encontrou o Tipo da Tributação --> erro
        If lErro = 27259 Then gError 27664
    End If

    DescTipoTrib.Caption = objTipoDeTributacao.sDescricao

    Exit Sub

Erro_TipoTributacao_Validate:

    Cancel = True

    Select Case gErr

        Case 27663

        Case 27664
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TIPOTRIBUTACAO", TipoTributacao.Text)

            If vbMsgRes = vbYes Then

                Call Chama_Tela("TipoDeTributacao", objTipoDeTributacao)

            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155038)

    End Select

    Exit Sub

End Sub

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

Public Sub DataRefIni_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataRefIni_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub DataRefIni_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub DataRefIni_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DataRefIni
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DataRefFim_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataRefFim_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub DataRefFim_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub DataRefFim_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DataRefFim
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_DataRefIni(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DataRefIni

    Set objGridInt.objControle = DataRefIni

    'verifica se a data está preenchida
    If Len(Trim(DataRefIni.ClipText)) > 0 Then

        'verifica se a data é válida
        lErro = Data_Critica(DataRefIni.Text)
        If lErro <> SUCESSO Then gError 136056
                
        'ALTERAÇÃO DE LINHAS EXISTENTES
        If (GridItens.Row - GridItens.FixedRows) = objGridItens.iLinhasExistentes Then
            objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 136057

    Saida_Celula_DataRefIni = SUCESSO

    Exit Function

Erro_Saida_Celula_DataRefIni:

    Saida_Celula_DataRefIni = gErr

    Select Case gErr

        Case 136056, 136057
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155039)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataRefFim(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DataRefFim

    Set objGridInt.objControle = DataRefFim

    'verifica se a data está preenchida
    If Len(Trim(DataRefFim.ClipText)) > 0 Then

        'verifica se a data é válida
        lErro = Data_Critica(DataRefFim.Text)
        If lErro <> SUCESSO Then gError 136058
                
        'ALTERAÇÃO DE LINHAS EXISTENTES
        If (GridItens.Row - GridItens.FixedRows) = objGridItens.iLinhasExistentes Then
            objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 136059

    Saida_Celula_DataRefFim = SUCESSO

    Exit Function

Erro_Saida_Celula_DataRefFim:

    Saida_Celula_DataRefFim = gErr

    Select Case gErr

        Case 136058, 136059
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155040)

    End Select

    Exit Function

End Function
'########################################


Private Sub BotaoMedicoes_Click()

Dim objMedicaoContrato As New ClassMedicaoContrato
Dim colSelecao As New Collection
    
    If Len(Trim(Codigo.Text)) <> 0 Then
    
        colSelecao.Add Codigo.Text
        colSelecao.Add giFilialEmpresa
        
        Call Chama_Tela("MedicaoContratosLista", colSelecao, objMedicaoContrato, Nothing, "Contrato = ? AND FilialEmpresa = ? ")
    
    End If

End Sub

Private Sub BotaoFaturamentos_Click()

Dim objNF As New ClassNFiscal
Dim colSelecao As New Collection
   
    If Len(Trim(Codigo.Text)) <> 0 Then
    
        colSelecao.Add Codigo.Text
        colSelecao.Add giFilialEmpresa
        
        Call Chama_Tela("ContratosFaturamentoLista", colSelecao, objNF, Nothing, "Contrato = ? AND FilialEmpresa = ?")
    
    End If

End Sub

Private Function Carrega_Serie() As Long
'Carrega a combo de Séries com as séries lidas do BD

Dim lErro As Long
Dim objSerie As ClassSerie
Dim colSerie As New colSerie
Dim sSerieAnt As String

On Error GoTo Erro_Carrega_Serie

    'Lê as séries
    lErro = CF("Series_Le", colSerie)
    If lErro <> SUCESSO Then gError 42121

    'Carrega na combo
    For Each objSerie In colSerie
        If UCase(sSerieAnt) <> UCase(Desconverte_Serie_Eletronica(objSerie.sSerie)) Then Serie.AddItem Desconverte_Serie_Eletronica(objSerie.sSerie)
        sSerieAnt = Desconverte_Serie_Eletronica(objSerie.sSerie)
    Next

    Carrega_Serie = SUCESSO

    Exit Function

Erro_Carrega_Serie:

    Carrega_Serie = gErr

    Select Case gErr

        Case 42121

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157166)

    End Select

    Exit Function

End Function

Private Sub Serie_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objSerie As New ClassSerie
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Serie_Validate
   
    'Verifica se foi preenchida
    If Len(Trim(Serie.Text)) = 0 Then Exit Sub
    
    'Verifica se foi selecionada
    If Serie.ListIndex = -1 Then

        'Tenta selecionar a serie
        lErro = Combo_Item_Igual(Serie)
        If lErro <> SUCESSO And lErro <> 12253 Then gError 42158
        'Se não está na combo
        If lErro <> SUCESSO Then
        
            objSerie.sSerie = Converte_Serie_Eletronica(Serie.Text, optNFe.Value)
            'Busca a série no BD
            lErro = CF("Serie_Le", objSerie)
            If lErro <> SUCESSO And lErro <> 22202 Then gError 51621
            If lErro <> SUCESSO Then gError 42159 'Se não encontrou
        
        End If

    End If
        
    Exit Sub

Erro_Serie_Validate:

    Cancel = True

    Select Case gErr

        Case 42159
            'Pergunta se deseja criar nova série
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_SERIE", Serie.Text)
            'Se a resposta for afirmativa
            If vbMsgRes = vbYes Then
                'Chama a tela de cadastro de séries
                Call Chama_Tela("SerieNFiscal", objSerie)
            End If
            'segura o foco na série

        Case 42158, 51621

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157188)

    End Select

End Sub

Private Sub Serie_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Serie_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub optNFe_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub QtdeParcelas_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QtdeParcelas_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub QtdeParcelas_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub QtdeParcelas_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = QtdeParcelas
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub UltParcCobrada_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UltParcCobrada_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub UltParcCobrada_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub UltParcCobrada_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = UltParcCobrada
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_QtdeParcelas(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_QtdeParcelas

    Set objGridInt.objControle = QtdeParcelas

    If Len(Trim(QtdeParcelas.ClipText)) <> 0 Then

        lErro = Valor_Positivo_Critica(QtdeParcelas.Text)
        If lErro <> SUCESSO Then gError 129440

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129444

    Saida_Celula_QtdeParcelas = SUCESSO

    Exit Function

Erro_Saida_Celula_QtdeParcelas:

    Saida_Celula_QtdeParcelas = gErr

    Select Case gErr

        Case 129440, 129444
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 155016)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UltParcCobrada(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_UltParcCobrada

    Set objGridInt.objControle = UltParcCobrada

    If Len(Trim(UltParcCobrada.ClipText)) <> 0 Then

        lErro = Valor_Positivo_Critica(UltParcCobrada.Text)
        If lErro <> SUCESSO Then gError 129440

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129444

    Saida_Celula_UltParcCobrada = SUCESSO

    Exit Function

Erro_Saida_Celula_UltParcCobrada:

    Saida_Celula_UltParcCobrada = gErr

    Select Case gErr

        Case 129440, 129444
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 155016)

    End Select

    Exit Function

End Function

