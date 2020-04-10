VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.UserControl PedComprasConsOcx 
   ClientHeight    =   9195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   KeyPreview      =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   16995
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   8295
      Index           =   2
      Left            =   90
      TabIndex        =   3
      Top             =   675
      Visible         =   0   'False
      Width           =   16695
      Begin VB.Frame Frame10 
         Caption         =   "Notas Fiscais de Entrada"
         Height          =   2595
         Left            =   30
         TabIndex        =   108
         Top             =   5700
         Width           =   9030
         Begin MSMask.MaskEdBox ItemNF 
            Height          =   225
            Left            =   5280
            TabIndex        =   140
            Top             =   630
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   5
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox UMNF 
            Height          =   225
            Left            =   3360
            TabIndex        =   25
            Top             =   330
            Width           =   1365
            _ExtentX        =   2408
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
         Begin MSMask.MaskEdBox QuantNF 
            Height          =   225
            Left            =   2265
            TabIndex        =   24
            Top             =   315
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
         Begin MSMask.MaskEdBox NFiscal 
            Height          =   225
            Left            =   360
            TabIndex        =   22
            Top             =   315
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   5
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Serie 
            Height          =   225
            Left            =   1305
            TabIndex        =   23
            Top             =   300
            Width           =   870
            _ExtentX        =   1535
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
         Begin MSFlexGridLib.MSFlexGrid GridNFs 
            Height          =   2325
            Left            =   45
            TabIndex        =   21
            Top             =   210
            Width           =   8805
            _ExtentX        =   15531
            _ExtentY        =   4101
            _Version        =   393216
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Valores"
         Height          =   2580
         Index           =   1
         Left            =   9405
         TabIndex        =   80
         Top             =   5715
         Width           =   7245
         Begin VB.Label DescontoValor 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5295
            TabIndex        =   94
            Top             =   660
            Width           =   1245
         End
         Begin VB.Label ValorProdutos 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   150
            TabIndex        =   93
            Top             =   660
            Width           =   1245
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
            Left            =   315
            TabIndex        =   92
            Top             =   465
            Width           =   765
         End
         Begin VB.Label ValorFrete 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1425
            TabIndex        =   91
            Top             =   660
            Width           =   1245
         End
         Begin VB.Label ValorSeguro 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2700
            TabIndex        =   90
            Top             =   660
            Width           =   1245
         End
         Begin VB.Label OutrasDespesas 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4005
            TabIndex        =   89
            Top             =   660
            Width           =   1245
         End
         Begin VB.Label IPIValor 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   135
            TabIndex        =   88
            Top             =   1425
            Width           =   1245
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
            Left            =   4110
            TabIndex        =   87
            Top             =   465
            Width           =   840
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
            Left            =   1770
            TabIndex        =   86
            Top             =   465
            Width           =   450
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
            Left            =   5445
            TabIndex        =   85
            Top             =   465
            Width           =   825
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
            Left            =   2910
            TabIndex        =   84
            Top             =   465
            Width           =   615
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
            Left            =   570
            TabIndex        =   83
            Top             =   1230
            Width           =   255
         End
         Begin VB.Label ValorTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1410
            TabIndex        =   82
            Top             =   1425
            Width           =   1245
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
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
            Left            =   1740
            TabIndex        =   81
            Top             =   1230
            Width           =   450
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Itens"
         Height          =   5625
         Left            =   30
         TabIndex        =   79
         Top             =   30
         Width           =   16620
         Begin VB.CommandButton BotaoEntrega 
            Caption         =   "Datas de Entrega"
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
            TabIndex        =   143
            Top             =   5190
            Width           =   1695
         End
         Begin VB.TextBox DescCompleta 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   870
            MaxLength       =   50
            TabIndex        =   135
            Top             =   975
            Width           =   5460
         End
         Begin MSMask.MaskEdBox TotalMoedaReal 
            Height          =   228
            Left            =   5832
            TabIndex        =   131
            Top             =   936
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
            Left            =   4608
            TabIndex        =   132
            Top             =   936
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
         Begin VB.TextBox DescricaoProduto 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   1350
            MaxLength       =   50
            TabIndex        =   6
            Top             =   330
            Width           =   4000
         End
         Begin MSMask.MaskEdBox UnidadeMed 
            Height          =   225
            Left            =   2940
            TabIndex        =   7
            Top             =   270
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   5
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PrecoTotal 
            Height          =   225
            Left            =   7545
            TabIndex        =   11
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
         Begin MSMask.MaskEdBox PrecoUnitario 
            Height          =   225
            Left            =   6480
            TabIndex        =   10
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
         Begin MSMask.MaskEdBox QuantRecebida 
            Height          =   225
            Left            =   5130
            TabIndex        =   9
            Top             =   285
            Width           =   1260
            _ExtentX        =   2223
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
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   225
            Left            =   4110
            TabIndex        =   8
            Top             =   240
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
            Left            =   15
            TabIndex        =   5
            Top             =   330
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.TextBox Observacao 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   5025
            MaxLength       =   255
            TabIndex        =   20
            Top             =   1875
            Width           =   2445
         End
         Begin VB.ComboBox RecebForaFaixa 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "PedComprasConsOcx.ctx":0000
            Left            =   2745
            List            =   "PedComprasConsOcx.ctx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1800
            Width           =   2235
         End
         Begin MSMask.MaskEdBox DataLimite 
            Height          =   225
            Left            =   2910
            TabIndex        =   14
            Top             =   1515
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox AliquotaICM 
            Height          =   225
            Left            =   6030
            TabIndex        =   17
            Top             =   1515
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
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
         Begin MSMask.MaskEdBox ValorIPI 
            Height          =   225
            Left            =   5010
            TabIndex        =   16
            Top             =   1515
            Width           =   960
            _ExtentX        =   1693
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
         Begin MSMask.MaskEdBox AliquotaIPI 
            Height          =   225
            Left            =   4050
            TabIndex        =   15
            Top             =   1515
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
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
         Begin MSMask.MaskEdBox ValorDesconto 
            Height          =   225
            Left            =   1830
            TabIndex        =   13
            Top             =   1560
            Width           =   1035
            _ExtentX        =   1826
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
         Begin MSMask.MaskEdBox PercentDesc 
            Height          =   225
            Left            =   855
            TabIndex        =   12
            Top             =   1515
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
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
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PercentMaisReceb 
            Height          =   225
            Left            =   105
            TabIndex        =   18
            Top             =   1875
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
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
         Begin MSFlexGridLib.MSFlexGrid GridItens 
            Height          =   1680
            Left            =   45
            TabIndex        =   4
            Top             =   225
            Width           =   16455
            _ExtentX        =   29025
            _ExtentY        =   2963
            _Version        =   393216
            Rows            =   6
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox TempoTransito 
            Height          =   225
            Left            =   5100
            TabIndex        =   144
            Top             =   3915
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   3
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DeliveryDate 
            Height          =   225
            Left            =   3750
            TabIndex        =   145
            Top             =   3960
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   8220
      Index           =   6
      Left            =   270
      TabIndex        =   121
      Top             =   780
      Visible         =   0   'False
      Width           =   16515
      Begin VB.Frame Frame11 
         Caption         =   "Notas"
         Height          =   7710
         Left            =   15
         TabIndex        =   122
         Top             =   165
         Width           =   16230
         Begin VB.TextBox NotaPC 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   1260
            MaxLength       =   150
            TabIndex        =   123
            Top             =   570
            Width           =   14145
         End
         Begin MSFlexGridLib.MSFlexGrid GridNotas 
            Height          =   3915
            Left            =   210
            TabIndex        =   124
            Top             =   255
            Width           =   15570
            _ExtentX        =   27464
            _ExtentY        =   6906
            _Version        =   393216
            Rows            =   5
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   0   'False
            FocusRect       =   2
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   8160
      Index           =   5
      Left            =   180
      TabIndex        =   112
      Top             =   780
      Visible         =   0   'False
      Width           =   16575
      Begin VB.Frame SSFrame1 
         Caption         =   "Bloqueios"
         Height          =   8025
         Left            =   105
         TabIndex        =   113
         Top             =   90
         Width           =   16350
         Begin VB.ComboBox TipoBloqueio 
            Height          =   315
            ItemData        =   "PedComprasConsOcx.ctx":0004
            Left            =   180
            List            =   "PedComprasConsOcx.ctx":0006
            TabIndex        =   114
            Top             =   570
            Width           =   2000
         End
         Begin MSMask.MaskEdBox ResponsavelLib 
            Height          =   270
            Left            =   6975
            TabIndex        =   115
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
            TabIndex        =   116
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
            TabIndex        =   117
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
            TabIndex        =   118
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
            TabIndex        =   119
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
            Height          =   7425
            Left            =   135
            TabIndex        =   120
            Top             =   375
            Width           =   16080
            _ExtentX        =   28363
            _ExtentY        =   13097
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
      Height          =   8040
      Index           =   4
      Left            =   165
      TabIndex        =   27
      Top             =   840
      Visible         =   0   'False
      Width           =   16575
      Begin VB.Frame Frame5 
         Caption         =   "Distribuição dos Produtos"
         Height          =   7410
         Left            =   45
         TabIndex        =   78
         Top             =   210
         Width           =   16335
         Begin VB.TextBox DescProduto 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   30
            Top             =   330
            Width           =   4000
         End
         Begin MSMask.MaskEdBox Quant 
            Height          =   225
            Left            =   6240
            TabIndex        =   34
            Top             =   360
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
         Begin MSMask.MaskEdBox UM 
            Height          =   225
            Left            =   5175
            TabIndex        =   33
            Top             =   240
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   5
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CentroCusto 
            Height          =   225
            Left            =   2880
            TabIndex        =   31
            Top             =   360
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Prod 
            Height          =   225
            Left            =   270
            TabIndex        =   29
            Top             =   360
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Almoxarifado 
            Height          =   225
            Left            =   4155
            TabIndex        =   32
            Top             =   360
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridDistribuicao 
            Height          =   6585
            Left            =   270
            TabIndex        =   28
            Top             =   495
            Width           =   15840
            _ExtentX        =   27940
            _ExtentY        =   11615
            _Version        =   393216
            Rows            =   7
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox ContaContabil 
            Height          =   225
            Left            =   7035
            TabIndex        =   111
            Top             =   240
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8265
      Index           =   3
      Left            =   105
      TabIndex        =   26
      Top             =   690
      Visible         =   0   'False
      Width           =   16680
      Begin VB.Frame Frame2 
         Caption         =   "Local de Entrega"
         Height          =   2910
         Left            =   210
         TabIndex        =   49
         Top             =   180
         Width           =   8505
         Begin VB.Frame FrameTipo 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   795
            Index           =   0
            Left            =   4800
            TabIndex        =   50
            Top             =   330
            Width           =   3495
            Begin VB.Label Label37 
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
               Left            =   390
               TabIndex        =   52
               Top             =   195
               Width           =   465
            End
            Begin VB.Label FilialEmpresa 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   960
               TabIndex        =   51
               Top             =   120
               Width           =   2145
            End
         End
         Begin VB.Frame FrameTipo 
            BorderStyle     =   0  'None
            Height          =   675
            Index           =   1
            Left            =   4785
            TabIndex        =   53
            Top             =   375
            Visible         =   0   'False
            Width           =   3495
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
               Index           =   1
               Left            =   90
               TabIndex        =   57
               Top             =   60
               Width           =   1035
            End
            Begin VB.Label Fornec 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   1200
               TabIndex        =   56
               Top             =   0
               Width           =   2145
            End
            Begin VB.Label Label32 
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
               Left            =   600
               TabIndex        =   55
               Top             =   405
               Width           =   465
            End
            Begin VB.Label FilialFornec 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   1170
               TabIndex        =   54
               Top             =   360
               Width           =   2145
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Tipo"
            Height          =   555
            Left            =   270
            TabIndex        =   58
            Top             =   450
            Width           =   3720
            Begin VB.OptionButton TipoDestino 
               Caption         =   "Fornecedor"
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
               Index           =   1
               Left            =   2145
               TabIndex        =   60
               Top             =   240
               Width           =   1335
            End
            Begin VB.OptionButton TipoDestino 
               Caption         =   "Filial Empresa"
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
               Index           =   0
               Left            =   480
               TabIndex        =   59
               Top             =   240
               Width           =   1515
            End
         End
         Begin VB.Label Pais 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4020
            TabIndex        =   72
            Top             =   2355
            Width           =   1995
         End
         Begin VB.Label Estado 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1260
            TabIndex        =   71
            Top             =   2400
            Width           =   495
         End
         Begin VB.Label CEP 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6675
            TabIndex        =   70
            Top             =   1920
            Width           =   930
         End
         Begin VB.Label Cidade 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4020
            TabIndex        =   69
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label Bairro 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1260
            TabIndex        =   68
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label Endereco 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1260
            TabIndex        =   67
            Top             =   1500
            Width           =   6345
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
            TabIndex        =   66
            Top             =   2400
            Width           =   495
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
            TabIndex        =   65
            Top             =   1995
            Width           =   465
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
            TabIndex        =   64
            Top             =   1995
            Width           =   585
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
            TabIndex        =   63
            Top             =   2400
            Width           =   675
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
            TabIndex        =   62
            Top             =   1995
            Width           =   675
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
            TabIndex        =   61
            Top             =   1515
            Width           =   915
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frete"
         Height          =   885
         Left            =   180
         TabIndex        =   73
         Top             =   3165
         Width           =   8535
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
            Left            =   4020
            TabIndex        =   77
            Top             =   435
            Width           =   1410
         End
         Begin VB.Label TipoFrete 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1590
            TabIndex        =   76
            Top             =   390
            Width           =   825
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
            Left            =   555
            TabIndex        =   75
            Top             =   450
            Width           =   945
         End
         Begin VB.Label Transportadora 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   5490
            TabIndex        =   74
            Top             =   390
            Width           =   2190
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   8235
      Index           =   1
      Left            =   105
      TabIndex        =   1
      Top             =   735
      Width           =   16620
      Begin VB.CommandButton BotaoCancelarBaixa 
         Caption         =   "Cancelar Baixa"
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
         Left            =   210
         TabIndex        =   137
         Top             =   6015
         Width           =   2220
      End
      Begin VB.CommandButton BotaoGerador 
         Caption         =   "Documento Gerador"
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
         Left            =   6675
         TabIndex        =   2
         Top             =   6015
         Width           =   2220
      End
      Begin VB.Frame Frame8 
         Caption         =   "Datas"
         Height          =   1050
         Left            =   225
         TabIndex        =   96
         Top             =   4455
         Width           =   9450
         Begin VB.Label DataRefFluxo 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7305
            TabIndex        =   139
            Top             =   645
            Width           =   1095
         End
         Begin VB.Label Label8 
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
            Left            =   6675
            TabIndex        =   138
            Top             =   675
            Width           =   525
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Baixa:"
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
            Left            =   6660
            TabIndex        =   110
            Top             =   270
            Width           =   540
         End
         Begin VB.Label DataBaixa 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7305
            TabIndex        =   109
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Data 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1770
            TabIndex        =   45
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label DataAlteracao 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1770
            TabIndex        =   46
            Top             =   645
            Width           =   1095
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Alteração:"
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
            TabIndex        =   105
            Top             =   675
            Width           =   885
         End
         Begin VB.Label DataEmissao 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4530
            TabIndex        =   47
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Emissão:"
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
            Left            =   3690
            TabIndex        =   104
            Top             =   300
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Envio:"
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
            Left            =   3900
            TabIndex        =   103
            Top             =   675
            Width           =   555
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Registro:"
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
            Left            =   885
            TabIndex        =   102
            Top             =   300
            Width           =   780
         End
         Begin VB.Label DataEnvio 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4530
            TabIndex        =   48
            Top             =   645
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Cabeçalho"
         Height          =   3900
         Left            =   240
         TabIndex        =   95
         Top             =   435
         Width           =   9450
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Tabela de Preço:"
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
            TabIndex        =   142
            Top             =   1995
            Width           =   1470
         End
         Begin VB.Label TabelaPreco 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6000
            TabIndex        =   141
            Top             =   1935
            Width           =   2145
         End
         Begin VB.Label Label5 
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
            Left            =   645
            TabIndex        =   134
            Top             =   1995
            Width           =   1035
         End
         Begin VB.Label LabelObsEmbalagem 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1740
            TabIndex        =   133
            Top             =   1935
            Width           =   2145
         End
         Begin VB.Label EmbalagemLabel 
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
            TabIndex        =   130
            Top             =   1995
            Width           =   1035
         End
         Begin VB.Label Taxa 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6000
            TabIndex        =   129
            Top             =   1545
            Width           =   2175
         End
         Begin VB.Label Moeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1755
            TabIndex        =   128
            Top             =   1530
            Width           =   2145
         End
         Begin VB.Label TaxaLabel 
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
            Left            =   5445
            TabIndex        =   127
            Top             =   1590
            Width           =   495
         End
         Begin VB.Label MoedaLabel 
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
            Left            =   1050
            TabIndex        =   126
            Top             =   1590
            Width           =   645
         End
         Begin VB.Label Embalagem 
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   -10000
            TabIndex        =   125
            Top             =   1935
            Width           =   2145
         End
         Begin VB.Label Label4 
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
            Index           =   0
            Left            =   600
            TabIndex        =   107
            Top             =   2385
            Width           =   1095
         End
         Begin VB.Label Observ 
            BorderStyle     =   1  'Fixed Single
            Height          =   1455
            Left            =   1755
            TabIndex        =   41
            Top             =   2340
            Width           =   6435
         End
         Begin VB.Label Contato 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6000
            TabIndex        =   44
            Top             =   1095
            Width           =   2175
         End
         Begin VB.Label Label21 
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
            Left            =   5205
            TabIndex        =   106
            Top             =   1155
            Width           =   735
         End
         Begin VB.Label Codigo 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1755
            TabIndex        =   38
            Top             =   285
            Width           =   900
         End
         Begin VB.Label CondPagto 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1755
            TabIndex        =   40
            Top             =   1095
            Width           =   2145
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
            Index           =   0
            Left            =   660
            TabIndex        =   101
            Top             =   750
            Width           =   1035
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
            Index           =   0
            Left            =   5475
            TabIndex        =   100
            Top             =   750
            Width           =   465
         End
         Begin VB.Label Fornecedor 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1755
            TabIndex        =   39
            Top             =   690
            Width           =   2145
         End
         Begin VB.Label Filial 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6000
            TabIndex        =   43
            Top             =   705
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
            Left            =   4965
            TabIndex        =   99
            Top             =   345
            Width           =   975
         End
         Begin VB.Label Comprador 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6000
            TabIndex        =   42
            Top             =   285
            Width           =   2145
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
            Left            =   630
            TabIndex        =   98
            Top             =   1155
            Width           =   1065
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
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   765
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   97
            Top             =   345
            Width           =   930
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   525
      Left            =   15270
      ScaleHeight     =   465
      ScaleWidth      =   1560
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   45
      Width           =   1620
      Begin VB.CommandButton BotaoEmail 
         Height          =   345
         Left            =   45
         Picture         =   "PedComprasConsOcx.ctx":0008
         Style           =   1  'Graphical
         TabIndex        =   136
         ToolTipText     =   "Enviar email"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   555
         Picture         =   "PedComprasConsOcx.ctx":09AA
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Imprimir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1050
         Picture         =   "PedComprasConsOcx.ctx":0AAC
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   8760
      Left            =   45
      TabIndex        =   0
      Top             =   345
      Width           =   16860
      _ExtentX        =   29739
      _ExtentY        =   15452
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
Attribute VB_Name = "PedComprasConsOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iLinhaAnt As Integer

Dim iAlterado As Integer
Dim iFrameAtual As Integer
Dim objGridItens As AdmGrid
Dim objGridDistribuicao As AdmGrid
Dim objGridNF As AdmGrid
Dim gcolItemPedido As Collection
Dim gobjPC As ClassPedidoCompras
Dim iFrameTipoDestinoAtual As Integer
Dim iChamaTela As Integer
Dim iFrameDestinoAtual As Integer

Dim objGridBloqueio As AdmGrid
Dim iGrid_TipoBloqueio_Col As Integer
Dim iGrid_DataBloqueio_Col As Integer
Dim iGrid_CodUsuario_Col As Integer
Dim iGrid_ResponsavelBL_Col As Integer
Dim iGrid_DataLiberacao_Col As Integer
Dim iGrid_ResponsavelLib_Col As Integer

'GridItens
Dim iGrid_Produto_Col As Integer
Dim iGrid_DescricaoProduto_Col As Integer
Dim iGrid_UnidadeMed_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_QuantRecebida_Col As Integer
Dim iGrid_PrecoUnitario_Col As Integer
Dim iGrid_PercentDesc_Col As Integer
Dim iGrid_Desconto_Col As Integer
Dim iGrid_PrecoTotal_Col As Integer
Dim iGrid_PrecoUnitarioMoedaReal_Col As Integer
Dim iGrid_TotalMoedaReal_Col As Integer
Dim iGrid_DataLimite_Col As Integer
Dim iGrid_AliquotaIPI_Col As Integer
Dim iGrid_ValorIPIItem_Col As Integer
Dim iGrid_AliquotaICMS_Col As Integer
Dim iGrid_PercentMaisReceb_Col As Integer
Dim iGrid_RecebForaFaixa_Col As Integer
Dim iGrid_Observacao_Col As Integer
Dim iGrid_DescCompleta_Col As Integer 'leo
Public iGrid_DeliveryDate_Col As Integer
Public iGrid_TempoTransito_Col As Integer

'GridDistribuicao
Dim iGrid_Prod_Col As Integer
Dim iGrid_DescProduto_Col As Integer
Dim iGrid_CentroCusto_Col As Integer
Dim iGrid_Almoxarifado_Col As Integer
Dim iGrid_UM_Col As Integer
Dim iGrid_Quant_Col As Integer
Dim iGrid_ContaContabil_Col As Integer

'GridNF
Dim iGrid_Serie_Col As Integer
Dim iGrid_NFiscal_Col As Integer
Dim iGrid_QuantNF_Col As Integer
Dim iGrid_UMNF_Col As Integer
Dim iGrid_ItemNF_Col As Integer

Dim objGridNotas As AdmGrid
Dim iGrid_NotaPC_Col As Integer

Dim bExibirColReal As Boolean

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1


Public Sub Form_Load()

Dim lErro As Long
Dim objComprador As New ClassComprador

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    iLinhaAnt = 0
    
    bExibirColReal = True

    '#############################
    'Inserido por Wagner
    Call Formata_Controles
    '#############################

    Set gcolItemPedido = New Collection
    Set gobjPC = New ClassPedidoCompras

    Set objGridBloqueio = New AdmGrid
    Set objGridItens = New AdmGrid
    Set objGridDistribuicao = New AdmGrid
    Set objGridNF = New AdmGrid
    Set objGridNotas = New AdmGrid
    
    Set objEventoCodigo = New AdmEvento

    'Faz a inicializacao do grid bloqueio
    lErro = Inicializa_Grid_Bloqueios(objGridBloqueio)
    If lErro <> SUCESSO Then gError 53187
    
    'Faz a inicializacao do grid itens
    lErro = Inicializa_Grid_Itens(objGridItens)
    If lErro <> SUCESSO Then gError 68402

    'Faz a inicializacao do grid distribuicao
    lErro = Inicializa_Grid_Distribuicao(objGridDistribuicao)
    If lErro <> SUCESSO Then gError 68403
    
    'Faz a inicializacao do grid NF
    lErro = Inicializa_Grid_NF(objGridNF)
    If lErro <> SUCESSO Then gError 68404
    
    'Inicializa mascara do produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 89167
    
    'Inicializa mascara da conta
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabil)
    If lErro <> SUCESSO Then gError 89256
    
    'Inicializa mascara do centro de custo
    lErro = CF("Inicializa_Mascara_Ccl_MaskEd", CentroCusto)
    If lErro <> SUCESSO Then gError 89256
    
    lErro = Inicializa_Grid_Notas(objGridNotas)
    If lErro <> SUCESSO Then gError 103327
    
    'Carrega a combo de Tipos de Bloqueio
    lErro = Carrega_TipoBloqueio()
    If lErro <> SUCESSO Then gError 53178
    
    Call Carrega_RecebForaFaixa

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 53178, 53187, 68402, 68403, 68404, 89167, 89256, 103327
            'Erros tratados nas rotinas chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164337)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Private Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Itens

Dim iIncremento As Integer

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Quant Recebida")
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Total")
    'Se a moeda for Diferente de Real => Exibe as Colunas de Comparacao
    If bExibirColReal = True Then
        objGridInt.colColuna.Add ("Preço (R$)")
        objGridInt.colColuna.Add ("Total (R$)")
    End If
    If gobjCOM.iPCExibeDeliveryDate = MARCADO Then
        objGridInt.colColuna.Add ("Delivery Date")
        objGridInt.colColuna.Add ("Tempo de Trânsito")
    End If
    objGridInt.colColuna.Add ("Data Limite")
    objGridInt.colColuna.Add ("Alíquota IPI")
    objGridInt.colColuna.Add ("Valor IPI ")
    objGridInt.colColuna.Add ("Alíquota ICMS")
    objGridInt.colColuna.Add ("% a Mais Receb")
    objGridInt.colColuna.Add ("Ação Receb Fora Faixa")
    objGridInt.colColuna.Add ("Observação")
    objGridInt.colColuna.Add ("Desc. Completa") 'leo

    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoProduto.Name)
    objGridInt.colCampo.Add (UnidadeMed.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (QuantRecebida.Name)
    objGridInt.colCampo.Add (PrecoUnitario.Name)
    objGridInt.colCampo.Add (PercentDesc.Name)
    objGridInt.colCampo.Add (ValorDesconto.Name)
    objGridInt.colCampo.Add (PrecoTotal.Name)
    'Se a moeda for Diferente de Real => Exibe as Colunas de Comparacao
    If bExibirColReal Then
        objGridInt.colCampo.Add (PrecoUnitarioMoedaReal.Name)
        objGridInt.colCampo.Add (TotalMoedaReal.Name)
    Else
        PrecoUnitarioMoedaReal.left = POSICAO_FORA_TELA
        PrecoUnitarioMoedaReal.TabStop = False
    
        TotalMoedaReal.left = POSICAO_FORA_TELA
        TotalMoedaReal.TabStop = False
    End If
    
    If gobjCOM.iPCExibeDeliveryDate = MARCADO Then
        objGridInt.colCampo.Add (DeliveryDate.Name)
        objGridInt.colCampo.Add (TempoTransito.Name)
    Else
        DeliveryDate.left = POSICAO_FORA_TELA
        DeliveryDate.TabStop = False
    
        TempoTransito.left = POSICAO_FORA_TELA
        TempoTransito.TabStop = False
    End If
    objGridInt.colCampo.Add (DataLimite.Name)
    objGridInt.colCampo.Add (AliquotaIPI.Name)
    objGridInt.colCampo.Add (ValorIPI.Name)
    objGridInt.colCampo.Add (AliquotaICM.Name)
    objGridInt.colCampo.Add (PercentMaisReceb.Name)
    objGridInt.colCampo.Add (RecebForaFaixa.Name)
    objGridInt.colCampo.Add (Observacao.Name)
    objGridInt.colCampo.Add (DescCompleta.Name) 'leo

   'indica onde estao situadas as colunas do grid
    iGrid_Produto_Col = 1
    iGrid_DescProduto_Col = 2
    iGrid_UnidadeMed_Col = 3
    iGrid_Quantidade_Col = 4
    iGrid_QuantRecebida_Col = 5
    iGrid_PrecoUnitario_Col = 6
    iGrid_PercentDesc_Col = 7
    iGrid_Desconto_Col = 8
    iGrid_PrecoTotal_Col = 9
    
    iIncremento = 0
    
    If bExibirColReal Then
        iGrid_PrecoUnitarioMoedaReal_Col = 10 + iIncremento
        iGrid_TotalMoedaReal_Col = 11 + iIncremento
        iIncremento = iIncremento + 2
    End If
    
    If gobjCOM.iPCExibeDeliveryDate = MARCADO Then
        iGrid_DeliveryDate_Col = 10 + iIncremento
        iGrid_TempoTransito_Col = 11 + iIncremento
        iIncremento = iIncremento + 2
    End If
    
    iGrid_DataLimite_Col = 10 + iIncremento
    iGrid_AliquotaIPI_Col = 11 + iIncremento
    iGrid_ValorIPIItem_Col = 12 + iIncremento
    iGrid_AliquotaICMS_Col = 13 + iIncremento
    iGrid_PercentMaisReceb_Col = 14 + iIncremento
    iGrid_RecebForaFaixa_Col = 15 + iIncremento
    iGrid_Observacao_Col = 16 + iIncremento
    iGrid_DescCompleta_Col = 17 + iIncremento

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridItens

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_PEDIDO_COMPRAS + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 12

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'proibido incluir e excluir linhas
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Itens = SUCESSO

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
    objGridInt.colCampo.Add (DescProduto.Name)
    objGridInt.colCampo.Add (CentroCusto.Name)
    objGridInt.colCampo.Add (Almoxarifado.Name)
    objGridInt.colCampo.Add (UM.Name)
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
    objGridInt.iLinhasVisiveis = 25

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'proibido incluir e excluir linhas
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Distribuicao = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_NF(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid NF

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Série")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Unidade Medida")

    objGridInt.colCampo.Add (Serie.Name)
    objGridInt.colCampo.Add (NFiscal.Name)
    objGridInt.colCampo.Add (ItemNF.Name)
    objGridInt.colCampo.Add (QuantNF.Name)
    objGridInt.colCampo.Add (UMNF.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_Serie_Col = 1
    iGrid_NFiscal_Col = 2
    iGrid_ItemNF_Col = 3
    iGrid_QuantNF_Col = 4
    iGrid_UMNF_Col = 5

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridNFs

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_NFS_ITEMPED + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 8

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'proibido incluir e excluir linhas
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_NF = SUCESSO

    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras
Dim sNomeRed As String

On Error GoTo Erro_Tela_Extrai

    sTabela = "PedCompraTodos_Fornecedor"
    
    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then gError 68405
    
    sNomeRed = Fornecedor.Caption
    
    'Preenche a coleção colCampoValor
    colCampoValor.Add "Codigo", objPedidoCompra.lCodigo, 0, "Codigo"
    colCampoValor.Add "OutrasDespesas", objPedidoCompra.dOutrasDespesas, 0, "OutrasDespesas"
    colCampoValor.Add "Data", objPedidoCompra.dtData, 0, "Data"
    colCampoValor.Add "DataAlteracao", objPedidoCompra.dtDataAlteracao, 0, "DataAlteracao"
    colCampoValor.Add "DataEnvio", objPedidoCompra.dtDataEnvio, 0, "DataEnvio"
    colCampoValor.Add "DataEmissao", objPedidoCompra.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "DataBaixa", objPedidoCompra.dtDataBaixa, 0, "DataBaixa"
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
    colCampoValor.Add "Embalagem", objPedidoCompra.iEmbalagem, 0, "Embalagem"
    colCampoValor.Add "Taxa", objPedidoCompra.dTaxa, 0, "Taxa"
    colCampoValor.Add "Moeda", objPedidoCompra.iMoeda, 0, "Moeda"
    colCampoValor.Add "ObsEmbalagem", objPedidoCompra.sObsEmbalagem, STRING_BUFFER_MAX_TEXTO, "ObsEmbalagem"
                
''    colCampoValor.Add "NomeReduzido", sNomeRed, STRING_BUFFER_MAX_TEXTO, "NomeReduzido"
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 68405
            'Erro tratado na rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164338)

    End Select

    Exit Sub


End Sub

Private Sub BotaoEmail_Click()

Dim lErro As Long, objBloqueioPC As ClassBloqueioPC
Dim objPedidoCompra As New ClassPedidoCompras
Dim objRelatorio As New AdmRelatorio
Dim sMailTo As String, sFiltro As String
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objEndereco As New ClassEndereco, sInfoEmail As String

On Error GoTo Erro_BotaoEmail_Click

    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then gError 68450

    If objPedidoCompra.lCodigo = 0 Then gError 76052

    lErro = CF("PedidoCompra_Le_Todos", objPedidoCompra)
    If lErro <> SUCESSO And lErro <> 68486 Then gError 76033
    
    'Se o Pedido não existe ==> erro
    If lErro = 68486 Then gError 76034
    
    If objPedidoCompra.dtDataRegAprov = DATA_NULA Then
        If gobjCOM.iPedCompraBloqEnvioSemAprov = MARCADO Then gError 213170
    End If
        
    lErro = CF("BloqueiosPC_Le", objPedidoCompra)
    If lErro <> SUCESSO Then gError 76058
    
    For Each objBloqueioPC In objPedidoCompra.colBloqueiosPC
            
        If objBloqueioPC.dtDataLib = DATA_NULA Then gError 76051
    
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
    DataEmissao.Caption = Format(gdtDataHoje, "dd/mm/yyyy")

    objPedidoCompra.dtDataEmissao = gdtDataHoje

    'Verifica se o Pedido de Compra está baixado
    If objPedidoCompra.dtDataBaixa <> DATA_NULA Then
    
        lErro = CF("PedidoCompraBaixado_Atualiza_DataEmissao", objPedidoCompra)
        If lErro <> SUCESSO And lErro <> 76070 Then gError 76064
        
        If lErro = 76070 Then gError 76074
        
    'Se o Pedido de Compra não está baixado
    Else
    
        'Atualiza data de emissao no BD para a data atual
        lErro = CF("PedidoCompra_Atualiza_DataEmissao", objPedidoCompra, True)
        If lErro <> SUCESSO And lErro <> 56348 Then gError 68451

        'se nao encontrar ---> erro
        If lErro = 56348 Then gError 68452
    
    End If
    
    sFiltro = "REL_PCOM.PC_NumIntDoc = @NPEDCOM"
    lErro = CF("Relatorio_ObterFiltro", "Pedido de Compra Consulta", sFiltro)
    If lErro <> SUCESSO Then gError 76035
    
    'Executa o relatório
    lErro = objRelatorio.ExecutarDiretoEmail("Pedido de Compra Consulta", sFiltro, 0, "PEDCOM", "NPEDCOM", objPedidoCompra.lNumIntDoc, "TTO_EMAIL", sMailTo, "TSUBJECT", "Pedido de Compra " & CStr(objPedidoCompra.lCodigo), "TALIASATTACH", "PedCompra" & CStr(objPedidoCompra.lCodigo), "TINFO_EMAIL", sInfoEmail)
    If lErro <> SUCESSO Then gError 76035
    
    Exit Sub

Erro_BotaoEmail_Click:

    Select Case gErr

        Case 68450, 68451, 76058, 129314, 129315
            'Erros tratados nas rotinas chamadas
            
        Case 68452, 76034, 76074
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", gErr, objPedidoCompra.lCodigo)

        Case 76033, 76035, 76064
        
        Case 76051
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_BLOQUEADO", gErr, objPedidoCompra.lCodigo)
            
        Case 76052
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDCOMPRA_IMPRESSAO", gErr)
            
        Case 129316
            Call Rotina_Erro(vbOKOnly, "ERRO_EMAIL_NAO_ENCONTRADO", gErr, objPedidoCompra.lCodigo)
            
        Case 213170
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_APROVADO", gErr, objPedidoCompra.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164339)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Activate()

Dim colSelecao As New Collection
Dim objPedidoCompra As ClassPedidoCompras

    'Se no Trata_Parametros nenhuma Pedido de Compras foi passado
    If iChamaTela = 1 Then
    
        'Chama a tela PedComprasTodosLista
        Call Chama_Tela("PedComprasTodosLista", colSelecao, objPedidoCompra, objEventoCodigo)
        iChamaTela = 0
        
    End If
    
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

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
    objPedidoCompra.dtDataBaixa = colCampoValor.Item("DataBaixa").vValor
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
    objPedidoCompra.iEmbalagem = colCampoValor.Item("Embalagem").vValor
    objPedidoCompra.iMoeda = colCampoValor.Item("Moeda").vValor
    objPedidoCompra.dTaxa = colCampoValor.Item("Taxa").vValor
    objPedidoCompra.sObsEmbalagem = colCampoValor.Item("ObsEmbalagem").vValor

    ' preenche a tela com os elementos do objPedidoCompra
    lErro = Trata_Parametros(objPedidoCompra)
    If lErro <> SUCESSO Then gError 68406

    iAlterado = 0

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 68406
            'Erro tratado na rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164340)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Private Function Traz_PedidoCompra_Tela(objPedidoCompra As ClassPedidoCompras) As Long
'Traz para a tela o Pedido de Compra armazenado em objPedidoCompra

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objItemPC As New ClassItemPedCompra
Dim objFilialEmpresa As New AdmFiliais
Dim objFilialCliente As New ClassFilialCliente
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim objObservacao As New ClassObservacao
Dim objFornecedor As New ClassFornecedor
Dim objCliente As New ClassCliente
Dim objEndereco As New ClassEndereco
Dim objTransportadora As New ClassTransportadora
Dim objComprador As New ClassComprador
Dim objUsuarios As New ClassUsuarios
Dim objEmbalagem As New ClassEmbalagem
Dim objMoeda As New ClassMoedas, objTabelaPreco As New ClassTabelaPreco

On Error GoTo Erro_Traz_PedidoCompra_Tela

    ' lê os itens do Pedido de compra
    lErro = CF("ItensPC_LeTodos", objPedidoCompra)
    If lErro <> SUCESSO Then gError 68407

    lErro = CF("BloqueiosPC_Le", objPedidoCompra)
    If lErro <> SUCESSO Then gError 86160


    lErro = CF("NotasPedCompras_Le", objPedidoCompra)
    If lErro <> SUCESSO Then gError 103355
    
    lErro = CF("ItensPCEntrega_Le", objPedidoCompra)
    If lErro <> SUCESSO Then gError 103355
   
   'Limpa a tela
    Call Limpa_Tela_PedidoComprasCons
 
    
    If objPedidoCompra.dTaxa > 0 Then
        Taxa.Caption = Format(objPedidoCompra.dTaxa, FORMATO_TAXA_CONVERSAO_MOEDA)
    Else
        Taxa.Caption = ""
    End If
    
    If objPedidoCompra.iEmbalagem > 0 Then
        
        objEmbalagem.iCodigo = objPedidoCompra.iEmbalagem
        
        lErro = CF("Embalagem_Le", objEmbalagem)
        If lErro <> SUCESSO And lErro <> 82763 Then gError 103390
        
        If lErro = SUCESSO Then
            Embalagem.Caption = objEmbalagem.sSigla
        Else
            Embalagem.Caption = ""
        End If
               
    End If
        
       
    objMoeda.iCodigo = objPedidoCompra.iMoeda
    
    lErro = Moedas_Le(objMoeda)
    If lErro <> SUCESSO And lErro <> 108821 Then gError 103391
    
    If lErro = SUCESSO Then
        Moeda.Caption = objMoeda.iCodigo & SEPARADOR & objMoeda.sNome
    Else
        Moeda.Caption = ""
    End If
        
    'Se a moeda selecionada for = REAL
    If objMoeda.iCodigo = MOEDA_REAL Then
    
        'Limpa a cotacao
        Taxa.Caption = ""
        
        bExibirColReal = False
        
    Else
            
        bExibirColReal = True
    
    End If
        
    'Coloca os dados na tela
    Codigo.Caption = objPedidoCompra.lCodigo
    Contato.Caption = objPedidoCompra.sContato

    If objPedidoCompra.iComprador <> 0 Then
    
        objComprador.iCodigo = objPedidoCompra.iComprador
        objComprador.iFilialEmpresa = objPedidoCompra.iFilialEmpresa
        
        'Lê o Comprador
        lErro = CF("Comprador_Le", objComprador)
        If lErro <> SUCESSO Then gError 68470
    
        objUsuarios.sCodUsuario = objComprador.sCodUsuario

        'le  o usuário contido na tabela de Usuarios
        lErro = CF("Usuarios_Le", objUsuarios, False)
        If lErro <> SUCESSO And lErro <> 40832 Then gError 68471
        If lErro <> SUCESSO Then gError 68472

        'Coloca nome reduzido do Comprador na tela
        Comprador.Caption = objUsuarios.sNomeReduzido

    End If
    
    Data.Caption = Format(objPedidoCompra.dtData, "dd/mm/yyyy")

    objFornecedor.lCodigo = objPedidoCompra.lFornecedor

    'Lê o Fornecedor
    lErro = CF("Fornecedor_Le", objFornecedor)
    If lErro <> SUCESSO And lErro <> 12729 Then gError 68429
    If lErro = 12729 Then gError 68430

    'Coloca o NomeReduzido do Fornecedor na tela
    Fornecedor.Caption = objFornecedor.sNomeReduzido

    'Passa o CodFornecedor e o CodFilial para o objfilialfornecedor
    objFilialFornecedor.lCodFornecedor = objPedidoCompra.lFornecedor
    objFilialFornecedor.iCodFilial = objPedidoCompra.iFilial

    'Lê o filialforncedor
    lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", objFornecedor.sNomeReduzido, objFilialFornecedor)
    If lErro <> SUCESSO And lErro <> 18272 Then gError 68409
    'Se nao encontrou
    If lErro = 18272 Then gError 68410

    'Coloca a filial na tela
    Filial.Caption = objPedidoCompra.iFilial & SEPARADOR & objFilialFornecedor.sNome

    If objPedidoCompra.dtDataAlteracao <> DATA_NULA Then
        DataAlteracao.Caption = Format(objPedidoCompra.dtDataAlteracao, "dd/mm/yyyy")
    End If
    
    If objPedidoCompra.dtDataEmissao <> DATA_NULA Then
        DataEmissao.Caption = Format(objPedidoCompra.dtDataEmissao, "dd/mm/yyyy")
    End If
    
    If objPedidoCompra.dtDataBaixa <> DATA_NULA Then
        DataBaixa.Caption = Format(objPedidoCompra.dtDataBaixa, "dd/mm/yyyy")
    End If
    
    'Preenche o TipoDestino
    TipoDestino(objPedidoCompra.iTipoDestino).Value = True

    If iFrameTipoDestinoAtual = TIPO_DESTINO_EMPRESA Then

        objFilialEmpresa.iCodFilial = objPedidoCompra.iFilialDestino

        'Lê a FilialEmpresa
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 68411
        If lErro = 27378 Then gError 68412

        'Coloca a FilialEmpresa na tela
        FilialEmpresa.Caption = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome

        ' preencher endereço
        If objFilialEmpresa.objEnderecoEntrega.lCodigo <> 0 Then
            Call Preenche_Endereco(objFilialEmpresa.objEnderecoEntrega)
        Else
            Call Preenche_Endereco(objFilialEmpresa.objEndereco)
        End If

    ElseIf iFrameTipoDestinoAtual = TIPO_DESTINO_FORNECEDOR Then

        objFornecedor.lCodigo = objPedidoCompra.lFornCliDestino

        'Lê o Fornecedor
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12729 Then gError 68413

        If lErro = 12729 Then gError 68414

        'Coloca o Fornecedor na tela.
        Fornec.Caption = objFornecedor.sNomeReduzido

        objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
        objFilialFornecedor.iCodFilial = objPedidoCompra.iFilialDestino

        'le a FilialFornecedor
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", objFornecedor.sNomeReduzido, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError 68415

        'Se nao encontrou
        If lErro = 18272 Then gError 68416

        'Coloca a Filial na tela
        FilialFornec.Caption = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome

        If objFornecedor.lEndereco > 0 Then

            objEndereco.lCodigo = objFornecedor.lEndereco

            lErro = CF("Endereco_Le", objEndereco)
            If lErro <> SUCESSO And lErro <> 12309 Then gError 68417

            'Se nao encontrou ---> erro
            If lErro = 12309 Then gError 68418

            ' preenche endereço
            Call Preenche_Endereco(objEndereco)

        End If

    End If

    'Verifica o TipoFrete
    If StrParaInt(objPedidoCompra.sTipoFrete) = TIPO_CIF Then

        TipoFrete.Caption = "CIF"

    Else

        TipoFrete.Caption = "FOB"

    End If

    If objPedidoCompra.iTransportadora <> 0 Then

        objTransportadora.iCodigo = objPedidoCompra.iTransportadora

        'le a transportadora
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError 68419

        'se nao encontrou ---> erro
        If lErro = 19250 Then gError 68420

        Transportadora.Caption = objTransportadora.sNomeReduzido

    End If

    If objPedidoCompra.iCondicaoPagto <> 0 Then

        objCondicaoPagto.iCodigo = objPedidoCompra.iCondicaoPagto

        'lê a cond. de pagto
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 68421

        'se nao encontrou --->erro
        If lErro = 19205 Then gError 68422

        CondPagto.Caption = objPedidoCompra.iCondicaoPagto & SEPARADOR & objCondicaoPagto.sDescReduzida

    End If

    If objPedidoCompra.iTabelaPreco <> 0 Then

        objTabelaPreco.iCodigo = objPedidoCompra.iTabelaPreco

        'Tenta ler TabelaPreço com esse código no BD
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError 68421

        'se nao encontrou --->erro
        If lErro <> SUCESSO Then gError 68422

        TabelaPreco.Caption = objPedidoCompra.iTabelaPreco & SEPARADOR & objTabelaPreco.sDescricao

    End If

    If objPedidoCompra.dtDataEnvio <> DATA_NULA Then
        DataEnvio.Caption = Format(objPedidoCompra.dtDataEnvio, "dd/mm/yyyy")
    End If
    
    If objPedidoCompra.dtDataRefFluxo <> DATA_NULA Then
        DataRefFluxo.Caption = Format(objPedidoCompra.dtDataRefFluxo, "dd/mm/yyyy")
    End If
    
    
    'lê a observacao
    If objPedidoCompra.lObservacao > 0 Then

        objObservacao.lNumInt = objPedidoCompra.lObservacao

        lErro = CF("Observacao_Le", objObservacao)
        If lErro <> SUCESSO And lErro <> 53827 Then gError 68423
        If lErro = 53827 Then gError 68424

        Observ.Caption = objObservacao.sObservacao

    End If
    LabelObsEmbalagem.Caption = objPedidoCompra.sObsEmbalagem

    If objPedidoCompra.dValorFrete > 0 Then ValorFrete.Caption = Format(objPedidoCompra.dValorFrete, "standard")
    If objPedidoCompra.dValorSeguro > 0 Then ValorSeguro.Caption = Format(objPedidoCompra.dValorSeguro, "standard")
    If objPedidoCompra.dOutrasDespesas > 0 Then OutrasDespesas.Caption = Format(objPedidoCompra.dOutrasDespesas, "standard")
    If objPedidoCompra.dValorDesconto > 0 Then DescontoValor.Caption = Format(objPedidoCompra.dValorDesconto, "standard")
    If objPedidoCompra.dValorIPI > 0 Then IPIValor.Caption = Format(objPedidoCompra.dValorIPI, "standard")

    'preenche o Grid com os Ítens do Pedido Compra
    lErro = Preenche_Grid_Itens(objPedidoCompra)
    If lErro <> SUCESSO Then gError 68425

    ' preenche o Grid de distribuicao atraves do objPedidoCompra
    lErro = Preenche_Grid_Distribuicao(objPedidoCompra)
    If lErro <> SUCESSO Then gError 68426

    lErro = Preenche_Grid_Bloqueio(objPedidoCompra)
    If lErro <> SUCESSO Then gError 86159

    'por leo
    'preenche o Grid com as Notas do Pedido Compra
    lErro = Preenche_Grid_Notas(objPedidoCompra)
    If lErro <> SUCESSO Then gError 103358
    
    'preenche o campo ValorTotal e ValorProdutos
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 68427

    iAlterado = 0

    Traz_PedidoCompra_Tela = SUCESSO

    Exit Function

Erro_Traz_PedidoCompra_Tela:

    Traz_PedidoCompra_Tela = gErr

    Select Case gErr

        Case 68429, 68407, 68409, 68411, 68413, 68415, 68421, 68425, 68426, 68427, _
             68423, 68417, 68419, 68468, 68470, 68471, 86159, 86160, 103390, 103391, _
             103355, 103358

        Case 68430
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)

        Case 68412
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case 68414
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO_2", gErr)

        Case 68410, 68416
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_FORNECEDOR_INEXISTENTE", gErr, objFilialFornecedor.lCodFornecedor, objFilialFornecedor.iCodFilial)

        Case 68422
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)

        Case 68424
            Call Rotina_Erro(vbOKOnly, "ERRO_OBSERVACAO_NAO_CADASTRADA", gErr, objPedidoCompra.lObservacao)

        Case 68418
            Call Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_NAO_CADASTRADO", gErr)

        Case 68420
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", gErr, objTransportadora.sNomeReduzido)

        Case 68472
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", gErr, objUsuarios.sCodUsuario)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164341)

    End Select

    Exit Function

End Function

Private Function Preenche_Grid_Itens(objPedidoCompra As ClassPedidoCompras) As Long
'Preenche o GridItens

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoEnxuto As String
Dim sProdutoMascarado As String
Dim dPercDesc As Double
Dim iItem As Integer
Dim objItemPC As New ClassItemPedCompra
Dim dPrecoTotal As Double
Dim sProdutoFormatado As String, objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim objObservacao As New ClassObservacao

On Error GoTo Erro_Preenche_Grid_Itens

    Set gcolItemPedido = New Collection
    Set gobjPC = objPedidoCompra

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridItens)

    iIndice = 0

    For Each objItemPC In objPedidoCompra.colItens

        iIndice = iIndice + 1

        lErro = Mascara_RetornaProdutoEnxuto(objItemPC.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 68501
        
        'Mascara o produto enxuto
        Produto.PromptInclude = False
        Produto.Text = sProdutoMascarado
        Produto.PromptInclude = True

        'Calcula o percentual de desconto
        '#######################################################################
        'ALTERADO POR WAGNER
        If (objItemPC.dPrecoUnitario * objItemPC.dQuantidade) <> 0 Then
            dPercDesc = objItemPC.dValorDesconto / (objItemPC.dPrecoUnitario * objItemPC.dQuantidade)
        Else
            dPercDesc = 0
        End If
        '#######################################################################
        
        
        dPrecoTotal = (objItemPC.dPrecoUnitario * objItemPC.dQuantidade) - objItemPC.dValorDesconto 'Alterado por Wagner

        'Coloca os dados dos itens na tela
        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text

        GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col) = objItemPC.sDescProduto
        GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItemPC.sUM
        If objItemPC.dQuantRecebida > 0 Then GridItens.TextMatrix(iIndice, iGrid_QuantRecebida_Col) = Formata_Estoque(objItemPC.dQuantRecebida)
        If objItemPC.dQuantidade > 0 Then GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemPC.dQuantidade)
        If objItemPC.dPrecoUnitario > 0 Then GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col) = Format(objItemPC.dPrecoUnitario, gobjCOM.sFormatoPrecoUnitario) ' "STANDARD") 'Alterado por Wagner
        If dPercDesc > 0 Then GridItens.TextMatrix(iIndice, iGrid_PercentDesc_Col) = Format(dPercDesc, "Percent")
        If objItemPC.dValorDesconto > 0 Then GridItens.TextMatrix(iIndice, iGrid_Desconto_Col) = Format(objItemPC.dValorDesconto, "Standard")
        If dPrecoTotal > 0 Then GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col) = Format(dPrecoTotal, PrecoTotal.Format) 'Alterado por Wagner
        
        'If objItemPC.dtDataLimite <> DATA_NULA Then GridItens.TextMatrix(iIndice, iGrid_DataLimite_Col) = Format(objItemPC.dtDataLimite, "dd/mm/yyyy")

        If objItemPC.dtDataLimite <> DATA_NULA Then
            GridItens.TextMatrix(iIndice, iGrid_DataLimite_Col) = Format(objItemPC.dtDataLimite, "dd/mm/yyyy")
            
            If gobjCOM.iPCExibeDeliveryDate = MARCADO Then
                
                If objItemPC.dtDeliveryDate = DATA_NULA Then objItemPC.dtDeliveryDate = objItemPC.dtDataLimite
            
                GridItens.TextMatrix(iIndice, iGrid_DeliveryDate_Col) = Format(objItemPC.dtDeliveryDate, "dd/mm/yyyy")
                
                TempoTransito.PromptInclude = False
                TempoTransito.Text = CStr(objItemPC.iTempoTransito)
                TempoTransito.PromptInclude = True
                
                GridItens.TextMatrix(iIndice, iGrid_TempoTransito_Col) = TempoTransito.Text
            
            End If
            
        End If
        
        If objItemPC.dPercentMaisReceb > 0 Then GridItens.TextMatrix(iIndice, iGrid_PercentMaisReceb_Col) = Format(objItemPC.dPercentMaisReceb, "Percent")
        If objItemPC.dAliquotaIPI > 0 Then GridItens.TextMatrix(iIndice, iGrid_AliquotaIPI_Col) = Format(objItemPC.dAliquotaIPI, "Percent")
        If objItemPC.dAliquotaICMS > 0 Then GridItens.TextMatrix(iIndice, iGrid_AliquotaICMS_Col) = Format(objItemPC.dAliquotaICMS, "Percent")

        'lê a observacao
        If objItemPC.lObservacao > 0 Then

            objObservacao.lNumInt = objItemPC.lObservacao

            lErro = CF("Observacao_Le", objObservacao)
            If lErro <> SUCESSO And lErro <> 53827 Then gError 68473
            If lErro = 53827 Then gError 68474

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
    
    If objPedidoCompra.iMoeda = MOEDA_REAL Then
        Call ComparativoMoedaReal_Calcula(1)
    ElseIf Len(Trim(Taxa.Caption)) > 0 Then
        Call ComparativoMoedaReal_Calcula(CDbl(Taxa.Caption))
    End If

    Exit Function

Erro_Preenche_Grid_Itens:

    Preenche_Grid_Itens = gErr

    Select Case gErr

        Case 56437
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)
        
        Case 68501, 68473, 56381
            'Erros tratados nas rotinas chamadas
        
        Case 68474
            Call Rotina_Erro(vbOKOnly, "ERRO_OBSERVACAO_NAO_CADASTRADA", gErr, objPedidoCompra.lObservacao)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164342)

    End Select

    Exit Function

End Function

Private Function Preenche_Grid_NF(colItemNF As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objItemNF As New ClassItemNF
Dim lNumNota As Long
Dim iLinha As Integer

On Error GoTo Erro_Preenche_Grid_NF
    
    'Limpa o GridNFs
    Call Grid_Limpa(objGridNF)
    
    iIndice = 0
    
    For iIndice = 1 To colItemNF.Count
        
        Set objItemNF = colItemNF.Item(iIndice)
        iLinha = objGridNF.iLinhasExistentes + 1
        
'        'Verifica se o número da nota fiscal é o mesmo
'        Do While objItemNF.lNumNFOrig = lNumNota And iIndice < colItemNF.Count
'            iIndice = iIndice + 1
'            Set objItemNF = colItemNF.Item(iIndice)
'        Loop
        
        GridNFs.TextMatrix(iLinha, iGrid_Serie_Col) = objItemNF.sSerieNFOrig
        GridNFs.TextMatrix(iLinha, iGrid_NFiscal_Col) = objItemNF.lNumNFOrig
        GridNFs.TextMatrix(iLinha, iGrid_QuantNF_Col) = Formata_Estoque(objItemNF.dQuantidade)
        GridNFs.TextMatrix(iLinha, iGrid_UMNF_Col) = objItemNF.sUnidadeMed
        GridNFs.TextMatrix(iLinha, iGrid_ItemNF_Col) = objItemNF.iItem
        
        'Atualiza o número de linhas existentes do grid
        objGridNF.iLinhasExistentes = iLinha
        
        lNumNota = objItemNF.lNumNFOrig
        
    Next
    
    Exit Function

Erro_Preenche_Grid_NF:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164343)

    End Select

    Exit Function

End Function

Private Function Preenche_Grid_Distribuicao(objPedidoCompra As ClassPedidoCompras) As Long
'Preenche o GridDistribuicao

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

            If Len(Trim(objLocalizacao.sCcl)) > 0 Then
                lErro = Mascara_RetornaCclEnxuta(objLocalizacao.sCcl, sCclMascarado)
                If lErro <> SUCESSO Then gError 68475
                
                CentroCusto.PromptInclude = False
                CentroCusto.Text = sCclMascarado
                CentroCusto.PromptInclude = True

                sCclMascarado = CentroCusto.Text

                GridDistribuicao.TextMatrix(iIndice, iGrid_CentroCusto_Col) = sCclMascarado

            End If

            If objLocalizacao.sContaContabil <> "" Then

                lErro = Mascara_RetornaContaEnxuta(objLocalizacao.sContaContabil, sContaMascarada)
                If lErro <> SUCESSO Then gError 68476

                ContaContabil.PromptInclude = False
                ContaContabil.Text = sContaMascarada
                ContaContabil.PromptInclude = True
                
                sContaMascarada = ContaContabil.Text
                
                GridDistribuicao.TextMatrix(iIndice, iGrid_ContaContabil_Col) = sContaMascarada

            End If

            If (objLocalizacao.iAlmoxarifado) > 0 Then
                objAlmoxarifado.iCodigo = objLocalizacao.iAlmoxarifado

                lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                If lErro <> SUCESSO Then gError 68477
                GridDistribuicao.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido
            End If
        
        objGridDistribuicao.iLinhasExistentes = iIndice
        
        Next

    Next

    Preenche_Grid_Distribuicao = SUCESSO

    Exit Function

Erro_Preenche_Grid_Distribuicao:

    Preenche_Grid_Distribuicao = gErr

    Select Case gErr

        Case 68475, 68476, 68477

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164344)

    End Select

    Exit Function

End Function

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

    Set objItemPC = New ClassItemPedCompra
    lNumInt = gcolItemPedido.Item(1)
    
    objItemPC.lNumIntDoc = lNumInt
    
    objPedidoCompra.lCodigo = StrParaLong(Codigo.Caption)
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164345)

    End Select

    Exit Sub

End Sub

Private Sub CodigoLabel_Click()

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras
Dim colSelecao As New Collection

On Error GoTo Erro_CodigoLabel_Click

    'Move os dados da tela
    lErro = Move_Tela_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then gError 68437
    
    'Chama a Tela de browse
    Call Chama_Tela("PedComprasTodosLista", colSelecao, objPedidoCompra, objEventoCodigo)

    Exit Sub

Erro_CodigoLabel_Click:

    Select Case gErr

        Case 68437
            'Erro tratado na rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164346)

    End Select

    Exit Sub

End Sub

Private Sub GridItens_Click()

'Dim sProdutoFormatado As String
'Dim iProdutoPreenchido As Integer
'Dim objPedidoCompra As New ClassPedidoCompras
'Dim objItemNFItemPC As New ClassItemNFItemPC
'Dim objItemPC As New ClassItemPedCompra
'Dim objNFiscal As New ClassNFiscal
'Dim colItemNF As New Collection
'Dim lNumIntDoc As Long
Dim lErro As Long

On Error GoTo Erro_GridItens_Click

'    If objGridItens.iLinhasExistentes = 0 Or GridItens.Row = 0 Then gError 89437
'
'    'Verifica se a linha está preenchida
'    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) > 0 Then
'
'        'Coloca o Produto no formato do BD
'        lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
'        If lErro <> SUCESSO Then gError 68453
'
'        objPedidoCompra.iFilialEmpresa = giFilialEmpresa
'        objPedidoCompra.lCodigo = StrParaLong(Codigo.Caption)
'
'        'Lê o Pedido de Compra
'        lErro = CF("PedidoCompra_Le_Todos", objPedidoCompra)
'        If lErro <> SUCESSO And lErro <> 68486 Then gError 68454
'
'        'Se não encontrou ==> erro
'        If lErro = 68486 Then gError 68467
'
'        'Lê os Itens do Pedido de Compra
'        lErro = CF("ItensPC_LeTodos", objPedidoCompra)
'        If lErro <> SUCESSO Then gError 68455
'
'        For Each objItemPC In objPedidoCompra.colItens
'
'            If objItemPC.sProduto = sProdutoFormatado Then
'
'                lNumIntDoc = objItemPC.lNumIntDoc
'                Exit For
'            End If
'        Next
'
'        'Busca no BD os Itens de Notas fiscais associados ao Pedido de Compra
'        lErro = CF("ItemNFItemPC_Le", objItemPC.lNumIntDoc, colItemNF)
'        If lErro <> SUCESSO And lErro <> 66711 Then gError 68456
'
'        'preenche o GridNF
'        lErro = Preenche_Grid_NF(colItemNF)
'        If lErro <> SUCESSO Then gError 68428
'
'    End If

    Exit Sub

Erro_GridItens_Click:

    Select Case gErr

'        Case 68453, 68454, 68455, 68456, 68428
'            'Erros tratados nas rotinas chamadas
'
'        Case 68467
'            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", Err, objPedidoCompra.lCodigo)
'
'        Case 89437
'            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
'
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164347)
            
    End Select

    Exit Sub

End Sub

Private Sub GridItens_EnterCell()

    Call Grid_Entrada_Celula(objGridItens, iAlterado)

End Sub

Private Sub GridItens_GotFocus()

    Call Grid_Recebe_Foco(objGridItens)

End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridItens)

End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridItens)

End Sub

Private Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGridItens)
    
    If iLinhaAnt <> GridItens.Row Then
        Call Trata_NF
    End If
    
    iLinhaAnt = GridItens.Row

End Sub

Private Sub GridItens_Scroll()

    Call Grid_Scroll(objGridItens)

End Sub

Private Sub GridNFs_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridNF, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridNF, iAlterado)
    End If

End Sub

Private Sub GridNFs_EnterCell()

    Call Grid_Entrada_Celula(objGridNF, iAlterado)

End Sub

Private Sub GridNFs_GotFocus()

    Call Grid_Recebe_Foco(objGridNF)

End Sub

Private Sub GridNFs_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridNF)

End Sub

Private Sub NFs_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridNF, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridNF, iAlterado)
    End If

End Sub

Private Sub GridNFs_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridNF)

End Sub

Private Sub GridNFs_RowColChange()

    Call Grid_RowColChange(objGridNF)

End Sub

Private Sub GridNFs_Scroll()

    Call Grid_Scroll(objGridNF)

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

Private Sub GridDistribuicao_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridDistribuicao)

End Sub

Private Sub GridDistribuicao_RowColChange()

    Call Grid_RowColChange(objGridDistribuicao)

End Sub

Private Sub GridDistribuicao_Scroll()

    Call Grid_Scroll(objGridDistribuicao)

End Sub



Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objPedidoCompra = obj1

    'Traz o Pedido de Compra para tela
    lErro = Traz_PedidoCompra_Tela(objPedidoCompra)
    If lErro <> SUCESSO Then gError 68438

    iAlterado = 0

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 68438
            'Erro tratado na rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164348)

    End Select

    Exit Sub

End Sub


Private Function Move_Tela_Memoria(objPedidoCompra As ClassPedidoCompras) As Long
'Recolhe os dados da tela e armazena em objPedidoCompra

Dim lErro As Long
Dim objFornecedor As ClassFornecedor
Dim objComprador As New ClassComprador

On Error GoTo Erro_Move_Tela_Memoria

    'guarda a FilialEmpresa e o Codigo em objPedidoCompra
    objPedidoCompra.lCodigo = StrParaLong(Codigo.Caption)
    objPedidoCompra.iFilialEmpresa = giFilialEmpresa

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164349)

    End Select

    Exit Function

End Function

Public Function Trata_Parametros(Optional objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long
Dim lCodigo As Long
Dim bAchou As Boolean

On Error GoTo Erro_Trata_Parametros

    bAchou = False
    'Verifica se algum pedido foi passada por parametro
    If Not (objPedidoCompra Is Nothing) Then

        If objPedidoCompra.lNumIntDoc > 0 Then

            'Le o Pedido de Compra
            lErro = CF("PedidoCompras_Le", objPedidoCompra)
            If lErro <> SUCESSO And lErro <> 56118 Then gError 68439

            If lErro <> SUCESSO Then
                
                'Le o Pedido de Compra
                lErro = CF("PedidoCompraBaixado_Le", objPedidoCompra)
                If lErro <> SUCESSO And lErro <> 89237 Then gError 68439
            
            Else
                
                objPedidoCompra.dtDataBaixa = DATA_NULA
                objPedidoCompra.iTipoBaixa = 0
                objPedidoCompra.sMotivoBaixa = ""
                
            End If

            If lErro = SUCESSO Then bAchou = True
            
        End If
        If objPedidoCompra.lCodigo > 0 Then
            
            lErro = CF("PedidoCompra_Le_Numero", objPedidoCompra)
            If lErro <> SUCESSO And lErro <> 56142 Then Error 62646
            
            If lErro <> SUCESSO Then
                        
                lErro = CF("PedidoCompraBaixado_Le_Numero", objPedidoCompra)
                If lErro <> SUCESSO And lErro <> 56137 Then gError 68439

            End If

            If lErro = SUCESSO Then bAchou = True

        End If

        'Se o Pedido não existe ==> erro
        If Not bAchou Then gError 68440

        'Traz o Pedido de Compra para tela
        lErro = Traz_PedidoCompra_Tela(objPedidoCompra)
        If lErro <> SUCESSO Then gError 68441


    'SE não há Pedido passado como parametro
    Else

        iChamaTela = 1
        
    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 68439, 68441
            'Erros tratados nas rotinas chamadas
            
        Case 68440
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", gErr, objPedidoCompra.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164350)

    End Select

    Exit Function

End Function

Private Sub TabStrip1_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

       If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual invisivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index

    End If

End Sub

Private Sub TipoDestino_Click(Index As Integer)

Dim lErro As Long

On Error GoTo Erro_TipoDestino_Click

    If Index = iFrameTipoDestinoAtual Then Exit Sub

    'Torna Frame atual invisivel
    FrameTipo(iFrameTipoDestinoAtual).Visible = False

    'Torna Frame correspondente a Index visivel
    FrameTipo(Index).Visible = True

    'Armazena novo valor de iFrameTipoDestinoAtual
    iFrameTipoDestinoAtual = Index

    Call Limpa_Frame_Endereco

    Exit Sub

Erro_TipoDestino_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164351)

    End Select

    Exit Sub

End Sub

Private Sub BotaoImprimir_Click()
'Imprime o Pedido de Compra

Dim lErro As Long, objBloqueioPC As ClassBloqueioPC
Dim objPedidoCompra As New ClassPedidoCompras
Dim objRelatorio As New AdmRelatorio, sFiltro As String

On Error GoTo Erro_BotaoImprimir_Click

    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then gError 68450

    If objPedidoCompra.lCodigo = 0 Then gError 76052

    lErro = CF("PedidoCompra_Le_Todos", objPedidoCompra)
    If lErro <> SUCESSO And lErro <> 68486 Then gError 76033
    
    'Se o Pedido não existe ==> erro
    If lErro = 68486 Then gError 76034
        
    lErro = CF("BloqueiosPC_Le", objPedidoCompra)
    If lErro <> SUCESSO Then gError 76058
    
    For Each objBloqueioPC In objPedidoCompra.colBloqueiosPC
            
        If objBloqueioPC.dtDataLib = DATA_NULA Then gError 76051
    
    Next
    
    'Preenche a Data de Entrada com a Data Atual
    DataEmissao.Caption = Format(gdtDataHoje, "dd/mm/yyyy")

    objPedidoCompra.dtDataEmissao = gdtDataHoje

    'Verifica se o Pedido de Compra está baixado
    If objPedidoCompra.dtDataBaixa <> DATA_NULA Then
    
        lErro = CF("PedidoCompraBaixado_Atualiza_DataEmissao", objPedidoCompra)
        If lErro <> SUCESSO And lErro <> 76070 Then gError 76064
        
        If lErro = 76070 Then gError 76074
        
    'Se o Pedido de Compra não está baixado
    Else
    
        'Atualiza data de emissao no BD para a data atual
        lErro = CF("PedidoCompra_Atualiza_DataEmissao", objPedidoCompra)
        If lErro <> SUCESSO And lErro <> 56348 Then gError 68451

        'se nao encontrar ---> erro
        If lErro = 56348 Then gError 68452
    
    End If
    
    sFiltro = "REL_PCOM.PC_NumIntDoc = @NPEDCOM"
    lErro = CF("Relatorio_ObterFiltro", "Pedido de Compra Consulta", sFiltro)
    If lErro <> SUCESSO Then gError 76035
    
    'Executa o relatório
    lErro = objRelatorio.ExecutarDireto("Pedido de Compra Consulta", sFiltro, 0, "PEDCOM", "NPEDCOM", objPedidoCompra.lNumIntDoc)
    If lErro <> SUCESSO Then gError 76035
    
    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case 68450, 68451, 76058
            'Erros tratados nas rotinas chamadas
            
        Case 68452, 76034, 76074
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", gErr, objPedidoCompra.lCodigo)

        Case 76033, 76035, 76064
        
        Case 76051
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_BLOQUEADO", gErr, objPedidoCompra.lCodigo)
            
        Case 76052
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDCOMPRA_IMPRESSAO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164352)

    End Select

    Exit Sub

End Sub


Private Sub BotaoLimpar_Click()
'Limpa a tela

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 68480

    'Limpa a tela
    Call Limpa_Tela_PedidoComprasCons

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 68480
            'Erro tratado na rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164353)

    End Select

    Exit Sub

End Sub

Sub Limpa_Tela_PedidoComprasCons()

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_PedidoComprasCons

    'Limpa a tela
    Call Limpa_Tela(Me)

    'Limpa os outros campos da tela
    Codigo.Caption = ""
    Contato.Caption = ""
    Data.Caption = ""
    Fornecedor.Caption = ""
    Observ.Caption = ""
    DataEnvio.Caption = ""
    DataRefFluxo.Caption = ""
    DataAlteracao.Caption = ""
    DataBaixa.Caption = ""
    DataEmissao.Caption = ""
    Filial.Caption = ""
    CondPagto.Caption = ""
    TabelaPreco.Caption = ""
    ValorTotal.Caption = ""
    ValorFrete.Caption = ""
    ValorSeguro.Caption = ""
    ValorProdutos.Caption = ""
    OutrasDespesas.Caption = ""
    DescontoValor.Caption = ""
    IPIValor.Caption = ""
    FilialEmpresa.Caption = ""
    TipoFrete.Caption = ""
    Transportadora.Caption = ""
    Fornec.Caption = ""
    FilialFornec.Caption = ""
'leo
    Endereco.Caption = ""
    Taxa.Caption = ""
    Moeda.Caption = ""
    
    iLinhaAnt = 0
    
    Call Limpa_Frame_Endereco

    'Limpa os grids
    Call Grid_Limpa(objGridItens)
    Call Grid_Limpa(objGridDistribuicao)
    Call Grid_Limpa(objGridNF)
    Call Grid_Limpa(objGridNotas)

    Set gcolItemPedido = New Collection
    Set gobjPC = New ClassPedidoCompras

    Exit Sub

Erro_Limpa_Tela_PedidoComprasCons:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164354)

    End Select

    Exit Sub

End Sub
'ja existe em PedidoCompra
Private Sub Limpa_Frame_Endereco()

    Endereco.Caption = ""
    Bairro.Caption = ""
    Cidade.Caption = ""
    CEP.Caption = ""
    Estado.Caption = ""
    Pais.Caption = ""

    Exit Sub

End Sub


Private Sub Preenche_Endereco(objEndereco As ClassEndereco)

Dim objPais As New ClassPais
Dim lErro As Long

On Error GoTo Erro_Preenche_Endereco

    objPais.iCodigo = objEndereco.iCodigoPais

    lErro = CF("Paises_Le", objPais)
    If lErro <> SUCESSO And lErro <> 47876 Then gError 68481
    If lErro = 47876 Then gError 68482

    Endereco.Caption = objEndereco.sEndereco
    Bairro.Caption = objEndereco.sBairro
    Estado.Caption = objEndereco.sSiglaEstado
    Cidade.Caption = objEndereco.sCidade
    Pais.Caption = objPais.sNome
    CEP.Caption = objEndereco.sCEP

    Exit Sub

Erro_Preenche_Endereco:

    Select Case gErr

        Case 68481
            'Erro tratado na rotina chamada
            
        Case 68482
            Call Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_CADASTRADO", gErr, objEndereco.iCodigoPais)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164355)

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
    ValorProdutos.Caption = Format(dPrecoTotal, PrecoTotal.Format) 'Altreado por Wagner
    dValorTotal = (dPrecoTotal + StrParaDbl(ValorFrete.Caption) + StrParaDbl(ValorSeguro.Caption) + StrParaDbl(OutrasDespesas.Caption) + StrParaDbl(IPIValor.Caption)) - StrParaDbl(DescontoValor.Caption)

    'Coloca na tela o valor total
    ValorTotal.Caption = Format(dValorTotal, PrecoTotal.Format) 'Altreado por Wagner

    ValorTotal_Calcula = SUCESSO

    Exit Function

Erro_ValorTotal_Calcula:

    ValorTotal_Calcula = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164356)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    'libera as variaveis globais
    Set objEventoCodigo = Nothing

    Set objGridItens = Nothing
    Set objGridDistribuicao = Nothing
    Set objGridNF = Nothing
    Set objGridBloqueio = Nothing
    Set objGridNotas = Nothing

    Set gcolItemPedido = Nothing
    Set gobjPC = Nothing

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164357)

    End Select

    Exit Sub

End Sub



'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Pedido de Compras - Consulta"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "PedComprasCons"

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


Private Sub FornecedorLabel_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(FornecedorLabel(Index), Source, X, Y)
End Sub

Private Sub FornecedorLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel(Index), Button, Shift, X, Y)
End Sub


Private Sub Data_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Data, Source, X, Y)
End Sub

Private Sub Data_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Data, Button, Shift, X, Y)
End Sub

Private Sub Codigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Codigo, Source, X, Y)
End Sub

Private Sub Codigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Codigo, Button, Shift, X, Y)
End Sub

Private Sub CondPagto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CondPagto, Source, X, Y)
End Sub

Private Sub CondPagto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CondPagto, Button, Shift, X, Y)
End Sub

Private Sub Contato_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Contato, Source, X, Y)
End Sub

Private Sub Contato_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Contato, Button, Shift, X, Y)
End Sub

Private Sub DataAlteracao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataAlteracao, Source, X, Y)
End Sub

Private Sub DataAlteracao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataAlteracao, Button, Shift, X, Y)
End Sub

Private Sub DataEmissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataEmissao, Source, X, Y)
End Sub

Private Sub DataEmissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataEmissao, Button, Shift, X, Y)
End Sub

Private Sub Fornecedor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Fornecedor, Source, X, Y)
End Sub

Private Sub Fornecedor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Fornecedor, Button, Shift, X, Y)
End Sub

Private Sub Filial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Filial, Source, X, Y)
End Sub

Private Sub Filial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Filial, Button, Shift, X, Y)
End Sub

Private Sub Comprador_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Comprador, Source, X, Y)
End Sub

Private Sub Comprador_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Comprador, Button, Shift, X, Y)
End Sub

Private Sub CondPagtoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CondPagtoLabel, Source, X, Y)
End Sub

Private Sub CondPagtoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CondPagtoLabel, Button, Shift, X, Y)
End Sub

Private Sub CodigoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CodigoLabel, Source, X, Y)
End Sub

Private Sub CodigoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CodigoLabel, Button, Shift, X, Y)
End Sub

Private Sub DataEnvio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataEnvio, Source, X, Y)
End Sub

Private Sub DataEnvio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataEnvio, Button, Shift, X, Y)
End Sub

Private Sub Observ_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Observ, Source, X, Y)
End Sub

Private Sub Observ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Observ, Button, Shift, X, Y)
End Sub

Private Sub ValorTotal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorTotal, Source, X, Y)
End Sub

Private Sub ValorTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorTotal, Button, Shift, X, Y)
End Sub

Private Sub IPIValor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(IPIValor, Source, X, Y)
End Sub

Private Sub IPIValor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(IPIValor, Button, Shift, X, Y)
End Sub

Private Sub OutrasDespesas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(OutrasDespesas, Source, X, Y)
End Sub

Private Sub OutrasDespesas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(OutrasDespesas, Button, Shift, X, Y)
End Sub

Private Sub ValorSeguro_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorSeguro, Source, X, Y)
End Sub

Private Sub ValorSeguro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorSeguro, Button, Shift, X, Y)
End Sub

Private Sub ValorFrete_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorFrete, Source, X, Y)
End Sub

Private Sub ValorFrete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorFrete, Button, Shift, X, Y)
End Sub

Private Sub ValorProdutos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorProdutos, Source, X, Y)
End Sub

Private Sub ValorProdutos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorProdutos, Button, Shift, X, Y)
End Sub

Private Sub DescontoValor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescontoValor, Source, X, Y)
End Sub

Private Sub DescontoValor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescontoValor, Button, Shift, X, Y)
End Sub

Private Sub TransportadoraLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TransportadoraLabel, Source, X, Y)
End Sub

Private Sub TransportadoraLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TransportadoraLabel, Button, Shift, X, Y)
End Sub

Private Sub TipoFrete_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoFrete, Source, X, Y)
End Sub

Private Sub TipoFrete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoFrete, Button, Shift, X, Y)
End Sub

Private Sub Transportadora_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Transportadora, Source, X, Y)
End Sub

Private Sub Transportadora_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Transportadora, Button, Shift, X, Y)
End Sub

Private Sub Fornec_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Fornec, Source, X, Y)
End Sub

Private Sub Fornec_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Fornec, Button, Shift, X, Y)
End Sub

Private Sub FilialFornec_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialFornec, Source, X, Y)
End Sub

Private Sub FilialFornec_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialFornec, Button, Shift, X, Y)
End Sub

Private Sub FilialEmpresa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialEmpresa, Source, X, Y)
End Sub

Private Sub FilialEmpresa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialEmpresa, Button, Shift, X, Y)
End Sub

Private Sub Pais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Pais, Source, X, Y)
End Sub

Private Sub Pais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Pais, Button, Shift, X, Y)
End Sub

Private Sub Estado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Estado, Source, X, Y)
End Sub

Private Sub Estado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Estado, Button, Shift, X, Y)
End Sub

Private Sub CEP_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CEP, Source, X, Y)
End Sub

Private Sub CEP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CEP, Button, Shift, X, Y)
End Sub

Private Sub Cidade_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Cidade, Source, X, Y)
End Sub

Private Sub Cidade_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Cidade, Button, Shift, X, Y)
End Sub

Private Sub Bairro_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Bairro, Source, X, Y)
End Sub

Private Sub Bairro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Bairro, Button, Shift, X, Y)
End Sub

Private Sub Endereco_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Endereco, Source, X, Y)
End Sub

Private Sub Endereco_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Endereco, Button, Shift, X, Y)
End Sub

Private Function Carrega_RecebForaFaixa() As Long

Dim lErro As Long

On Error GoTo Erro_Carrega_RecebForaFaixa

    'Limpa a combo
    RecebForaFaixa.Clear

    RecebForaFaixa.AddItem MENSAGEM_NAO_AVISA_ACEITA_RECEBIMENTO
    RecebForaFaixa.ItemData(RecebForaFaixa.NewIndex) = NAO_AVISA_E_ACEITA_RECEBIMENTO

    RecebForaFaixa.AddItem MENSAGEM_REJEITA_RECEBIMENTO
    RecebForaFaixa.ItemData(RecebForaFaixa.NewIndex) = ERRO_E_REJEITA_RECEBIMENTO

    RecebForaFaixa.AddItem MENSAGEM_ACEITA_RECEBIMENTO
    RecebForaFaixa.ItemData(RecebForaFaixa.NewIndex) = AVISA_E_ACEITA_RECEBIMENTO

    Exit Function

Erro_Carrega_RecebForaFaixa:

    Carrega_RecebForaFaixa = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164358)

    End Select

    Exit Function

End Function



Private Sub Label4_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label4(Index), Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4(Index), Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label6(Index), Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6(Index), Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label15(Index), Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15(Index), Button, Shift, X, Y)
End Sub


Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub DataBaixa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataBaixa, Source, X, Y)
End Sub

Private Sub DataBaixa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataBaixa, Button, Shift, X, Y)
End Sub

Private Sub Label30_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label30, Source, X, Y)
End Sub

Private Sub Label30_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label30, Button, Shift, X, Y)
End Sub

Private Sub Label29_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label29, Source, X, Y)
End Sub

Private Sub Label29_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label29, Button, Shift, X, Y)
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

Private Sub Label21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label21, Source, X, Y)
End Sub

Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label21, Button, Shift, X, Y)
End Sub

Private Sub Label28_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label28, Source, X, Y)
End Sub

Private Sub Label28_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label28, Button, Shift, X, Y)
End Sub

Private Sub Label41_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label41, Source, X, Y)
End Sub

Private Sub Label41_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label41, Button, Shift, X, Y)
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

Private Sub Label25_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label25, Source, X, Y)
End Sub

Private Sub Label25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label25, Button, Shift, X, Y)
End Sub

Private Sub Label37_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label37, Source, X, Y)
End Sub

Private Sub Label37_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label37, Button, Shift, X, Y)
End Sub

Private Sub Label32_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label32, Source, X, Y)
End Sub

Private Sub Label32_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label32, Button, Shift, X, Y)
End Sub

Private Sub Label63_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label63, Source, X, Y)
End Sub

Private Sub Label63_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label63, Button, Shift, X, Y)
End Sub

Private Sub Label65_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label65, Source, X, Y)
End Sub

Private Sub Label65_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label65, Button, Shift, X, Y)
End Sub

Private Sub Label70_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label70, Source, X, Y)
End Sub

Private Sub Label70_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label70, Button, Shift, X, Y)
End Sub

Private Sub Label71_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label71, Source, X, Y)
End Sub

Private Sub Label71_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label71, Button, Shift, X, Y)
End Sub

Private Sub Label72_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label72, Source, X, Y)
End Sub

Private Sub Label72_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label72, Button, Shift, X, Y)
End Sub

Private Sub Label73_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label73, Source, X, Y)
End Sub

Private Sub Label73_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label73, Button, Shift, X, Y)
End Sub

Private Sub Label31_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label31, Source, X, Y)
End Sub

Private Sub Label31_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label31, Button, Shift, X, Y)
End Sub
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

Private Function Preenche_Grid_Bloqueio(objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long
Dim iIndice As Integer
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

        objTipoDeBloqueioPC.iCodigo = objBloqueioPC.iTipoBloqueio

        lErro = CF("TipoDeBloqueioPC_Le", objTipoDeBloqueioPC)
        If lErro <> SUCESSO And lErro <> 49143 Then Error 57250
        If lErro = 49143 Then Error 57251

        GridBloqueios.TextMatrix(iIndice, iGrid_TipoBloqueio_Col) = objBloqueioPC.iTipoBloqueio & SEPARADOR & objTipoDeBloqueioPC.sNomeReduzido

        If objBloqueioPC.dtDataLib <> DATA_NULA Then GridBloqueios.TextMatrix(iIndice, iGrid_DataLiberacao_Col) = Format(objBloqueioPC.dtDataLib, "dd/mm/yyyy")
        If (objBloqueioPC.dtData <> DATA_NULA) Then GridBloqueios.TextMatrix(iIndice, iGrid_DataBloqueio_Col) = Format(objBloqueioPC.dtData, "dd/mm/yyyy")
    Next

    objGridBloqueio.iLinhasExistentes = iIndice

    Preenche_Grid_Bloqueio = SUCESSO

    Exit Function

Erro_Preenche_Grid_Bloqueio:

    Preenche_Grid_Bloqueio = Err

    Select Case Err

        Case 57251
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODEBLOQUEIOPC_NAO_CADASTRADO", Err, objTipoDeBloqueioPC.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164359)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164360)

    End Select

    Exit Function

End Function

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
'    iGrid_NotaPC_Col = 1

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridNotas

    'Linhas do grid
    objGridInt.objGrid.Rows = 20

    GridBloqueios.ColWidth(0) = 300

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 11

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGridInt.iProibidoIncluir = PROIBIDO_INCLUIR
    
    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridNotas = SUCESSO

    Exit Function

End Function

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164361)

    End Select

    Call Comando_Fechar(lComando)

End Function

Private Sub ComparativoMoedaReal_Calcula(ByVal dTaxa As Double)
'Preenche as colunas INFORMATIVAS de proporção da moeda R$.

Dim iIndice As Integer

On Error GoTo Erro_ComparativoMoedaReal_Calcula

    'Para cada linha do grid de Itens será claculado o correspondente em R$
    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        '##################################################################
        'ALTERADO POR WAGNER
        If (StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col)) * dTaxa) > 0 Then
            'Preço Unitário em R$ = Preço Unitário na Moeda selecionada dividido pela taxa de conversão
            GridItens.TextMatrix(iIndice, iGrid_PrecoUnitarioMoedaReal_Col) = Format((StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col)) - (StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col)) / StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col)))) * dTaxa, gobjCOM.sFormatoPrecoUnitario) 'Alterado por Wagner
        Else
            GridItens.TextMatrix(iIndice, iGrid_PrecoUnitarioMoedaReal_Col) = ""
        End If
        
        'Preço Total em R$ = Preço Unitário em R$ x Quantidade do produto
        GridItens.TextMatrix(iIndice, iGrid_TotalMoedaReal_Col) = Format(StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoUnitarioMoedaReal_Col)) * StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col)), TotalMoedaReal.Format) 'Alterado por Wagner
        '##################################################################
        
    Next

    Exit Sub
    
Erro_ComparativoMoedaReal_Calcula:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164362)

    End Select

    Exit Sub

End Sub


Private Function Inicializa_Grid_Notas(objGridInt As AdmGrid) As Long
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
    objGridInt.objGrid.Rows = 30

    GridBloqueios.ColWidth(0) = 300

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 22

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGridInt.iProibidoIncluir = PROIBIDO_INCLUIR
    
    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Notas = SUCESSO

    Exit Function

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164363)

    End Select

    Exit Function

End Function

'##############################################
'Inserido por Wagner
Private Sub Formata_Controles()

    PrecoUnitario.Format = gobjCOM.sFormatoPrecoUnitario
    PrecoUnitarioMoedaReal.Format = gobjCOM.sFormatoPrecoUnitario

End Sub
'##############################################

'##############################################
'Inserido por Wagner
Private Sub BotaoCancelarBaixa_Click()
'Chama a tela de ConcorrenciaCons ou PedCotacaoCons, de acordo com
'a forma de geracao do Pedido de Compra (por Concorrencia ou PedidoCotacao)

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras

On Error GoTo Erro_BotaoCancelarBaixa_Click
   
    objPedidoCompra.lCodigo = StrParaLong(Codigo.Caption)
    objPedidoCompra.iFilialEmpresa = giFilialEmpresa
    
    'Lê o Pedido de Compra
    lErro = CF("PedidoCompra_Le_Todos", objPedidoCompra)
    If lErro <> SUCESSO Then gError 141349
    
    If objPedidoCompra.iStatus <> PEDIDOCOMPRA_STATUS_BAIXADO Then gError 141350
        
    lErro = CF("PedidoCompra_Cancelar_Baixa", objPedidoCompra)
    If lErro <> SUCESSO Then gError 141351
    
    Call Limpa_Tela_PedidoComprasCons
 
    Exit Sub

Erro_BotaoCancelarBaixa_Click:

    Select Case gErr
    
        Case 141349, 141351
        
        Case 141350
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_BAIXADO", gErr, objPedidoCompra.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141871)

    End Select

    Exit Sub

End Sub
'####################################################################

Private Sub Trata_NF()

Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objPedidoCompra As New ClassPedidoCompras
Dim objItemNFItemPC As New ClassItemNFItemPC
Dim objItemPC As New ClassItemPedCompra
Dim objNFiscal As New ClassNFiscal
Dim colItemNF As New Collection
Dim lNumIntDoc As Long
Dim lErro As Long

On Error GoTo Erro_Trata_NF

    If objGridItens.iLinhasExistentes = 0 Or GridItens.Row = 0 Then gError ERRO_SEM_MENSAGEM '89437
    
    'Verifica se a linha está preenchida
    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) > 0 Then

        'Coloca o Produto no formato do BD
        lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 68453

        objPedidoCompra.iFilialEmpresa = giFilialEmpresa
        objPedidoCompra.lCodigo = StrParaLong(Codigo.Caption)

        'Lê o Pedido de Compra
        lErro = CF("PedidoCompra_Le_Todos", objPedidoCompra)
        If lErro <> SUCESSO And lErro <> 68486 Then gError 68454

        'Se não encontrou ==> erro
        If lErro = 68486 Then gError 68467
        
        'Lê os Itens do Pedido de Compra
        lErro = CF("ItensPC_LeTodos", objPedidoCompra)
        If lErro <> SUCESSO Then gError 68455

        For Each objItemPC In objPedidoCompra.colItens

            If objItemPC.sProduto = sProdutoFormatado Then

                lNumIntDoc = objItemPC.lNumIntDoc
                Exit For
            End If
        Next

        'Busca no BD os Itens de Notas fiscais associados ao Pedido de Compra
        lErro = CF("ItemNFItemPC_Le", objItemPC.lNumIntDoc, colItemNF)
        If lErro <> SUCESSO And lErro <> 66711 Then gError 68456
        
        'preenche o GridNF
        lErro = Preenche_Grid_NF(colItemNF)
        If lErro <> SUCESSO Then gError 68428

    End If

    Exit Sub

Erro_Trata_NF:

    Select Case gErr

        Case 68453, 68454, 68455, 68456, 68428
            'Erros tratados nas rotinas chamadas
            
        Case 68467
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", Err, objPedidoCompra.lCodigo)
            
        Case ERRO_SEM_MENSAGEM '89437
            'Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164347)
            
    End Select

    Exit Sub

End Sub

Private Sub BotaoEntrega_Click()

Dim lErro As Long
Dim sProdutoTela As String
Dim dQuantidade As Double
Dim objItemPC As ClassItemPedCompra
Dim sProduto As String, iPreenchido As Integer

On Error GoTo Erro_BotaoEntrega_Click

    If GridItens.Row = 0 Then gError 183202

    sProdutoTela = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)

    dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))

    If Len(sProdutoTela) = 0 Then gError 183203
    
    Set objItemPC = gobjPC.colItens(GridItens.Row)
    
    gobjPC.lCodigo = StrParaLong(Codigo.Caption)
    
    'Formata o produto
    lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    objItemPC.sProduto = sProduto
    objItemPC.sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)
    objItemPC.dQuantidade = dQuantidade

    Call Chama_Tela_Modal("DataEntregaCOM", gobjPC, objItemPC, True)
    
    If giRetornoTela = vbOK Then
        If objItemPC.colDataEntrega.Count > 0 Then
            GridItens.TextMatrix(GridItens.Row, iGrid_DataLimite_Col) = Format(objItemPC.colDataEntrega.Item(1).dtDataEntrega, "dd/mm/yyyy")
        End If
    End If

    Exit Sub

Erro_BotaoEntrega_Click:

    Select Case gErr

        Case 183202
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 183203
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183196)

    End Select

    Exit Sub
    
End Sub
