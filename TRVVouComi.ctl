VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl TRVVouComi 
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   9510
   Begin VB.ComboBox EmiCargo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   255
      TabIndex        =   99
      Top             =   8130
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Frame Frame5 
      Caption         =   "OVER - Comissão Direta (C.D.) e Indireta (C.I.)"
      Height          =   1950
      Left            =   105
      TabIndex        =   88
      Top             =   4425
      Width           =   6000
      Begin MSMask.MaskEdBox OVEREmi 
         Height          =   225
         Left            =   480
         TabIndex        =   94
         Top             =   750
         Width           =   1740
         _ExtentX        =   3069
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
      Begin MSMask.MaskEdBox OverPercTotal 
         Height          =   225
         Left            =   1890
         TabIndex        =   95
         Top             =   765
         Width           =   735
         _ExtentX        =   1296
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
         Format          =   "0%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox OVERPercVou 
         Height          =   225
         Left            =   2475
         TabIndex        =   96
         Top             =   750
         Width           =   735
         _ExtentX        =   1296
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
         Format          =   "0%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox OverValor 
         Height          =   225
         Left            =   3195
         TabIndex        =   97
         Top             =   765
         Width           =   825
         _ExtentX        =   1455
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
         Format          =   "0%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox OverCargo 
         Height          =   225
         Left            =   4110
         TabIndex        =   98
         Top             =   780
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
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridOver 
         Height          =   1260
         Left            =   30
         TabIndex        =   93
         Top             =   240
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   2223
         _Version        =   393216
         Cols            =   8
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         FocusRect       =   2
      End
   End
   Begin VB.CommandButton BotaoHistOcor 
      Caption         =   " Ocorrências"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6465
      TabIndex        =   5
      ToolTipText     =   "Exibe o histórico de ocorrências"
      Top             =   75
      Width           =   1740
   End
   Begin VB.Frame Frame4 
      Caption         =   "Comissão Interna"
      Height          =   1950
      Left            =   6210
      TabIndex        =   81
      Top             =   4425
      Width           =   3180
      Begin VB.Frame Frame7 
         Caption         =   "Outros Vendedores"
         Height          =   1365
         Left            =   90
         TabIndex        =   82
         Top             =   510
         Width           =   2985
         Begin MSMask.MaskEdBox Vendedor 
            Height          =   225
            Left            =   525
            TabIndex        =   15
            Top             =   555
            Width           =   2000
            _ExtentX        =   3519
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
         Begin MSFlexGridLib.MSFlexGrid GridVendedores 
            Height          =   480
            Left            =   45
            TabIndex        =   14
            Top             =   180
            Width           =   2900
            _ExtentX        =   5106
            _ExtentY        =   847
            _Version        =   393216
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
      End
      Begin MSMask.MaskEdBox PromotorNovo 
         Height          =   270
         Left            =   975
         TabIndex        =   13
         Top             =   210
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   476
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin VB.Label LabelPromotor 
         Caption         =   "Promotor:"
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
         Left            =   75
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   83
         Top             =   210
         Width           =   825
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   8325
      ScaleHeight     =   450
      ScaleWidth      =   1005
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "TRVVouComi.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Gravar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   540
         Picture         =   "TRVVouComi.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoHist 
      Caption         =   "Detalhamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4680
      TabIndex        =   4
      Top             =   75
      Width           =   1740
   End
   Begin VB.Frame Frame2 
      Caption         =   "Voucher"
      Height          =   1065
      Left            =   105
      TabIndex        =   26
      Top             =   450
      Width           =   9300
      Begin VB.CheckBox Antc 
         Caption         =   "Antc"
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
         Left            =   8040
         TabIndex        =   66
         Top             =   210
         Width           =   945
      End
      Begin VB.CheckBox Cartao 
         Caption         =   "Cartão"
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
         Left            =   7140
         TabIndex        =   65
         Top             =   210
         Width           =   945
      End
      Begin VB.Label Label1 
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
         Index           =   17
         Left            =   945
         TabIndex        =   101
         Top             =   810
         Width           =   735
      End
      Begin VB.Label AgenciaVou 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   1680
         TabIndex        =   100
         Top             =   765
         Width           =   3330
      End
      Begin VB.Label ValorBase 
         BorderStyle     =   1  'Fixed Single
         Height          =   275
         Left            =   5850
         TabIndex        =   85
         Top             =   150
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Base R$:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   78
         Left            =   5040
         TabIndex        =   84
         Top             =   210
         Width           =   810
      End
      Begin VB.Label Destino 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   1680
         TabIndex        =   63
         Top             =   465
         Width           =   3330
      End
      Begin VB.Label Produto 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   5850
         TabIndex        =   61
         Top             =   465
         Width           =   3345
      End
      Begin VB.Label Label1 
         Caption         =   "Destino:"
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
         Index           =   44
         Left            =   900
         TabIndex        =   64
         Top             =   510
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "Produto:"
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
         Left            =   5100
         TabIndex        =   62
         Top             =   495
         Width           =   1020
      End
      Begin VB.Label DataEmissaoVou 
         BorderStyle     =   1  'Fixed Single
         Height          =   275
         Left            =   1680
         TabIndex        =   29
         Top             =   150
         Width           =   1290
      End
      Begin VB.Label ValorBrutoVou 
         BorderStyle     =   1  'Fixed Single
         Height          =   275
         Left            =   3885
         TabIndex        =   27
         Top             =   150
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Data de Emissão:"
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
         Index           =   5
         Left            =   150
         TabIndex        =   30
         Top             =   210
         Width           =   1620
      End
      Begin VB.Label Label1 
         Caption         =   "Bruto R$:"
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
         Index           =   3
         Left            =   3015
         TabIndex        =   28
         Top             =   210
         Width           =   825
      End
   End
   Begin VB.CommandButton BotaoTrazerVou 
      Height          =   330
      Left            =   4140
      Picture         =   "TRVVouComi.ctx":02D8
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Trazer Dados"
      Top             =   90
      Width           =   360
   End
   Begin VB.Frame FrameSuporte 
      Caption         =   "Simular Comissão"
      Height          =   1680
      Left            =   90
      TabIndex        =   37
      Top             =   9315
      Width           =   9270
      Begin VB.CheckBox Import 
         Caption         =   "Simular importação"
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
         Left            =   2925
         TabIndex        =   20
         Top             =   585
         Width           =   2085
      End
      Begin VB.CommandButton BotaoPrimeiraComissao 
         Caption         =   "Primeira Comissão"
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
         Left            =   165
         TabIndex        =   79
         Top             =   1365
         Width           =   1800
      End
      Begin VB.CheckBox CartaoNovo 
         Caption         =   "Cartão"
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
         Left            =   5865
         TabIndex        =   17
         Top             =   240
         Width           =   945
      End
      Begin VB.CheckBox AntcNovo 
         Caption         =   "Antc"
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
         Left            =   7845
         TabIndex        =   18
         Top             =   225
         Width           =   945
      End
      Begin VB.CommandButton BotaoExcluirComissao 
         Caption         =   "Excluir Comissão"
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
         Left            =   2040
         TabIndex        =   22
         Top             =   1380
         Width           =   1740
      End
      Begin MSMask.MaskEdBox BrutoNovo 
         Height          =   315
         Left            =   1680
         TabIndex        =   19
         Top             =   540
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
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
      Begin MSMask.MaskEdBox ProdutoNovo 
         Height          =   315
         Left            =   1680
         TabIndex        =   16
         Top             =   180
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox DestinoNovo 
         Height          =   315
         Left            =   5865
         TabIndex        =   21
         Top             =   555
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Cliente 
         Height          =   315
         Left            =   5865
         TabIndex        =   78
         Top             =   900
         Visible         =   0   'False
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox BaseNovo 
         Height          =   315
         Left            =   1680
         TabIndex        =   86
         Top             =   915
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
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
      Begin VB.Label Label1 
         Caption         =   "Base:"
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
         Index           =   4
         Left            =   1080
         TabIndex        =   87
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Destino:"
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
         Index           =   22
         Left            =   5055
         TabIndex        =   77
         Top             =   600
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "Produto:"
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
         Index           =   21
         Left            =   870
         TabIndex        =   76
         Top             =   225
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Bruto:"
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
         Index           =   20
         Left            =   1065
         TabIndex        =   75
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Comissão Externa Nova"
      Height          =   1455
      Left            =   105
      TabIndex        =   36
      Top             =   2955
      Width           =   9285
      Begin MSMask.MaskEdBox RepresentanteNovo 
         Height          =   270
         Left            =   1665
         TabIndex        =   7
         Top             =   480
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   476
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox PercComiRepNovo 
         Height          =   270
         Left            =   5850
         TabIndex        =   8
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
      Begin MSMask.MaskEdBox CorrentistaNovo 
         Height          =   270
         Left            =   1665
         TabIndex        =   9
         Top             =   795
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   476
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox PercComiCorNovo 
         Height          =   270
         Left            =   5850
         TabIndex        =   10
         Top             =   795
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox EmissorNovo 
         Height          =   270
         Left            =   1665
         TabIndex        =   11
         Top             =   1110
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   476
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox PercComiEmiNovo 
         Height          =   270
         Left            =   5850
         TabIndex        =   12
         Top             =   1110
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox PercComiAgeNovo 
         Height          =   275
         Left            =   5850
         TabIndex        =   6
         Top             =   165
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         Caption         =   "%CI:"
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
         Index           =   15
         Left            =   8145
         TabIndex        =   92
         Top             =   1140
         Width           =   375
      End
      Begin VB.Label CIPercTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   8535
         TabIndex        =   91
         Top             =   1095
         Width           =   705
      End
      Begin VB.Label Label1 
         Caption         =   "%CD:"
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
         Index           =   13
         Left            =   6960
         TabIndex        =   90
         Top             =   1140
         Width           =   465
      End
      Begin VB.Label CDPercTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   7425
         TabIndex        =   89
         Top             =   1095
         Width           =   705
      End
      Begin VB.Label AgenciaNovo 
         BorderStyle     =   1  'Fixed Single
         Height          =   275
         Left            =   1665
         TabIndex        =   80
         Top             =   165
         Width           =   3000
      End
      Begin VB.Label Label1 
         Caption         =   "% Comissão:"
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
         Index           =   18
         Left            =   4755
         TabIndex        =   74
         Top             =   210
         Width           =   1140
      End
      Begin VB.Label LabelAgencia 
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
         Height          =   330
         Left            =   930
         TabIndex        =   73
         Top             =   225
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "% Comissão:"
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
         Index           =   16
         Left            =   4755
         TabIndex        =   72
         Top             =   510
         Width           =   1140
      End
      Begin VB.Label LabelEmissor 
         Caption         =   "Emissor:"
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
         Left            =   870
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   71
         Top             =   1125
         Width           =   690
      End
      Begin VB.Label Label1 
         Caption         =   "% Comissão:"
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
         Index           =   14
         Left            =   4755
         TabIndex        =   70
         Top             =   810
         Width           =   1140
      End
      Begin VB.Label LabelCorrentista 
         Caption         =   "Correntista:"
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
         Left            =   585
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   69
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "% Comissão:"
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
         Index           =   6
         Left            =   4755
         TabIndex        =   68
         Top             =   1140
         Width           =   1140
      End
      Begin VB.Label LabelRepresentante 
         Caption         =   "Representante:"
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
         Left            =   255
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   67
         Top             =   525
         Width           =   1410
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Comissão Externa Atual"
      Height          =   1440
      Left            =   105
      TabIndex        =   31
      Top             =   1500
      Width           =   9285
      Begin VB.Label VlrComiAge 
         BorderStyle     =   1  'Fixed Single
         Height          =   275
         Left            =   7770
         TabIndex        =   59
         Top             =   195
         Width           =   1425
      End
      Begin VB.Label PercComiAge 
         BorderStyle     =   1  'Fixed Single
         Height          =   275
         Left            =   5850
         TabIndex        =   57
         Top             =   195
         Width           =   1065
      End
      Begin VB.Label Agencia 
         BorderStyle     =   1  'Fixed Single
         Height          =   275
         Left            =   1710
         TabIndex        =   56
         Top             =   195
         Width           =   2925
      End
      Begin VB.Label Label1 
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
         Height          =   240
         Index           =   12
         Left            =   7125
         TabIndex        =   60
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "% Comissão:"
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
         Index           =   11
         Left            =   4785
         TabIndex        =   58
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label1 
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
         Height          =   330
         Index           =   10
         Left            =   960
         TabIndex        =   55
         Top             =   240
         Width           =   750
      End
      Begin VB.Label VlrComiEmi 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   7770
         TabIndex        =   50
         Top             =   1095
         Width           =   1425
      End
      Begin VB.Label VlrComiCor 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   7770
         TabIndex        =   48
         Top             =   795
         Width           =   1425
      End
      Begin VB.Label VlrComiRep 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   7770
         TabIndex        =   46
         Top             =   495
         Width           =   1425
      End
      Begin VB.Label PercComiEmi 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   5850
         TabIndex        =   44
         Top             =   1095
         Width           =   1065
      End
      Begin VB.Label PercComiCor 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   5850
         TabIndex        =   40
         Top             =   795
         Width           =   1065
      End
      Begin VB.Label PercComiRep 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   5850
         TabIndex        =   34
         Top             =   495
         Width           =   1065
      End
      Begin VB.Label Label1 
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
         Height          =   240
         Index           =   9
         Left            =   7125
         TabIndex        =   51
         Top             =   1140
         Width           =   1140
      End
      Begin VB.Label Label1 
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
         Height          =   240
         Index           =   8
         Left            =   7125
         TabIndex        =   49
         Top             =   825
         Width           =   1140
      End
      Begin VB.Label Label1 
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
         Height          =   240
         Index           =   7
         Left            =   7125
         TabIndex        =   47
         Top             =   525
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "% Comissão:"
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
         Index           =   61
         Left            =   4785
         TabIndex        =   45
         Top             =   1140
         Width           =   1140
      End
      Begin VB.Label Emissor 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   1710
         TabIndex        =   43
         Top             =   1095
         Width           =   2925
      End
      Begin VB.Label Label1 
         Caption         =   "Emissor:"
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
         Index           =   50
         Left            =   900
         TabIndex        =   42
         Top             =   1140
         Width           =   690
      End
      Begin VB.Label Label1 
         Caption         =   "% Comissão:"
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
         Index           =   59
         Left            =   4785
         TabIndex        =   41
         Top             =   825
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Correntista:"
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
         Index           =   60
         Left            =   615
         TabIndex        =   39
         Top             =   855
         Width           =   1080
      End
      Begin VB.Label Correntista 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   1710
         TabIndex        =   38
         Top             =   795
         Width           =   2925
      End
      Begin VB.Label Label1 
         Caption         =   "% Comissão:"
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
         Index           =   58
         Left            =   4785
         TabIndex        =   35
         Top             =   525
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Representante:"
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
         Index           =   57
         Left            =   285
         TabIndex        =   33
         Top             =   540
         Width           =   1410
      End
      Begin VB.Label Representante 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   1710
         TabIndex        =   32
         Top             =   495
         Width           =   2925
      End
   End
   Begin MSMask.MaskEdBox TipoVou 
      Height          =   315
      Left            =   735
      TabIndex        =   0
      Top             =   105
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   1
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox SerieVou 
      Height          =   315
      Left            =   1755
      TabIndex        =   1
      Top             =   105
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   1
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NumeroVou 
      Height          =   315
      Left            =   2985
      TabIndex        =   2
      Top             =   105
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   9
      Mask            =   "#########"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
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
      Height          =   330
      Index           =   0
      Left            =   240
      TabIndex        =   54
      Top             =   150
      Width           =   435
   End
   Begin VB.Label Label1 
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
      Height          =   330
      Index           =   2
      Left            =   1200
      TabIndex        =   53
      Top             =   150
      Width           =   480
   End
   Begin VB.Label LabelNumVou 
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
      Height          =   330
      Left            =   2265
      TabIndex        =   52
      Top             =   150
      Width           =   750
   End
End
Attribute VB_Name = "TRVVouComi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim lNumvouAnt As Long
Dim sSerieAnt As String
Dim sTipoAnt As String
Dim lAgenciaAnt As Long

Private WithEvents objEventoEmissor As AdmEvento
Attribute objEventoEmissor.VB_VarHelpID = -1
Private WithEvents objEventoEmissorGrid As AdmEvento
Attribute objEventoEmissorGrid.VB_VarHelpID = -1
Private WithEvents objEventoRepresentante As AdmEvento
Attribute objEventoRepresentante.VB_VarHelpID = -1
Private WithEvents objEventoCorrentista As AdmEvento
Attribute objEventoCorrentista.VB_VarHelpID = -1
Private WithEvents objEventoAgencia As AdmEvento
Attribute objEventoAgencia.VB_VarHelpID = -1
Private WithEvents objEventoPromotor As AdmEvento
Attribute objEventoPromotor.VB_VarHelpID = -1

Dim objGridVendedores As AdmGrid
Dim iGrid_Vendedor_Col As Integer

Dim objGridOVER As AdmGrid
Dim iGrid_OVEREmi_Col As Integer
Dim iGrid_OVERPercTotal_Col As Integer
Dim iGrid_OVERPercVou_Col As Integer
Dim iGrid_OVERValor_Col As Integer
Dim iGrid_OVERCargo_Col As Integer

Dim iAlterado As Integer

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Alterações no comissionamento do voucher"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TRVVouComi"

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

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property
'**** fim do trecho a ser copiado *****

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objEventoEmissor = Nothing
    Set objEventoEmissorGrid = Nothing
    Set objEventoRepresentante = Nothing
    Set objEventoCorrentista = Nothing
    Set objEventoAgencia = Nothing
    Set objEventoPromotor = Nothing

    Set objGridVendedores = Nothing
    Set objGridOVER = Nothing

    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198028)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iAlterado = 0
    
    Set objEventoEmissor = New AdmEvento
    Set objEventoEmissorGrid = New AdmEvento
    Set objEventoRepresentante = New AdmEvento
    Set objEventoCorrentista = New AdmEvento
    Set objEventoAgencia = New AdmEvento
    Set objEventoPromotor = New AdmEvento
    
    Set objGridVendedores = New AdmGrid
    Set objGridOVER = New AdmGrid
    
    If giVersaoTRV = TRV_VERSAO_FATURAMENTO Then
        FrameSuporte.Visible = False
    End If
    
    lErro = Inicializa_Grid_Vendedores(objGridVendedores)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Inicializa_Grid_OVER(objGridOVER)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_TRVEMICARGO, EmiCargo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198029)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objVou As ClassTRVVouchers) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objVou Is Nothing) Then

        lErro = Traz_TRVVouchers_Tela(objVou)
        If lErro <> SUCESSO Then gError 198030

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 198030

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198031)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(ByVal objVou As ClassTRVVouchers) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objVendedor As ClassVendedor
Dim objVouEmiCI As ClassTRVVouEmiCI

On Error GoTo Erro_Move_Tela_Memoria

    objVou.lNumVou = StrParaLong(NumeroVou.Text)
    objVou.sSerie = SerieVou.Text
    objVou.sTipoDoc = TRV_TIPODOC_VOU_TEXTO
    objVou.sTipVou = TipoVou.Text
    
    lErro = CF("TRVVouchers_Le", objVou)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 198200
    
    If lErro <> SUCESSO Then gError 198201
    
    If Len(Trim(PercComiAgeNovo.Text)) <> 0 Then
        objVou.lClienteComissao = LCodigo_Extrai(AgenciaNovo.Caption)
        objVou.dComissaoAg = StrParaDbl(PercComiAgeNovo.Text) / 100
    End If
    
    If LCodigo_Extrai(RepresentanteNovo.Text) <> 0 Then
        objVou.lRepresentante = LCodigo_Extrai(RepresentanteNovo.Text)
        objVou.dComissaoRep = StrParaDbl(PercComiRepNovo.Text) / 100
    End If
    
    If LCodigo_Extrai(CorrentistaNovo.Text) <> 0 Then
        objVou.lCorrentista = LCodigo_Extrai(CorrentistaNovo.Text)
        objVou.dComissaoCorr = StrParaDbl(PercComiCorNovo.Text) / 100
    End If
    
    If LCodigo_Extrai(EmissorNovo.Text) <> 0 Then
        objVou.lEmissor = LCodigo_Extrai(EmissorNovo.Text)
        objVou.dComissaoEmissor = StrParaDbl(PercComiEmiNovo.Text) / 100
    End If

    objVou.lPromotor = LCodigo_Extrai(PromotorNovo.Text)
    
    Set objVou.colVendedores = New Collection
    
    For iIndice = 1 To objGridVendedores.iLinhasExistentes
    
        Set objVendedor = New ClassVendedor
        
        objVendedor.iCodigo = Codigo_Extrai(GridVendedores.TextMatrix(iIndice, iGrid_Vendedor_Col))
    
        objVou.colVendedores.Add objVendedor
    
    Next
    
    Set objVou.colEmiCI = New Collection
    For iIndice = 2 To objGridOVER.iLinhasExistentes
    
        Set objVouEmiCI = New ClassTRVVouEmiCI
        
        objVouEmiCI.lFornEmissor = LCodigo_Extrai(GridOver.TextMatrix(iIndice, iGrid_OVEREmi_Col))
        objVouEmiCI.dPercCI = PercentParaDbl(GridOver.TextMatrix(iIndice, iGrid_OVERPercTotal_Col))
    
        objVou.colEmiCI.Add objVouEmiCI
    
    Next

    If CartaoNovo.Value = vbChecked Then
        objVou.iCartao = MARCADO
    Else
        objVou.iCartao = DESMARCADO
    End If
    
    If AntcNovo.Value = vbChecked Then
        objVou.idiasantc = MARCADO
    Else
        objVou.idiasantc = DESMARCADO
    End If
    
    If Len(Trim(DestinoNovo.Text)) > 0 Then
        objVou.sDestino = DestinoNovo.Text
    End If
    
    If Len(Trim(ProdutoNovo.Text)) > 0 Then
        objVou.sProduto = ProdutoNovo.Text
    End If
    
    If StrParaDbl(BrutoNovo.Text) <> 0 Then
        objVou.dValorBruto = StrParaDbl(BrutoNovo.Text)
    End If

    If StrParaDbl(BaseNovo.Text) <> 0 Then
        objVou.dValorBaseComis = StrParaDbl(BaseNovo.Text)
    End If
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 198200
        
        Case 198201
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_NAO_CADASTRADO", gErr, objVou.lNumVou, objVou.sSerie, objVou.sTipVou)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198032)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objVou As New ClassTRVVouchers

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TRVVouchers"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objVou)
    If lErro <> SUCESSO Then gError 198033

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "TipoDoc", TRV_TIPODOC_VOU_TEXTO, 0, "Codigo"
    colCampoValor.Add "NumVou", objVou.lNumVou, 0, "Codigo"
    colCampoValor.Add "Serie", objVou.sSerie, 0, "Codigo"
    colCampoValor.Add "TiVou", objVou.sTipVou, 0, "Codigo"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 198033

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198034)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objVou As New ClassTRVVouchers

On Error GoTo Erro_Tela_Preenche

    objVou.lNumVou = colCampoValor.Item("NumVou").vValor
    objVou.sSerie = colCampoValor.Item("NumVou").vValor
    objVou.sTipVou = colCampoValor.Item("NumVou").vValor
    
    If objVou.lNumVou <> 0 Then

        lErro = Traz_TRVVouchers_Tela(objVou)
        If lErro <> SUCESSO Then gError 198035

    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 198035

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198036)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objVou As New ClassTRVVouchers
Dim bSimulaImport As Boolean

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(NumeroVou.Text)) = 0 Then gError 198037
    '#####################

    'Preenche o objTRVTiposOcorrencia
    lErro = Move_Tela_Memoria(objVou)
    If lErro <> SUCESSO Then gError 198038

    lErro = CF("TRVEmissor_Valida", objVou.lEmissor, objVou.lClienteComissao)
    If lErro <> SUCESSO Then gError 198038

    lErro = Trata_Alteracao(objVou, objVou.sTipVou, objVou.sSerie, objVou.lNumVou)
    If lErro <> SUCESSO Then gError 198039
    
    If Import.Value = vbChecked Then
        bSimulaImport = True
    Else
        bSimulaImport = False
    End If

    'Grava o/a TRVTiposOcorrencia no Banco de Dados
    lErro = CF("TRVVouComi_Grava", objVou, False, bSimulaImport)
    If lErro <> SUCESSO Then gError 198040
    
    'Limpa Tela
    Call Limpa_Tela_TRVVouchers
    
    Call Traz_TRVVouchers_Tela(objVou)

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 198037
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TRVTIPOSOCORRENCIA_NAO_PREENCHIDO", gErr)

        Case 198038, 198039, 198040

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198041)

    End Select

    Exit Function

End Function

Function Limpa_Tela_TRVVouchers() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_TRVVouchers

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Cartao.Value = vbUnchecked
    Antc.Value = vbUnchecked
    CartaoNovo.Value = Cartao.Value
    AntcNovo.Value = Antc.Value
    Import.Value = vbUnchecked
    
    DataEmissaoVou.Caption = ""
    ValorBrutoVou.Caption = ""
    ValorBase.Caption = ""
    Produto.Caption = ""
    Destino.Caption = ""
    Representante.Caption = ""
    Correntista.Caption = ""
    Emissor.Caption = ""
    Agencia.Caption = ""
    AgenciaVou.Caption = ""
    AgenciaNovo.Caption = ""
    PercComiRep.Caption = ""
    PercComiCor.Caption = ""
    PercComiEmi.Caption = ""
    PercComiAge.Caption = ""
    VlrComiRep.Caption = ""
    VlrComiCor.Caption = ""
    VlrComiEmi.Caption = ""
    VlrComiAge.Caption = ""
    
    lAgenciaAnt = 0
    
    Call Grid_Limpa(objGridOVER)
    Call Grid_Limpa(objGridVendedores)
    CDPercTotal.Caption = ""
    CIPercTotal.Caption = ""

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    iAlterado = 0

    Limpa_Tela_TRVVouchers = SUCESSO

    Exit Function

Erro_Limpa_Tela_TRVVouchers:

    Limpa_Tela_TRVVouchers = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198042)

    End Select

    Exit Function

End Function

Function Traz_TRVVouchers_Tela(ByVal objVou As ClassTRVVouchers) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_TRVVouchers_Tela

    Call Limpa_Tela_TRVVouchers
    
    NumeroVou.PromptInclude = False
    NumeroVou.Text = CStr(objVou.lNumVou)
    NumeroVou.PromptInclude = True
    
    SerieVou.Text = objVou.sSerie
    TipoVou.Text = objVou.sTipVou
    
    Call TrazerVou_Click

    iAlterado = 0

    Traz_TRVVouchers_Tela = SUCESSO

    Exit Function

Erro_Traz_TRVVouchers_Tela:

    Traz_TRVVouchers_Tela = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198044)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 198045
'
'    'Limpa Tela
'    Call Limpa_Tela_TRVVouchers

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 198045

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198046)

    End Select

    Exit Sub

End Sub

Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198047)

    End Select

    Exit Sub

End Sub

Private Sub AgenciaNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AntcNovo_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub BrutoNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CartaoNovo_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CorrentistaNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DestinoNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub EmissorNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PercComiAgeNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PercComiCorNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PercComiEmiNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PercComiRepNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ProdutoNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PromotorNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub RepresentanteNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub SerieVou_Validate(Cancel As Boolean)
    If SerieVou.Text <> sSerieAnt Then
        Call Limpa_Vou
    End If
End Sub

Private Sub TipoVou_Validate(Cancel As Boolean)
    If TipoVou.Text <> sTipoAnt Then
        Call Limpa_Vou
    End If
End Sub

Private Sub TipoVou_Change()
    iAlterado = REGISTRO_ALTERADO
    If Len(Trim(TipoVou.ClipText)) > 0 Then
        If SerieVou.Visible Then SerieVou.SetFocus
    End If
End Sub

Private Sub SerieVou_Change()
    iAlterado = REGISTRO_ALTERADO
    If Len(Trim(SerieVou.ClipText)) > 0 Then
        If NumeroVou.Visible Then NumeroVou.SetFocus
    End If
End Sub

Private Sub TipoVou_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub SerieVou_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub NumeroVou_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NumeroVou_Validate(Cancel As Boolean)
    If StrParaLong(NumeroVou.Text) <> lNumvouAnt Then
        Call Limpa_Vou
    End If
End Sub

Private Sub BotaoTrazerVou_Click()
    Call TrazerVou_Click
End Sub

Private Sub TrazerVou_Click()

Dim lErro As Long
Dim objVoucher As New ClassTRVVouchers
Dim objTRVVoucherInfo As New ClassTRVVoucherInfo
Dim objcliente As New ClassCliente
Dim objForn As New ClassFornecedor
Dim objVendedor As New ClassVendedor
Dim iLinha As Integer
Dim vValor As Variant
Dim objVouEmiCI As ClassTRVVouEmiCI
Dim colAux As New Collection

On Error GoTo Erro_TrazerVou_Click

    objVoucher.lNumVou = StrParaLong(NumeroVou.Text)
    objVoucher.sSerie = SerieVou.Text
    objVoucher.sTipoDoc = TRV_TIPODOC_VOU_TEXTO
    objVoucher.sTipVou = TipoVou.Text
    
    lErro = CF("TRVVouchers_Le", objVoucher)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
    
    If lErro <> SUCESSO Then gError 194425
    
    lErro = CF("TRVVouEmiCI_Le", objVoucher)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If objVoucher.iStatus = STATUS_TRV_VOU_CANCELADO Then gError 194422
    
    'Busca o Cliente no BD
    If objVoucher.lRepresentante <> 0 Then
        Cliente.Text = objVoucher.lRepresentante
        lErro = TP_Cliente_Le2(Cliente, objcliente)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        Representante.Caption = Cliente.Text
    End If
    
    'Busca o Cliente no BD
    If objVoucher.lCorrentista <> 0 Then
        Cliente.Text = objVoucher.lCorrentista
        lErro = TP_Cliente_Le2(Cliente, objcliente)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        Correntista.Caption = Cliente.Text
    End If
    
    'Busca o Cliente no BD
    If objVoucher.lClienteComissao <> 0 Then
        Cliente.Text = objVoucher.lClienteComissao
        lErro = TP_Cliente_Le2(Cliente, objcliente)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        Agencia.Caption = Cliente.Text
        AgenciaNovo.Caption = Cliente.Text
    End If
    If objVoucher.lClienteVou <> 0 Then
        Cliente.Text = objVoucher.lClienteVou
        lErro = TP_Cliente_Le2(Cliente, objcliente)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        AgenciaVou.Caption = Cliente.Text
    End If
    
    'Busca o Cliente no BD
    If objVoucher.lEmissor <> 0 Then
        Cliente.Text = objVoucher.lEmissor
        lErro = TP_Fornecedor_Le2(Cliente, objForn)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        Emissor.Caption = Cliente.Text
    End If
    
    'Busca o Cliente no BD
    If objVoucher.lPromotor <> 0 Then
        PromotorNovo.Text = objVoucher.lPromotor
        lErro = CF("TP_Vendedor_LeTRV", PromotorNovo, objVendedor)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    DataEmissaoVou.Caption = Format(objVoucher.dtData, "dd/mm/yyyy")
    
    ValorBrutoVou.Caption = Format(objVoucher.dValorBrutoComOCR, "STANDARD")
    ValorBase.Caption = Format(IIf(objVoucher.dValorBruto > DELTA_VALORMONETARIO, objVoucher.dValorBaseComis * (objVoucher.dValorBrutoComOCR / objVoucher.dValorBruto), 0), "STANDARD")
    
    Produto.Caption = objVoucher.sProduto
    
    PercComiCor.Caption = Format(objVoucher.dComissaoCorr, "PERCENT")
    PercComiAge.Caption = Format(objVoucher.dComissaoAg, "PERCENT")
    PercComiRep.Caption = Format(objVoucher.dComissaoRep, "PERCENT")
    PercComiEmi.Caption = Format(objVoucher.dComissaoEmissor, "PERCENT")
    
'    VlrComiCor.Caption = Format(objVoucher.dValorBruto * objVoucher.dComissaoCorr, "STANDARD")
'    VlrComiAge.Caption = Format(objVoucher.dValorBruto * objVoucher.dComissaoAg, "STANDARD")
'    VlrComiRep.Caption = Format(objVoucher.dValorBruto * objVoucher.dComissaoRep, "STANDARD")
'    VlrComiEmi.Caption = Format(objVoucher.dValorBruto * objVoucher.dComissaoEmissor, "STANDARD")
'
    VlrComiCor.Caption = Format(objVoucher.dValorCMC, "STANDARD")
    VlrComiAge.Caption = Format(objVoucher.dValorCMAComOCR + objVoucher.dValorCMCC, "STANDARD")
    VlrComiRep.Caption = Format(objVoucher.dValorCMR, "STANDARD")
    VlrComiEmi.Caption = Format(objVoucher.dValorCME, "STANDARD")
    
    If objVoucher.iCartao = MARCADO Then
        Cartao.Value = vbChecked
    Else
        Cartao.Value = vbUnchecked
    End If
    
    If objVoucher.idiasantc = MARCADO Then
        Antc.Value = vbChecked
    Else
        Antc.Value = vbUnchecked
    End If
    
    CartaoNovo.Value = Cartao.Value
    AntcNovo.Value = Antc.Value
    
    Destino.Caption = objVoucher.sDestino
    
    Call Grid_Limpa(objGridVendedores)
    iLinha = 0
    For Each objVendedor In objVoucher.colVendedores
    
        iLinha = iLinha + 1
        
        Cliente.Text = CStr(objVendedor.iCodigo)
        
        lErro = CF("TP_Vendedor_LeTRV", Cliente, objVendedor)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        GridVendedores.TextMatrix(iLinha, iGrid_Vendedor_Col) = Cliente.Text
       
    Next

    objGridVendedores.iLinhasExistentes = iLinha
    
    Set objVouEmiCI = New ClassTRVVouEmiCI
    objVouEmiCI.lFornEmissor = objVoucher.lEmissor
    colAux.Add objVouEmiCI
    
    For Each objVouEmiCI In objVoucher.colEmiCI
        colAux.Add objVouEmiCI
    Next
    
    Call Grid_Limpa(objGridOVER)
    iLinha = 0
    For Each objVouEmiCI In colAux
    
        iLinha = iLinha + 1
        
        Cliente.Text = objVouEmiCI.lFornEmissor
        lErro = TP_Fornecedor_Le2(Cliente, objForn)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
         
        GridOver.TextMatrix(iLinha, iGrid_OVEREmi_Col) = Cliente.Text
        GridOver.TextMatrix(iLinha, iGrid_OVERPercTotal_Col) = Format(objVouEmiCI.dPercCI, "PERCENT")
        
        Call Trata_Emissor_Grid(iLinha, objVouEmiCI.lFornEmissor)
       
    Next

    objGridOVER.iLinhasExistentes = iLinha
    
    Call Ajusta_Over
    
    lNumvouAnt = objVoucher.lNumVou
    sSerieAnt = objVoucher.sSerie
    sTipoAnt = objVoucher.sTipVou
    lAgenciaAnt = objVoucher.lClienteComissao

    Exit Sub

Erro_TrazerVou_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 194422
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_JA_CANCELADO", gErr)
            
'        Case 194424
'            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_SEM_DADOS_SIGAV", gErr)
'
        Case 194425
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_NAO_CADASTRADO", gErr, objVoucher.lNumVou, objVoucher.sSerie, objVoucher.sTipVou)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194426)

    End Select

    Exit Sub
    
End Sub

Private Sub Limpa_Vou()

Dim lNumVou As Long
Dim sSerie As String
Dim sTipVou As String

    lNumVou = StrParaLong(NumeroVou.Text)
    sSerie = SerieVou.Text
    sTipVou = TipoVou.Text

    If lNumVou <> lNumvouAnt Or sSerie <> sSerieAnt Or sTipVou <> sTipoAnt Then
    
        Call Limpa_Tela_TRVVouchers
        
        NumeroVou.PromptInclude = False
        NumeroVou.Text = CStr(lNumVou)
        NumeroVou.PromptInclude = True
        
        SerieVou.Text = sSerie
        TipoVou.Text = sTipVou
        
    End If

End Sub

Private Sub RepresentanteNovo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objcliente As New ClassCliente

On Error GoTo Erro_RepresentanteNovo_Validate

    If Len(Trim(RepresentanteNovo.Text)) > 0 Then
    
        RepresentanteNovo.Text = LCodigo_Extrai(RepresentanteNovo.Text)

        'Tenta ler o Vendedor (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(RepresentanteNovo, objcliente)
        If lErro <> SUCESSO Then gError 195843
        
    End If
    
    Exit Sub

Erro_RepresentanteNovo_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 195843

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195845)
    
    End Select

End Sub

Private Sub CorrentistaNovo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objcliente As New ClassCliente

On Error GoTo Erro_CorrentistaNovo_Validate

    If Len(Trim(CorrentistaNovo.Text)) > 0 Then
    
        CorrentistaNovo.Text = LCodigo_Extrai(CorrentistaNovo.Text)

        'Tenta ler o Vendedor (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(CorrentistaNovo, objcliente)
        If lErro <> SUCESSO Then gError 195843
        
    End If
    
    Exit Sub

Erro_CorrentistaNovo_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 195843

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195845)
    
    End Select

End Sub

Private Sub AgenciaNovo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim objVoucher As New ClassTRVVouchers

On Error GoTo Erro_AgenciaNovo_Validate

    If Len(Trim(AgenciaNovo.Caption)) > 0 Then
    
        AgenciaNovo.Caption = LCodigo_Extrai(AgenciaNovo.Caption)

        'Tenta ler o Vendedor (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(AgenciaNovo, objcliente)
        If lErro <> SUCESSO Then gError 195843
        
        'Se trocou a agência
        If LCodigo_Extrai(AgenciaNovo.Caption) <> LCodigo_Extrai(Agencia.Caption) Then
          
            objVoucher.lNumVou = StrParaLong(NumeroVou.Text)
            objVoucher.sSerie = SerieVou.Text
            objVoucher.sTipoDoc = TRV_TIPODOC_VOU_TEXTO
            objVoucher.sTipVou = TipoVou.Text
              
            lErro = CF("TRVVouchers_Le", objVoucher)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 195843
            
            If lErro = SUCESSO Then
              
                lErro = CF("TRVVou_Le_Dados_Comis_Cliente", objVoucher)
                If lErro <> SUCESSO Then gError 195843
                  
                If LCodigo_Extrai(Representante.Caption) <> objVoucher.lRepresentante And objVoucher.lRepresentante <> 0 Then
                    RepresentanteNovo.Text = objVoucher.lRepresentante
                    Call RepresentanteNovo_Validate(bSGECancelDummy)
                    PercComiRepNovo.Text = objVoucher.dComissaoRep * 100
                    Call PercComiRepNovo_Validate(bSGECancelDummy)
                End If
              
                If LCodigo_Extrai(Correntista.Caption) <> objVoucher.lCorrentista And objVoucher.lCorrentista <> 0 Then
                    CorrentistaNovo.Text = objVoucher.lCorrentista
                    Call CorrentistaNovo_Validate(bSGECancelDummy)
                    PercComiCorNovo.Text = objVoucher.dComissaoCorr * 100
                    Call PercComiCorNovo_Validate(bSGECancelDummy)
                End If
                  
                If LCodigo_Extrai(Emissor.Caption) <> objVoucher.lEmissor And objVoucher.lEmissor <> 0 Then
                    EmissorNovo.Text = objVoucher.lEmissor
                    Call EmissorNovo_Validate(bSGECancelDummy)
                    PercComiEmiNovo.Text = objVoucher.dComissaoEmissor * 100
                    Call PercComiEmiNovo_Validate(bSGECancelDummy)
                End If
                  
                PercComiAgeNovo.Text = objVoucher.dComissaoAg * 100
                Call PercComiAgeNovo_Validate(bSGECancelDummy)
                              
            End If
              
        End If

    End If
       
    Exit Sub

Erro_AgenciaNovo_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 195843

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195845)
    
    End Select

End Sub

Private Sub EmissorNovo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objForn As New ClassFornecedor
Dim objCli As New ClassCliente

On Error GoTo Erro_EmissorNovo_Validate

    If Len(Trim(EmissorNovo.Text)) > 0 Then
    
        EmissorNovo.Text = LCodigo_Extrai(EmissorNovo.Text)

        'Tenta ler o Fornecedor (NomeReduzido ou Código)
        lErro = TP_Fornecedor_Le2(EmissorNovo, objForn)
        If lErro <> SUCESSO Then gError 195843
        
        'Não pode deixar colocar um emissor que não seja desse cliente
        
        'Tenta ler o Cliente (NomeReduzido ou Código)
        objCli.lCodigo = LCodigo_Extrai(AgenciaVou.Caption)
        If lErro <> SUCESSO Then gError 195843
        
        lErro = CF("TRVEmissor_Valida", objForn.lCodigo, objCli.lCodigo)
        If lErro <> SUCESSO Then gError 195843
        
        GridOver.TextMatrix(1, iGrid_OVEREmi_Col) = EmissorNovo.Text
        
    End If
    
    Exit Sub

Erro_EmissorNovo_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 195843

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195845)
    
    End Select

End Sub

Private Sub PromotorNovo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_PromotorNovo_Validate

    If Len(Trim(PromotorNovo.Text)) > 0 Then
    
        PromotorNovo.Text = LCodigo_Extrai(PromotorNovo.Text)

        'Tenta ler o Vendedor (NomeReduzido ou Código)
        lErro = CF("TP_Vendedor_LeTRV", PromotorNovo, objVendedor)
        If lErro <> SUCESSO Then gError 195843
        
    End If
    
    Exit Sub

Erro_PromotorNovo_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 195843

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195845)
    
    End Select

End Sub

Private Sub PercComiAgeNovo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercComiAgeNovo_Validate

    'Verifica se foi preenchido a Comissao de Venda
    If Len(Trim(PercComiAgeNovo.Text)) = 0 Then Exit Sub

    'Critica se é porcentagem
    lErro = Porcentagem_Critica(PercComiAgeNovo.Text)
    If lErro <> SUCESSO Then Error 195853

    'Formata
    PercComiAgeNovo.Text = Format(PercComiAgeNovo.Text, "Fixed")

    Exit Sub

Erro_PercComiAgeNovo_Validate:

    Cancel = True

    Select Case gErr

        Case 195853
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195854)
    
    End Select

    Exit Sub

End Sub

Private Sub PercComiCorNovo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercComiCorNovo_Validate

    'Verifica se foi preenchido a Comissao de Venda
    If Len(Trim(PercComiCorNovo.Text)) = 0 Then Exit Sub

    'Critica se é porcentagem
    lErro = Porcentagem_Critica(PercComiCorNovo.Text)
    If lErro <> SUCESSO Then Error 195853

    'Formata
    PercComiCorNovo.Text = Format(PercComiCorNovo.Text, "Fixed")

    Exit Sub

Erro_PercComiCorNovo_Validate:

    Cancel = True

    Select Case gErr

        Case 195853
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195854)
    
    End Select

    Exit Sub

End Sub

Private Sub PercComiRepNovo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercComiRepNovo_Validate

    'Verifica se foi preenchido a Comissao de Venda
    If Len(Trim(PercComiRepNovo.Text)) = 0 Then Exit Sub

    'Critica se é porcentagem
    lErro = Porcentagem_Critica(PercComiRepNovo.Text)
    If lErro <> SUCESSO Then Error 195853

    'Formata
    PercComiRepNovo.Text = Format(PercComiRepNovo.Text, "Fixed")

    Exit Sub

Erro_PercComiRepNovo_Validate:

    Cancel = True

    Select Case gErr

        Case 195853
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195854)
    
    End Select

    Exit Sub

End Sub

Private Sub PercComiEmiNovo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercComiEmiNovo_Validate

    'Verifica se foi preenchido a Comissao de Venda
    If Len(Trim(PercComiEmiNovo.Text)) = 0 Then Exit Sub

    'Critica se é porcentagem
    lErro = Porcentagem_Critica(PercComiEmiNovo.Text)
    If lErro <> SUCESSO Then Error 195853

    'Formata
    PercComiEmiNovo.Text = Format(PercComiEmiNovo.Text, "Fixed")

    Call Ajusta_Over

    Exit Sub

Erro_PercComiEmiNovo_Validate:

    Cancel = True

    Select Case gErr

        Case 195853
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195854)
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoHist_Click()

Dim lErro As Long
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoHist_Click

    colSelecao.Add StrParaLong(NumeroVou.Text)
    colSelecao.Add TipoVou.Text
    colSelecao.Add SerieVou.Text

    Call Chama_Tela("VoucherInfoLista", colSelecao, Nothing, Nothing, "NumVou= ? AND TipVou = ? AND Serie = ?")

    Exit Sub

Erro_BotaoHist_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190226)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluirComissao_Click()

Dim lErro As Long
Dim objVou As New ClassTRVVouchers

On Error GoTo Erro_BotaoExcluirComissao_Click

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(NumeroVou.Text)) = 0 Then gError 198216
    '#####################

    'Preenche o objTRVTiposOcorrencia
    lErro = Move_Tela_Memoria(objVou)
    If lErro <> SUCESSO Then gError 198217

    'Grava o/a TRVTiposOcorrencia no Banco de Dados
    lErro = CF("TRVVoucher_Exclui_Comissao", objVou)
    If lErro <> SUCESSO Then gError 198218

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluirComissao_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 198216
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TRVTIPOSOCORRENCIA_NAO_PREENCHIDO", gErr)
            
        Case 198217, 198218

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198219)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoPrimeiraComissao_Click()

Dim lErro As Long
Dim objVou As New ClassTRVVouchers

On Error GoTo Erro_BotaoPrimeiraComissao_Click

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(NumeroVou.Text)) = 0 Then gError 198216
    '#####################

    'Preenche o objTRVTiposOcorrencia
    lErro = Move_Tela_Memoria(objVou)
    If lErro <> SUCESSO Then gError 198217

    'Grava o/a TRVTiposOcorrencia no Banco de Dados
    lErro = CF("TRVVouComi_Grava", objVou, True)
    If lErro <> SUCESSO Then gError 198218

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoPrimeiraComissao_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 198216
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TRVTIPOSOCORRENCIA_NAO_PREENCHIDO", gErr)
            
        Case 198217, 198218

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198219)

    End Select

    Exit Sub
    
End Sub

Public Sub LabelRepresentante_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As Collection
Dim sNomeBrowse As String

    'Preenche NomeReduzido com o cliente da tela
    If Len(Trim(RepresentanteNovo.Text)) > 0 Then
        objcliente.lCodigo = LCodigo_Extrai(RepresentanteNovo.Text)
        objcliente.sNomeReduzido = RepresentanteNovo.Text
    End If

    sNomeBrowse = "ClientesLista"

    Call CF("Cliente_Obtem_NomeBrowse", sNomeBrowse)

    'Chama Tela ClienteLista
    Call Chama_Tela(sNomeBrowse, colSelecao, objcliente, objEventoRepresentante)

End Sub

Public Sub objEventoRepresentante_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente
Dim bCancel As Boolean

    Set objcliente = obj1

    RepresentanteNovo.Text = objcliente.lCodigo
    Call RepresentanteNovo_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Public Sub LabelCorrentista_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As Collection
Dim sNomeBrowse As String

    'Preenche NomeReduzido com o cliente da tela
    If Len(Trim(CorrentistaNovo.Text)) > 0 Then
        objcliente.lCodigo = LCodigo_Extrai(CorrentistaNovo.Text)
        objcliente.sNomeReduzido = CorrentistaNovo.Text
    End If

    sNomeBrowse = "ClientesLista"

    Call CF("Cliente_Obtem_NomeBrowse", sNomeBrowse)

    'Chama Tela ClienteLista
    Call Chama_Tela(sNomeBrowse, colSelecao, objcliente, objEventoCorrentista)

End Sub

Public Sub objEventoCorrentista_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente
Dim bCancel As Boolean

    Set objcliente = obj1

    CorrentistaNovo.Text = objcliente.lCodigo
    Call CorrentistaNovo_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Public Sub LabelAgencia_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As Collection
Dim sNomeBrowse As String

'    'Preenche NomeReduzido com o cliente da tela
'    If Len(Trim(AgenciaNovo.Text)) > 0 Then
'        objcliente.lCodigo = LCodigo_Extrai(AgenciaNovo.Text)
'        objcliente.sNomeReduzido = AgenciaNovo.Text
'    End If
'
'    sNomeBrowse = "ClientesLista"
'
'    Call CF("Cliente_Obtem_NomeBrowse", sNomeBrowse)
'
'    'Chama Tela ClienteLista
'    Call Chama_Tela(sNomeBrowse, colSelecao, objcliente, objEventoAgencia)

End Sub

Public Sub objEventoAgencia_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente
Dim bCancel As Boolean

    Set objcliente = obj1
'
'    AgenciaNovo.Text = objcliente.lCodigo
'    Call AgenciaNovo_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Public Sub LabelEmissor_Click()

Dim objForn As New ClassFornecedor
Dim colSelecao As New Collection

    'Preenche NomeReduzido com o cliente da tela
    If Len(Trim(EmissorNovo.Text)) > 0 Then
        objForn.lCodigo = LCodigo_Extrai(EmissorNovo.Text)
        objForn.sNomeReduzido = EmissorNovo.Text
    End If

    colSelecao.Add LCodigo_Extrai(Agencia.Caption)

    'Chama Tela ClienteLista
    Call Chama_Tela("FornecedorLista", colSelecao, objForn, objEventoEmissor, "Codigo IN (SELECT FornEmissor FROM TRVCliEmissores WHERE Cliente = ?)")

End Sub

Public Sub objEventoEmissor_evSelecao(obj1 As Object)

Dim objForn As ClassFornecedor
Dim bCancel As Boolean

    Set objForn = obj1

    EmissorNovo.Text = objForn.lCodigo
    Call EmissorNovo_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Public Sub LabelPromotor_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection

    'Preenche NomeReduzido com o cliente da tela
    If Len(Trim(PromotorNovo.Text)) > 0 Then
        objVendedor.iCodigo = Codigo_Extrai(PromotorNovo.Text)
        objVendedor.sNomeReduzido = PromotorNovo.Text
    End If

    'Chama Tela ClienteLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoPromotor)

End Sub

Public Sub objEventoPromotor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor
Dim bCancel As Boolean

    Set objVendedor = obj1

    PromotorNovo.Text = objVendedor.iCodigo
    Call PromotorNovo_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is AgenciaNovo Then Call LabelAgencia_Click
        If Me.ActiveControl Is RepresentanteNovo Then Call LabelRepresentante_Click
        If Me.ActiveControl Is CorrentistaNovo Then Call LabelCorrentista_Click
        If Me.ActiveControl Is EmissorNovo Then Call LabelEmissor_Click
        If Me.ActiveControl Is OVEREmi Then Call BotaoEmissorGrid_Click
        If Me.ActiveControl Is PromotorNovo Then Call LabelPromotor_Click
    
    End If
    
End Sub

Private Function Inicializa_Grid_Vendedores(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Nome")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Vendedor.Name)

    'Colunas do GridRepr
    iGrid_Vendedor_Col = 1

    'Grid do GridInterno
    objGridInt.objGrid = GridVendedores
    
    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 100

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 3

    'Largura da primeira coluna
    GridVendedores.ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Vendedores = SUCESSO

    Exit Function

End Function

Public Sub GridVendedores_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridVendedores, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridVendedores, iAlterado)
    End If
    
End Sub

Public Sub GridVendedores_EnterCell()
    Call Grid_Entrada_Celula(objGridVendedores, iAlterado)
End Sub

Public Sub GridVendedores_GotFocus()
    Call Grid_Recebe_Foco(objGridVendedores)
End Sub

Public Sub GridVendedores_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call Grid_Trata_Tecla1(KeyCode, objGridVendedores)
    
End Sub

Public Sub GridVendedores_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridVendedores, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridVendedores, iAlterado)
    End If
    
End Sub

Public Sub GridVendedores_LeaveCell()
    Call Saida_Celula(objGridVendedores)
End Sub

Public Sub GridVendedores_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridVendedores)
End Sub

Public Sub GridVendedores_RowColChange()
    Call Grid_RowColChange(objGridVendedores)
End Sub

Public Sub GridVendedores_Scroll()
    Call Grid_Scroll(objGridVendedores)
End Sub


Public Sub Vendedor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Vendedor_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridVendedores)
End Sub

Public Sub Vendedor_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridVendedores)
End Sub

Public Sub Vendedor_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridVendedores.objControle = Vendedor
    lErro = Grid_Campo_Libera_Foco(objGridVendedores)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Function Saida_Celula_Vendedor(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Vendedor do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objVendedor As New ClassVendedor
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Vendedor

    Set objGridInt.objControle = Vendedor

    'Verifica se vendedor está preenchido
    If Len(Trim(Vendedor.Text)) > 0 Then

        Vendedor.Text = LCodigo_Extrai(Vendedor.Text)
        
        'Verifica se Vendedor existe
        lErro = CF("TP_Vendedor_LeTRV", Vendedor, objVendedor, DESMARCADO)
        If lErro <> SUCESSO And lErro <> 25018 And lErro <> 25020 Then gError 200587

        If lErro = 25018 Then gError 200588
        If lErro = 25020 Then gError 200589

        'verifica se precisa preencher o grid com uma nova linha
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 200590

    Saida_Celula_Vendedor = SUCESSO

    Exit Function

Erro_Saida_Celula_Vendedor:

    Saida_Celula_Vendedor = gErr

    Select Case gErr

        Case 200587, 25689, 200590, 200588, 200589
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200591)

    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

         'GridVendedores
        If objGridInt.objGrid.Name = GridVendedores.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col

                 Case iGrid_Vendedor_Col

                    lErro = Saida_Celula_Vendedor(objGridInt)
                    If lErro <> SUCESSO Then gError 200557
                
            End Select
                    
        ElseIf objGridInt.objGrid.Name = GridOver.Name Then
                    
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col

                 Case iGrid_OVEREmi_Col

                    lErro = Saida_Celula_OverEmi(objGridInt)
                    If lErro <> SUCESSO Then gError 200557
                
                 Case iGrid_OVERPercTotal_Col

                    lErro = Saida_Celula_OverPercTotal(objGridInt)
                    If lErro <> SUCESSO Then gError 200557
                
            End Select
            
        End If
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 200558

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 200553 To 200557

        Case 200558
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200559)

    End Select

    Exit Function

End Function

Private Sub BotaoHistOcor_Click()

Dim lErro As Long
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoHistOcor_Click

    colSelecao.Add StrParaLong(NumeroVou.Text)
    colSelecao.Add TipoVou.Text
    colSelecao.Add SerieVou.Text

    Call Chama_Tela("OcorrenciasHistLista", colSelecao, Nothing, Nothing)

    Exit Sub

Erro_BotaoHistOcor_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190226)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Grid_OVER(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Emissor")
    objGridInt.colColuna.Add ("%Over")
    objGridInt.colColuna.Add ("%Vou")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Cargo")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (OVEREmi.Name)
    objGridInt.colCampo.Add (OverPercTotal.Name)
    objGridInt.colCampo.Add (OVERPercVou.Name)
    objGridInt.colCampo.Add (OverValor.Name)
    objGridInt.colCampo.Add (OverCargo.Name)

    'Colunas do GridOver
    iGrid_OVEREmi_Col = 1
    iGrid_OVERPercTotal_Col = 2
    iGrid_OVERPercVou_Col = 3
    iGrid_OVERValor_Col = 4
    iGrid_OVERCargo_Col = 5

    'Grid do GridInterno
    objGridInt.objGrid = GridOver
    
    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 100

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5

    'Largura da primeira coluna
    GridOver.ColWidth(0) = 300

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_OVER = SUCESSO

    Exit Function

End Function

Public Sub GridOVER_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridOVER, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridOVER, iAlterado)
    End If
    
End Sub

Public Sub GridOVER_EnterCell()
    Call Grid_Entrada_Celula(objGridOVER, iAlterado)
End Sub

Public Sub GridOVER_GotFocus()
    Call Grid_Recebe_Foco(objGridOVER)
End Sub

Public Sub GridOVER_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyDelete Or GridOver.Row <> 1 Then
    
        Call Grid_Trata_Tecla1(KeyCode, objGridOVER)
        
    End If
    
    Call Ajusta_Over
    
End Sub

Public Sub GridOVER_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridOVER, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridOVER, iAlterado)
    End If
    
End Sub

Public Sub GridOVER_LeaveCell()
    Call Saida_Celula(objGridOVER)
End Sub

Public Sub GridOVER_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridOVER)
End Sub

Public Sub GridOVER_RowColChange()
    Call Grid_RowColChange(objGridOVER)
End Sub

Public Sub GridOVER_Scroll()
    Call Grid_Scroll(objGridOVER)
End Sub

Public Sub OVEREmi_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub OVEREmi_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridOVER)
End Sub

Public Sub OVEREmi_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOVER)
End Sub

Public Sub OVEREmi_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridOVER.objControle = OVEREmi
    lErro = Grid_Campo_Libera_Foco(objGridOVER)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub OVERPercTotal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub OVERPercTotal_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridOVER)
End Sub

Public Sub OVERPercTotal_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOVER)
End Sub

Public Sub OVERPercTotal_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridOVER.objControle = OverPercTotal
    lErro = Grid_Campo_Libera_Foco(objGridOVER)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable
              
    Select Case objControl.Name
    
        Case OVEREmi.Name, OverPercTotal.Name
            If iLinha <> 1 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
            
    End Select
        
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 190665)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_OverEmi(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objForn As New ClassFornecedor
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_OverEmi

    Set objGridInt.objControle = OVEREmi

    'Verifica se vendedor está preenchido
    If Len(Trim(OVEREmi.Text)) > 0 Then

        If LCodigo_Extrai(OVEREmi.Text) <> 0 Then OVEREmi.Text = LCodigo_Extrai(OVEREmi.Text)
        Cliente.Text = OVEREmi.Text
        
        lErro = TP_Fornecedor_Le2(Cliente, objForn)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = Trata_Emissor_Grid(GridOver.Row, objForn.lCodigo)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_OverEmi = SUCESSO

    Exit Function

Erro_Saida_Celula_OverEmi:

    Saida_Celula_OverEmi = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200591)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_OverPercTotal(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dPercentDesc As Double

On Error GoTo Erro_Saida_Celula_OverPercTotal

    Set objGridInt.objControle = OverPercTotal

    'Verifica se vendedor está preenchido
    If Len(Trim(OverPercTotal.Text)) > 0 Then

         'Critica a porcentagem
        lErro = Porcentagem_Critica(OverPercTotal.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        dPercentDesc = CDbl(OverPercTotal.Text)
        
        OverPercTotal.Text = Format(dPercentDesc, "Fixed")
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call Ajusta_Over

    Saida_Celula_OverPercTotal = SUCESSO

    Exit Function

Erro_Saida_Celula_OverPercTotal:

    Saida_Celula_OverPercTotal = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200591)

    End Select

    Exit Function

End Function

Private Sub BotaoEmissorGrid_Click()

Dim objForn As New ClassFornecedor
Dim colSelecao As New Collection

    If GridOver.Row <> 0 Then

        colSelecao.Add LCodigo_Extrai(Agencia.Caption)
    
        'Chama Tela ClienteLista
        Call Chama_Tela("FornecedorLista", colSelecao, objForn, objEventoEmissorGrid, "Codigo IN (SELECT FornEmissor FROM TRVCliEmissores WHERE Cliente = ?)")

    End If

End Sub

Public Sub objEventoEmissorGrid_evSelecao(obj1 As Object)

Dim objForn As ClassFornecedor
Dim bCancel As Boolean

    Set objForn = obj1

    Call Trata_Emissor_Grid(GridOver.Row, objForn.lCodigo)

    Me.Show

    Exit Sub

End Sub

Private Function Trata_Emissor_Grid(ByVal iLinha As Integer, ByVal lCodigo As Long) As Long

Dim lErro As Long
Dim objForn As New ClassFornecedor
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objCli As New ClassCliente
Dim objCliTRV As New ClassClienteTRV
Dim objEmi As ClassCliEmissoresTRV

On Error GoTo Erro_Trata_Emissor_Grid

    Cliente.Text = lCodigo
    
    lErro = TP_Fornecedor_Le2(Cliente, objForn)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    OVEREmi.Text = Cliente.Text
    
    objCli.lCodigo = LCodigo_Extrai(AgenciaVou.Caption)
    
    lErro = CF("TRVEmissor_Valida", objForn.lCodigo, objCli.lCodigo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF("Cliente_Le_Customizado", objCli, True)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Set objCliTRV = objCli.objInfoUsu
    
    For Each objEmi In objCliTRV.colCliEmissoresTRV
        If objEmi.lFornEmissor = objForn.lCodigo Then
            Exit For
        End If
    Next
    
    Call Combo_Seleciona_ItemData(EmiCargo, objEmi.lCargo)
    GridOver.TextMatrix(iLinha, iGrid_OVERCargo_Col) = EmiCargo.Text
    GridOver.TextMatrix(iLinha, iGrid_OVEREmi_Col) = OVEREmi.Text
        
    'verifica se precisa preencher o grid com uma nova linha
    If GridOver.Row - GridOver.FixedRows = objGridOVER.iLinhasExistentes Then
        objGridOVER.iLinhasExistentes = objGridOVER.iLinhasExistentes + 1
    End If
        
    Trata_Emissor_Grid = SUCESSO

    Exit Function

Erro_Trata_Emissor_Grid:

    Trata_Emissor_Grid = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200591)

    End Select

    Exit Function

End Function

Private Sub Ajusta_Over()

Dim lErro As Long
Dim colEmiCI As New Collection
Dim objVouEmiCI As ClassTRVVouEmiCI
Dim dPercCITotal As Double, dPercOver As Double
Dim iLinha As Integer, iIndice As Integer

On Error GoTo Erro_Ajusta_Over

   If objGridOVER.iLinhasExistentes > 0 Then
    
        For iIndice = 2 To objGridOVER.iLinhasExistentes
        
            Set objVouEmiCI = New ClassTRVVouEmiCI
            
            objVouEmiCI.lFornEmissor = LCodigo_Extrai(GridOver.TextMatrix(iIndice, iGrid_OVEREmi_Col))
            objVouEmiCI.dPercCI = PercentParaDbl(GridOver.TextMatrix(iIndice, iGrid_OVERPercTotal_Col))
        
            colEmiCI.Add objVouEmiCI
            
            dPercCITotal = dPercCITotal + objVouEmiCI.dPercCI
        
        Next
        
        If dPercCITotal > 1 Then gError 206587
        
        If Len(Trim(PercComiEmiNovo.Text)) > 0 Then
            dPercOver = StrParaDbl(PercComiEmiNovo.Text) / 100
        Else
            dPercOver = PercentParaDbl(PercComiEmi.Caption)
        End If
        
        iLinha = 1
        GridOver.TextMatrix(iLinha, iGrid_OVERPercTotal_Col) = Format(1 - dPercCITotal, "PERCENT")
        GridOver.TextMatrix(iLinha, iGrid_OVERPercVou_Col) = Format(dPercOver * (1 - dPercCITotal), "PERCENT")
        GridOver.TextMatrix(iLinha, iGrid_OVERValor_Col) = Format(StrParaDbl(ValorBase.Caption) * dPercOver * (1 - dPercCITotal), "STANDARD")
    
        CDPercTotal.Caption = Format(1 - dPercCITotal, "PERCENT")
        CIPercTotal.Caption = Format(dPercCITotal, "PERCENT")
    
        For Each objVouEmiCI In colEmiCI
            iLinha = iLinha + 1
            GridOver.TextMatrix(iLinha, iGrid_OVERPercTotal_Col) = Format(objVouEmiCI.dPercCI, "PERCENT")
            GridOver.TextMatrix(iLinha, iGrid_OVERPercVou_Col) = Format(dPercOver * objVouEmiCI.dPercCI, "PERCENT")
            GridOver.TextMatrix(iLinha, iGrid_OVERValor_Col) = Format(StrParaDbl(ValorBase.Caption) * dPercOver * objVouEmiCI.dPercCI, "STANDARD")
        Next
    
    End If

    Exit Sub

Erro_Ajusta_Over:

    Select Case gErr

        Case 206587
            Call Rotina_Erro(vbOKOnly, "ERRO_PERC_CI_MAIOR_1", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206588)

    End Select

    Exit Sub
End Sub
