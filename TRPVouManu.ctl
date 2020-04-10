VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl TRPVouManu 
   ClientHeight    =   6240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame12"
      Height          =   5250
      Index           =   3
      Left            =   180
      TabIndex        =   73
      Top             =   855
      Visible         =   0   'False
      Width           =   9195
      Begin VB.Frame Frame8 
         Caption         =   "Outros"
         Height          =   2640
         Left            =   5280
         TabIndex        =   154
         Top             =   2475
         Width           =   3885
         Begin VB.CheckBox ImprimirValor 
            Caption         =   "Imprimir o Valor no Voucher"
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
            Left            =   975
            TabIndex        =   59
            Top             =   1860
            Width           =   2715
         End
         Begin VB.ComboBox Idioma 
            Height          =   315
            ItemData        =   "TRPVouManu.ctx":0000
            Left            =   975
            List            =   "TRPVouManu.ctx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   405
            Width           =   1650
         End
         Begin VB.Label Label1 
            Caption         =   "C.Pagto:"
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
            Index           =   35
            Left            =   180
            TabIndex        =   161
            Top             =   1410
            Width           =   765
         End
         Begin VB.Label CondPagto 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   975
            TabIndex        =   160
            Top             =   1365
            Width           =   1725
         End
         Begin VB.Label Label1 
            Caption         =   "Controle:"
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
            Index           =   33
            Left            =   150
            TabIndex        =   159
            Top             =   900
            Width           =   780
         End
         Begin VB.Label Controle 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   975
            TabIndex        =   158
            Top             =   870
            Width           =   2850
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Idioma:"
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
            Index           =   32
            Left            =   300
            TabIndex        =   156
            Top             =   450
            Width           =   630
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Vendedores"
         Height          =   2640
         Left            =   45
         TabIndex        =   153
         Top             =   2475
         Width           =   5160
         Begin MSMask.MaskEdBox VendedorPerc 
            Height          =   255
            Left            =   3795
            TabIndex        =   164
            Top             =   1125
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Vendedor 
            Height          =   225
            Left            =   525
            TabIndex        =   155
            Top             =   1140
            Width           =   2955
            _ExtentX        =   5212
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
         Begin MSFlexGridLib.MSFlexGrid GridVendedores 
            Height          =   480
            Left            =   60
            TabIndex        =   57
            Top             =   615
            Width           =   5010
            _ExtentX        =   8837
            _ExtentY        =   847
            _Version        =   393216
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox PromotorComissao 
            Height          =   315
            Left            =   1020
            TabIndex        =   55
            Top             =   225
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox PercComiProNovo 
            Height          =   315
            Left            =   3990
            TabIndex        =   56
            Top             =   225
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Format          =   "#0.#0\%"
            PromptChar      =   " "
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
            Index           =   1
            Left            =   2895
            TabIndex        =   165
            Top             =   255
            Width           =   1140
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
            Left            =   120
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   114
            Top             =   255
            Width           =   825
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Comissão"
         Height          =   2310
         Left            =   45
         TabIndex        =   115
         Top             =   45
         Width           =   9120
         Begin VB.CommandButton BotaoRecalcComis 
            Caption         =   "Recalcular"
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
            Left            =   5595
            Picture         =   "TRPVouManu.ctx":0021
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Recalcula os % de comissão"
            Top             =   1755
            Width           =   1080
         End
         Begin MSMask.MaskEdBox RepresentanteComissao 
            Height          =   315
            Left            =   1425
            TabIndex        =   49
            Top             =   645
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox PercComiRepNovo 
            Height          =   315
            Left            =   5610
            TabIndex        =   50
            Top             =   645
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CorrentistaComissao 
            Height          =   315
            Left            =   1425
            TabIndex        =   51
            Top             =   1035
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox PercComiCorNovo 
            Height          =   315
            Left            =   5610
            TabIndex        =   52
            Top             =   1035
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PercComiEmiNovo 
            Height          =   315
            Left            =   5610
            TabIndex        =   53
            Top             =   1410
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PercComiAgeNovo 
            Height          =   315
            Left            =   5610
            TabIndex        =   48
            Top             =   240
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            Caption         =   "Valor R$:"
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
            Index           =   13
            Left            =   6825
            TabIndex        =   139
            Top             =   270
            Width           =   885
         End
         Begin VB.Label Label1 
            Caption         =   "Valor R$:"
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
            Index           =   15
            Left            =   6840
            TabIndex        =   138
            Top             =   1470
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Valor R$:"
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
            Index           =   19
            Left            =   6840
            TabIndex        =   137
            Top             =   1065
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Valor R$:"
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
            Left            =   6840
            TabIndex        =   136
            Top             =   660
            Width           =   870
         End
         Begin VB.Label Label1 
            Caption         =   "Faturável:"
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
            Left            =   6765
            TabIndex        =   135
            Top             =   1860
            Width           =   885
         End
         Begin VB.Label ValorFaturavelRS 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7725
            TabIndex        =   134
            Top             =   1800
            Width           =   1275
         End
         Begin VB.Label ClienteComissao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1425
            TabIndex        =   133
            Top             =   240
            Width           =   3015
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
            Index           =   3
            Left            =   735
            TabIndex        =   132
            Top             =   285
            Width           =   735
         End
         Begin VB.Label EmissorComissao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1425
            TabIndex        =   131
            Top             =   1410
            Width           =   3015
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
            Height          =   330
            Index           =   50
            Left            =   675
            TabIndex        =   130
            Top             =   1455
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
            Index           =   18
            Left            =   4515
            TabIndex        =   129
            Top             =   270
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
            Height          =   330
            Index           =   11
            Left            =   4515
            TabIndex        =   128
            Top             =   675
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
            Height          =   330
            Index           =   14
            Left            =   4515
            TabIndex        =   127
            Top             =   1050
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
            Left            =   390
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   126
            Top             =   1080
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
            Index           =   12
            Left            =   4515
            TabIndex        =   125
            Top             =   1440
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
            Left            =   60
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   124
            Top             =   690
            Width           =   1410
         End
         Begin VB.Label VlrComiAge 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7725
            TabIndex        =   123
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label VlrComiEmi 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7725
            TabIndex        =   122
            Top             =   1410
            Width           =   1275
         End
         Begin VB.Label VlrComiCor 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7725
            TabIndex        =   121
            Top             =   1035
            Width           =   1275
         End
         Begin VB.Label VlrComiRep 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7725
            TabIndex        =   120
            Top             =   645
            Width           =   1275
         End
         Begin VB.Label BrutoNaMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1425
            TabIndex        =   119
            Top             =   1785
            Width           =   1080
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
            Height          =   300
            Index           =   21
            Left            =   855
            TabIndex        =   118
            Top             =   1860
            Width           =   585
         End
         Begin VB.Label BrutoRS 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3375
            TabIndex        =   117
            Top             =   1785
            Width           =   1080
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
            Height          =   300
            Index           =   22
            Left            =   2520
            TabIndex        =   116
            Top             =   1860
            Width           =   915
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5310
      Index           =   1
      Left            =   90
      TabIndex        =   61
      Top             =   825
      Width           =   9360
      Begin VB.TextBox Observacao 
         Height          =   705
         Left            =   1200
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   4245
         Width           =   8085
      End
      Begin VB.Frame Frame18 
         Caption         =   "Passageiros"
         Height          =   2295
         Left            =   15
         TabIndex        =   84
         Top             =   1935
         Width           =   9315
         Begin MSMask.MaskEdBox PaxSobreNome 
            Height          =   315
            Left            =   555
            TabIndex        =   163
            Top             =   570
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
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
         Begin VB.CheckBox PaxCancelado 
            Height          =   315
            Left            =   8295
            TabIndex        =   97
            Top             =   1350
            Width           =   510
         End
         Begin VB.CheckBox PaxTitular 
            Height          =   315
            Left            =   8340
            TabIndex        =   96
            Top             =   1035
            Width           =   480
         End
         Begin VB.ComboBox PaxSexo 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "TRPVouManu.ctx":032E
            Left            =   5820
            List            =   "TRPVouManu.ctx":0338
            Style           =   2  'Dropdown List
            TabIndex        =   95
            Top             =   1020
            Width           =   555
         End
         Begin MSMask.MaskEdBox PaxNumDoc 
            Height          =   315
            Left            =   4440
            TabIndex        =   85
            Top             =   1005
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
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
         Begin MSMask.MaskEdBox PaxTipoDoc 
            Height          =   315
            Left            =   3630
            TabIndex        =   86
            Top             =   1005
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
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
         Begin MSMask.MaskEdBox PaxDataNasc 
            Height          =   315
            Left            =   2625
            TabIndex        =   87
            Top             =   1005
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
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
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PaxNome 
            Height          =   315
            Left            =   405
            TabIndex        =   88
            Top             =   1005
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
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
         Begin MSMask.MaskEdBox PaxValorEmi 
            Height          =   315
            Left            =   6480
            TabIndex        =   93
            Top             =   1020
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PaxValor 
            Height          =   315
            Left            =   7395
            TabIndex        =   94
            Top             =   1020
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin VB.CommandButton BotaoRecalcularPax 
            Caption         =   "Recalcular"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   30
            Picture         =   "TRPVouManu.ctx":0342
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Recalcula os Valores dos Passageiros e o Bruto"
            Top             =   1680
            Visible         =   0   'False
            Width           =   1455
         End
         Begin MSFlexGridLib.MSFlexGrid GridPax 
            Height          =   330
            Left            =   30
            TabIndex        =   18
            Top             =   195
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   582
            _Version        =   393216
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
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
            Height          =   255
            Index           =   17
            Left            =   7095
            TabIndex        =   113
            Top             =   1995
            Width           =   915
         End
         Begin VB.Label BrutoPax 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   8025
            TabIndex        =   112
            Top             =   1935
            Width           =   1245
         End
         Begin VB.Label Label1 
            Caption         =   "Bruto Emissão:"
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
            Index           =   31
            Left            =   1755
            TabIndex        =   92
            Top             =   1995
            Width           =   1380
         End
         Begin VB.Label BrutoUSSPaxEmi 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3150
            TabIndex        =   91
            Top             =   1935
            Width           =   1140
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
            Height          =   255
            Index           =   27
            Left            =   4695
            TabIndex        =   90
            Top             =   1995
            Width           =   570
         End
         Begin VB.Label BrutoUSSPax 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   5295
            TabIndex        =   89
            Top             =   1935
            Width           =   1245
         End
      End
      Begin VB.CommandButton BotaoConsulta 
         Caption         =   "Consulta"
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
         Left            =   7845
         TabIndex        =   21
         ToolTipText     =   "Abre a tela de consulta do Voucher"
         Top             =   4950
         Width           =   1470
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
         Height          =   330
         Left            =   30
         TabIndex        =   20
         ToolTipText     =   "Detalha as informações de valores do Voucher"
         Top             =   4950
         Width           =   1470
      End
      Begin VB.Frame FrameDadosBasicos 
         Caption         =   "Dados Básicos"
         Height          =   1935
         Left            =   15
         TabIndex        =   63
         Top             =   15
         Width           =   9315
         Begin VB.Frame FrameDadosBasicosAux 
            BorderStyle     =   0  'None
            Height          =   645
            Index           =   1
            Left            =   5775
            TabIndex        =   180
            Top             =   1200
            Width           =   3450
            Begin VB.ComboBox Moeda 
               Height          =   315
               Left            =   1830
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   15
               Width           =   1650
            End
            Begin MSMask.MaskEdBox Cambio 
               Height          =   315
               Left            =   1830
               TabIndex        =   16
               Top             =   360
               Width           =   1080
               _ExtentX        =   1905
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
               Format          =   "#,##0.00##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox TarifaUnitaria 
               Height          =   315
               Left            =   0
               TabIndex        =   15
               Top             =   360
               Width           =   780
               _ExtentX        =   1376
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
            Begin VB.Label Vigencia 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   0
               TabIndex        =   184
               Top             =   15
               Width           =   540
            End
            Begin VB.Label Label4 
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
               Height          =   330
               Left            =   570
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   183
               Top             =   60
               Width           =   450
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Câmbio:"
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
               Height          =   285
               Index           =   30
               Left            =   840
               TabIndex        =   182
               Top             =   405
               Width           =   945
            End
            Begin VB.Label Label1 
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
               ForeColor       =   &H00000080&
               Height          =   195
               Index           =   25
               Left            =   1155
               TabIndex        =   181
               Top             =   45
               Width           =   690
            End
         End
         Begin VB.Frame FrameDadosBasicosAux 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   0
            Left            =   420
            TabIndex        =   178
            Top             =   495
            Width           =   4335
            Begin MSMask.MaskEdBox Cliente 
               Height          =   315
               Left            =   750
               TabIndex        =   5
               Top             =   30
               Width           =   3540
               _ExtentX        =   6244
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin VB.Label LabelCliente 
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
               Left            =   0
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   179
               Top             =   75
               Width           =   660
            End
         End
         Begin VB.ComboBox Grupo 
            Height          =   315
            ItemData        =   "TRPVouManu.ctx":064F
            Left            =   7605
            List            =   "TRPVouManu.ctx":065C
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   180
            Width           =   1650
         End
         Begin VB.ComboBox Emissor 
            Height          =   315
            Left            =   5775
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   525
            Width           =   3480
         End
         Begin VB.ComboBox Destino 
            Height          =   315
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   870
            Width           =   3540
         End
         Begin VB.ComboBox Produto 
            Height          =   315
            Left            =   5775
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   870
            Width           =   3480
         End
         Begin MSMask.MaskEdBox VigenciaDe 
            Height          =   315
            Left            =   1170
            TabIndex        =   9
            Top             =   1215
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownVigenciaDe 
            Height          =   300
            Left            =   2325
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1200
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox VigenciaAte 
            Height          =   315
            Left            =   3300
            TabIndex        =   11
            Top             =   1215
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownVigenciaAte 
            Height          =   300
            Left            =   4455
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1215
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox QtdePax 
            Height          =   315
            Left            =   1170
            TabIndex        =   14
            Top             =   1560
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   3
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Plano:"
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
            Index           =   6
            Left            =   6990
            TabIndex        =   157
            Top             =   225
            Width           =   555
         End
         Begin VB.Label TarifaFolheto 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3300
            TabIndex        =   111
            Top             =   1560
            Width           =   915
         End
         Begin VB.Label Label1 
            Caption         =   "Taf. UN Folheto:"
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
            Index           =   70
            Left            =   1755
            TabIndex        =   110
            Top             =   1605
            Width           =   1470
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Taf. Unitária:"
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
            Height          =   315
            Index           =   16
            Left            =   4440
            TabIndex        =   109
            Top             =   1590
            Width           =   1245
         End
         Begin VB.Label Label1 
            Caption         =   "Qtde Pax:"
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
            Height          =   285
            Index           =   28
            Left            =   195
            TabIndex        =   83
            Top             =   1590
            Width           =   885
         End
         Begin VB.Label Label2 
            Caption         =   "Vigência:"
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
            Left            =   4905
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   82
            Top             =   1245
            Width           =   795
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Volta:"
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
            Height          =   315
            Index           =   26
            Left            =   2655
            TabIndex        =   81
            Top             =   1275
            Width           =   540
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Ida:"
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
            Height          =   315
            Index           =   10
            Left            =   510
            TabIndex        =   80
            Top             =   1275
            Width           =   540
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00000080&
            Height          =   315
            Index           =   9
            Left            =   4920
            TabIndex        =   79
            Top             =   585
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00000080&
            Height          =   315
            Index           =   8
            Left            =   255
            TabIndex        =   78
            Top             =   930
            Width           =   810
         End
         Begin VB.Label LabelProduto 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00000080&
            Height          =   330
            Left            =   4905
            TabIndex        =   77
            Top             =   915
            Width           =   810
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
            Left            =   690
            TabIndex        =   71
            Top             =   225
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
            Left            =   1725
            TabIndex        =   70
            Top             =   225
            Width           =   480
         End
         Begin VB.Label LabelNumVou2 
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
            Left            =   2910
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   69
            Top             =   225
            Width           =   750
         End
         Begin VB.Label TipoVou 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1170
            TabIndex        =   67
            Top             =   180
            Width           =   480
         End
         Begin VB.Label SerieVou 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2280
            TabIndex        =   66
            Top             =   180
            Width           =   480
         End
         Begin VB.Label NumeroVou 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3660
            TabIndex        =   65
            Top             =   180
            Width           =   1035
         End
         Begin VB.Label DataEmissaoVou 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   5775
            TabIndex        =   64
            Top             =   180
            Width           =   1140
         End
         Begin VB.Label Label1 
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
            Height          =   330
            Index           =   5
            Left            =   4965
            TabIndex        =   68
            Top             =   225
            Width           =   780
         End
      End
      Begin VB.Label DataLib 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5865
         TabIndex        =   173
         Top             =   4980
         Width           =   1065
      End
      Begin VB.Label UsuarioLib 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3435
         TabIndex        =   172
         Top             =   4980
         Width           =   1395
      End
      Begin VB.Label HoraLib 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   6930
         TabIndex        =   171
         Top             =   4980
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Data\Hora:"
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
         Left            =   4875
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   170
         Top             =   5025
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Usuário:"
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
         Index           =   40
         Left            =   2700
         TabIndex        =   169
         Top             =   5025
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Liberação:"
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
         Index           =   37
         Left            =   1770
         TabIndex        =   168
         Top             =   5025
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Index           =   36
         Left            =   495
         TabIndex        =   167
         Top             =   4260
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5250
      Index           =   2
      Left            =   180
      TabIndex        =   62
      Top             =   855
      Visible         =   0   'False
      Width           =   9195
      Begin VB.Frame FramePagoCC 
         Caption         =   "Pago com Cartão de Crédito"
         Height          =   465
         Left            =   3540
         TabIndex        =   151
         Top             =   2625
         Width           =   5610
         Begin VB.OptionButton OptCCNao 
            Caption         =   "Não"
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
            Left            =   3255
            TabIndex        =   37
            Top             =   195
            Value           =   -1  'True
            Width           =   705
         End
         Begin VB.OptionButton OptCCSim 
            Caption         =   "Sim"
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
            Left            =   1590
            TabIndex        =   36
            Top             =   195
            Width           =   645
         End
      End
      Begin VB.Frame FrameCC 
         Caption         =   "Dados do Pagamento com Cartão de crédito"
         Height          =   2130
         Left            =   3555
         TabIndex        =   143
         Top             =   3060
         Width           =   5595
         Begin VB.TextBox TitularCC 
            Height          =   300
            Left            =   1305
            TabIndex        =   38
            Top             =   270
            Width           =   4230
         End
         Begin VB.ComboBox BandeiraCC 
            Height          =   315
            Left            =   4125
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   630
            Width           =   1425
         End
         Begin VB.ComboBox NumParcelasCC 
            Height          =   315
            Left            =   4125
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   1005
            Width           =   1425
         End
         Begin MSMask.MaskEdBox CPFTitularCC 
            Height          =   315
            Left            =   1305
            TabIndex        =   39
            Top             =   626
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   14
            Mask            =   "##############"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NumeroCC 
            Height          =   315
            Left            =   1305
            TabIndex        =   41
            Top             =   1005
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   16
            Mask            =   "################"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NumAutoCC 
            Height          =   315
            Left            =   1305
            TabIndex        =   45
            Top             =   1740
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValidadeCC 
            Height          =   315
            Left            =   4125
            TabIndex        =   44
            Top             =   1365
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataAutoCC 
            Height          =   315
            Left            =   4125
            TabIndex        =   46
            Top             =   1740
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataAutoCC 
            Height          =   300
            Left            =   5280
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   1725
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox CodSegurancaCC 
            Height          =   315
            Left            =   1305
            TabIndex        =   43
            Top             =   1380
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            Caption         =   "Cód.Segur.:"
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
            Index           =   34
            Left            =   165
            TabIndex        =   166
            Top             =   1425
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Data Autorização:"
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
            Height          =   285
            Index           =   29
            Left            =   2400
            TabIndex        =   152
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Titular:"
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
            Index           =   38
            Left            =   570
            TabIndex        =   150
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Bandeira:"
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
            Index           =   39
            Left            =   3255
            TabIndex        =   149
            Top             =   660
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "Parcelas:"
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
            Index           =   42
            Left            =   3255
            TabIndex        =   148
            Top             =   1065
            Width           =   1395
         End
         Begin VB.Label Label1 
            Caption         =   "CPF:"
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
            Index           =   62
            Left            =   765
            TabIndex        =   147
            Top             =   675
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Cartão:"
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
            Index           =   7
            Left            =   570
            TabIndex        =   146
            Top             =   1065
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Autorização:"
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
            Index           =   23
            Left            =   135
            TabIndex        =   145
            Top             =   1785
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Validade:"
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
            Index           =   24
            Left            =   3270
            TabIndex        =   144
            Top             =   1395
            Width           =   795
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "Em caso de emergência ligar para"
         Height          =   2595
         Left            =   0
         TabIndex        =   140
         Top             =   2625
         Width           =   3525
         Begin MSMask.MaskEdBox ContatoTelefone 
            Height          =   225
            Left            =   2100
            TabIndex        =   141
            Top             =   795
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
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
         Begin MSMask.MaskEdBox ContatoNome 
            Height          =   225
            Left            =   300
            TabIndex        =   142
            Top             =   795
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
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
         Begin MSFlexGridLib.MSFlexGrid GridContatos 
            Height          =   480
            Left            =   30
            TabIndex        =   35
            Top             =   240
            Width           =   3465
            _ExtentX        =   6112
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
      Begin VB.Frame Frame5 
         Caption         =   "Endereço Passageiro Principal"
         Height          =   2655
         Left            =   0
         TabIndex        =   72
         Top             =   -15
         Width           =   9150
         Begin VB.ComboBox Pais 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4470
            TabIndex        =   31
            Top             =   1215
            Width           =   2160
         End
         Begin VB.ComboBox Estado 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1200
            TabIndex        =   30
            Top             =   1200
            Width           =   630
         End
         Begin VB.TextBox Endereco 
            Height          =   315
            Left            =   1200
            MaxLength       =   40
            TabIndex        =   26
            Top             =   270
            Width           =   7770
         End
         Begin MSMask.MaskEdBox Cidade 
            Height          =   315
            Left            =   4485
            TabIndex        =   28
            Top             =   735
            Width           =   2430
            _ExtentX        =   4286
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Bairro 
            Height          =   315
            Left            =   1200
            TabIndex        =   27
            Top             =   735
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   12
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CEP 
            Height          =   315
            Left            =   7905
            TabIndex        =   29
            Top             =   735
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#####-###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Telefone1 
            Height          =   315
            Left            =   1200
            TabIndex        =   32
            Top             =   1680
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   18
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Telefone2 
            Height          =   315
            Left            =   4470
            TabIndex        =   33
            Top             =   1680
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   18
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Email 
            Height          =   315
            Left            =   1200
            TabIndex        =   34
            Top             =   2205
            Width           =   5430
            _ExtentX        =   9578
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Contato 
            Height          =   315
            Left            =   1200
            TabIndex        =   98
            Top             =   1860
            Visible         =   0   'False
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin VB.Label PaisLabel 
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
            Left            =   3945
            TabIndex        =   108
            Top             =   1260
            Width           =   495
         End
         Begin VB.Label Label42 
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
            Left            =   390
            TabIndex        =   107
            Top             =   1905
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.Label Label46 
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
            Left            =   7395
            TabIndex        =   106
            Top             =   810
            Width           =   465
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            Caption         =   "E-mail:"
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
            TabIndex        =   105
            Top             =   2250
            Width           =   585
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            Caption         =   "Telefone 2:"
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
            Left            =   3405
            TabIndex        =   104
            Top             =   1725
            Width           =   1005
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            Caption         =   "Telefone 1:"
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
            TabIndex        =   103
            Top             =   1710
            Width           =   1005
         End
         Begin VB.Label Label54 
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
            Left            =   555
            TabIndex        =   102
            Top             =   780
            Width           =   585
         End
         Begin VB.Label Label55 
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
            Left            =   465
            TabIndex        =   101
            Top             =   1245
            Width           =   675
         End
         Begin VB.Label LabelCidade 
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
            Left            =   3780
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   100
            Top             =   810
            Width           =   675
         End
         Begin VB.Label Label56 
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
            Left            =   225
            TabIndex        =   99
            Top             =   330
            Width           =   915
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7320
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   162
      TabStop         =   0   'False
      Top             =   15
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "TRPVouManu.ctx":067C
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "TRPVouManu.ctx":07FA
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoCancelar 
         Height          =   360
         Left            =   595
         Picture         =   "TRPVouManu.ctx":0D2C
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Cancelar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "TRPVouManu.ctx":0EB6
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoTrazerVou 
      Height          =   315
      Left            =   2910
      Picture         =   "TRPVouManu.ctx":1010
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Trazer Dados"
      Top             =   105
      Width           =   435
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5685
      Left            =   30
      TabIndex        =   60
      Top             =   480
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   10028
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Voucher"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Complemento"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Comissão"
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
   Begin MSMask.MaskEdBox TipoVouP 
      Height          =   315
      Left            =   285
      TabIndex        =   0
      Top             =   120
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   1
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox SerieVouP 
      Height          =   315
      Left            =   855
      TabIndex        =   1
      Top             =   120
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   1
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NumeroVouP 
      Height          =   315
      Left            =   2025
      TabIndex        =   2
      Top             =   120
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   6
      Mask            =   "######"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      Caption         =   "Em vigência:"
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
      Index           =   43
      Left            =   3600
      TabIndex        =   177
      Top             =   165
      Width           =   1155
   End
   Begin VB.Label EmVigencia 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   4755
      TabIndex        =   176
      Top             =   120
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "Faturado:"
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
      Index           =   41
      Left            =   5625
      TabIndex        =   175
      Top             =   165
      Width           =   840
   End
   Begin VB.Label Faturado 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   6480
      TabIndex        =   174
      Top             =   120
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "T:"
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
      Index           =   65
      Left            =   60
      TabIndex        =   74
      Top             =   165
      Width           =   435
   End
   Begin VB.Label Label1 
      Caption         =   "S:"
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
      Index           =   64
      Left            =   630
      TabIndex        =   76
      Top             =   165
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
      Left            =   1275
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   75
      Top             =   165
      Width           =   750
   End
End
Attribute VB_Name = "TRPVouManu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iFrameAtual As Integer

Dim gobjVoucher As ClassTRPVouchers
Dim gcolAdmMeioPagto As Collection

Dim gbTrazendodados As Boolean
Dim gbVoucherBloqueado As Boolean
Dim gbVoucherValorBloqueado As Boolean

Dim sClienteAnt As String
Dim sProdutoAnt As String
Dim iVigenciaAnt As Integer
Dim iMoedaAnt As Integer
Dim iDestinoAnt As Integer

Dim objGridPax As AdmGrid
Dim iGrid_PaxNome_Col As Integer
Dim iGrid_PaxSobreNome_Col As Integer
Dim iGrid_PaxDataNasc_Col As Integer
Dim iGrid_PaxTipoDoc_Col As Integer
Dim iGrid_PaxNumDoc_Col As Integer
Dim iGrid_PaxSexo_Col As Integer
Dim iGrid_PaxValor_Col As Integer
Dim iGrid_PaxValorEmi_Col As Integer
Dim iGrid_PaxTitular_Col As Integer
Dim iGrid_PaxCancelado_Col As Integer

Dim objGridContatos As AdmGrid
Dim iGrid_ContatoNome_Col As Integer
Dim iGrid_ContatoTelefone_Col As Integer

Dim objGridVendedores As AdmGrid
Dim iGrid_Vendedor_Col As Integer
Dim iGrid_VendedorPerc_Col As Integer

Private WithEvents objEventoVoucher As AdmEvento
Attribute objEventoVoucher.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoPais As AdmEvento
Attribute objEventoPais.VB_VarHelpID = -1
Private WithEvents objEventoCidade As AdmEvento
Attribute objEventoCidade.VB_VarHelpID = -1
Private WithEvents objEventoRepresentante As AdmEvento
Attribute objEventoRepresentante.VB_VarHelpID = -1
Private WithEvents objEventoCorrentista As AdmEvento
Attribute objEventoCorrentista.VB_VarHelpID = -1
Private WithEvents objEventoPromotor As AdmEvento
Attribute objEventoPromotor.VB_VarHelpID = -1

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Manutenção de Vouchers"
    Call Form_Load

End Function

Public Function Name() As String
    Name = "TRPVouManu"
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

Private Sub BandeiraCC_Change()
    iAlterado = REGISTRO_ALTERADO
    Call Carrega_NumParcelas
End Sub

Private Sub BandeiraCC_Click()
    iAlterado = REGISTRO_ALTERADO
    Call Carrega_NumParcelas
End Sub

Private Sub BotaoConsulta_Click()

Dim lErro As Long
Dim objVou As New ClassTRPVouchers

On Error GoTo Erro_BotaoConsulta_Click

    objVou.lNumVou = StrParaLong(NumeroVouP.Text)
    objVou.sSerie = SerieVouP.Text
    objVou.sTipVou = TipoVouP.Text
    
    Call Chama_Tela("TRPVoucher", objVou)

    Exit Sub

Erro_BotaoConsulta_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200447)

    End Select

    Exit Sub
    
End Sub

Private Sub Cambio_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CorrentistaComissao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CPFTitularCC_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub


Private Sub Destino_Change()
    Dim sProd As String
    
    iAlterado = REGISTRO_ALTERADO
    If Codigo_Extrai(Destino.Text) <> iDestinoAnt Then
    
        sProd = SCodigo_Extrai(Produto.Text)
        
        Call CF("Carrega_Combo_WebProduto", Produto, Codigo_Extrai(Destino.Text))
        
        Call CF("SCombo_Seleciona2", Produto, sProd)
        
        'Call Calcula_Tarifa
    End If
    iDestinoAnt = Codigo_Extrai(Destino.Text)
End Sub

Private Sub Destino_Click()
    Dim sProd As String
    
    If Codigo_Extrai(Destino.Text) <> iDestinoAnt Then
    
        sProd = SCodigo_Extrai(Produto.Text)
    
        Call CF("Carrega_Combo_WebProduto", Produto, Codigo_Extrai(Destino.Text))
        
        Call CF("SCombo_Seleciona2", Produto, sProd)

        'Call Calcula_Tarifa
    End If
    iDestinoAnt = Codigo_Extrai(Destino.Text)
End Sub

Private Sub Emissor_Change()
    iAlterado = REGISTRO_ALTERADO
    EmissorComissao.Caption = Emissor.Text
End Sub

Private Sub Emissor_Click()
    EmissorComissao.Caption = Emissor.Text
End Sub

Private Sub Grupo_Change()
    iAlterado = REGISTRO_ALTERADO
    
    If Grupo.ListIndex <> -1 Then
        If Grupo.ItemData(Grupo.ListIndex) = TRP_GRUPO_INDIVIDUAL Then
            QtdePax.PromptInclude = False
            QtdePax.Text = 1
            QtdePax.PromptInclude = True
'            QtdePax.Enabled = False
'        Else
'            QtdePax.Enabled = True
        End If
    End If
    Call BotaoRecalcularPax_Click
End Sub

Private Sub Grupo_Click()
    If Grupo.ListIndex <> -1 Then
        If Grupo.ItemData(Grupo.ListIndex) = TRP_GRUPO_INDIVIDUAL Then
            QtdePax.PromptInclude = False
            QtdePax.Text = 1
            QtdePax.PromptInclude = True
'            QtdePax.Enabled = False
'        Else
'            QtdePax.Enabled = True
        End If
    End If
    Call BotaoRecalcularPax_Click
End Sub

Private Sub LabelNumVou_Click()
    Call BotaoVou_Click
End Sub

Private Sub LabelNumVou2_Click()
    Call BotaoVou_Click
End Sub

Private Sub Moeda_Change()
    iAlterado = REGISTRO_ALTERADO
    Call Calcula_Cambio
End Sub

Private Sub Moeda_Click()
    Call Calcula_Cambio
End Sub

Private Sub NumAutoCC_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NumAutoCC_GotFocus()
    Call MaskEdBox_TrataGotFocus(NumAutoCC, iAlterado)
End Sub

Private Sub NumeroCC_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NumeroCC_GotFocus()
    Call MaskEdBox_TrataGotFocus(NumeroCC, iAlterado)
End Sub

Private Sub NumParcelasCC_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub OptCCSim_Click()
    iAlterado = REGISTRO_ALTERADO
    If OptCCSim.Value Then
        FrameCC.Enabled = True
    Else
        Call Limpa_FrameCC
    End If
    Call Calcula_Valores
End Sub

Private Sub OptCCNao_Click()
    iAlterado = REGISTRO_ALTERADO
    If OptCCNao.Value Then
        Call Limpa_FrameCC
    Else
        FrameCC.Enabled = True
    End If
    Call Calcula_Valores
End Sub

Private Sub Limpa_FrameCC()

    OptCCNao.Value = True
    FrameCC.Enabled = False
    
    TitularCC.Text = ""

    CPFTitularCC.PromptInclude = False
    CPFTitularCC.Text = ""
    CPFTitularCC.PromptInclude = True
    BandeiraCC.ListIndex = -1
    NumParcelasCC.Clear

    NumeroCC.PromptInclude = False
    NumeroCC.Text = ""
    NumeroCC.PromptInclude = True

    ValidadeCC.PromptInclude = False
    ValidadeCC.Text = ""
    ValidadeCC.PromptInclude = True

    NumAutoCC.PromptInclude = False
    NumAutoCC.Text = ""
    NumAutoCC.PromptInclude = True

    DataAutoCC.PromptInclude = False
    DataAutoCC.Text = ""
    DataAutoCC.PromptInclude = True
    
    CodSegurancaCC.PromptInclude = False
    CodSegurancaCC.Text = ""
    CodSegurancaCC.PromptInclude = True
        
End Sub

Public Sub PaxSobreNome_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PaxSobreNome_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPax)
End Sub

Public Sub PaxSobreNome_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPax)
End Sub

Public Sub PaxSobreNome_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridPax.objControle = PaxSobreNome
    lErro = Grid_Campo_Libera_Foco(objGridPax)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub PercComiProNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Produto_Change()
    iAlterado = REGISTRO_ALTERADO
    Call Calcula_Tarifa
End Sub

Private Sub Produto_Click()
    Call Calcula_Tarifa
End Sub

Private Sub PromotorComissao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub QtdePax_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub RepresentanteComissao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub TarifaUnitaria_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub TipoVouP_Change()
    If Len(Trim(TipoVouP.ClipText)) > 0 Then
        If SerieVouP.Visible Then SerieVouP.SetFocus
    End If
End Sub

Private Sub SerieVouP_Change()
    If Len(Trim(SerieVouP.ClipText)) > 0 Then
        If NumeroVouP.Visible Then NumeroVouP.SetFocus
    End If
End Sub

Private Sub TipoVouP_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub SerieVouP_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TitularCC_Change()
    iAlterado = REGISTRO_ALTERADO
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

    Set objEventoVoucher = Nothing
    Set objEventoCliente = Nothing
    Set objEventoProduto = Nothing
    Set objEventoPais = Nothing
    Set objEventoCidade = Nothing
    Set objEventoCorrentista = Nothing
    Set objEventoPromotor = Nothing
    Set objEventoRepresentante = Nothing
    
    Set objGridVendedores = Nothing
    Set objGridPax = Nothing
    Set objGridContatos = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)
    
    Set gobjVoucher = Nothing
    Set gcolAdmMeioPagto = Nothing

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200448)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long
Dim colAdmMeioPagto As New Collection
Dim objAdmMeioPagto As ClassAdmMeioPagto
Dim colCodigoNome As New AdmCollCodigoNome
Dim objCodigoNome As AdmlCodigoNome
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome
Dim colCodigo As New Collection
Dim vCodigo As Variant
Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    gbTrazendodados = False
    gbVoucherBloqueado = False
    gbVoucherValorBloqueado = False

    Set objEventoVoucher = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoPais = New AdmEvento
    Set objEventoCidade = New AdmEvento
    Set objEventoCorrentista = New AdmEvento
    Set objEventoPromotor = New AdmEvento
    Set objEventoRepresentante = New AdmEvento

    Set objGridPax = New AdmGrid
    Set objGridContatos = New AdmGrid
    Set objGridVendedores = New AdmGrid

    'Carrega a combo Destino
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_DESTINO_VIAGEM, Destino)
    If lErro <> SUCESSO Then gError 200449

    lErro = Inicializa_Grid_Vendedores(objGridVendedores)
    If lErro <> SUCESSO Then gError 200451

    lErro = Inicializa_Grid_Pax(objGridPax)
    If lErro <> SUCESSO Then gError 200452

    lErro = Inicializa_Grid_Contatos(objGridContatos)
    If lErro <> SUCESSO Then gError 200453
    
    lErro = Carrega_Moeda
    If lErro <> SUCESSO Then gError 200454
    
    BandeiraCC.Clear
    lErro = CF("AdmMeioPagto_Le_TipoMeioPagto", 5, colAdmMeioPagto)
    If lErro <> SUCESSO Then gError 200455
    
    For Each objAdmMeioPagto In colAdmMeioPagto
    
        lErro = CF("AdmMeioPagtoCondPagto_Le", objAdmMeioPagto)
        If lErro <> SUCESSO Then gError 200456
    
        BandeiraCC.AddItem objAdmMeioPagto.sNome
        BandeiraCC.ItemData(BandeiraCC.NewIndex) = objAdmMeioPagto.iCodigo
        
    Next
    
    'Lê cada código da tabela Estados e poe na coleção colCodigo
    lErro = CF("Codigos_Le", "Estados", "Sigla", TIPO_STR, colCodigo, STRING_ESTADOS_SIGLA)
    If lErro <> SUCESSO Then gError 200457

    For Each vCodigo In colCodigo
        Estado.AddItem vCodigo
    Next

    Set colCodigoDescricao = New AdmColCodigoNome

    'Lê cada código e descrição da tabela Paises
    lErro = CF("Cod_Nomes_Le", "Paises", "Codigo", "Nome", STRING_PAISES_NOME, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 200458

    'Preenche cada ComboBox País com os objetos da coleção colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao

        Pais.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        Pais.ItemData(Pais.NewIndex) = objCodigoDescricao.iCodigo

    Next

    'Seleciona Brasil se existir
    For iIndice = 0 To Pais.ListCount - 1
        If right(Pais.List(iIndice), 6) = "Brasil" Then
            Pais.ListIndex = iIndice
            Exit For
        End If
    Next
    
    Set gcolAdmMeioPagto = colAdmMeioPagto

    iAlterado = 0
    
    iFrameAtual = 1

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 200449 To 200458

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200459)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objVou As ClassTRPVouchers) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objVou Is Nothing) Then
    
        lErro = Traz_TRPVouchers_Tela(objVou)
        If lErro <> SUCESSO Then gError 200460

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 200460

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200461)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objVou As New ClassTRPVouchers

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TRPVouchers"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objVou, False)
    If lErro <> SUCESSO Then gError 200462

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumVou", objVou.lNumVou, 0, "NumVou"
    colCampoValor.Add "Serie", objVou.sSerie, STRING_TRP_OCR_SERIE, "Serie"
    colCampoValor.Add "TipVou", objVou.sTipVou, STRING_TRP_OCR_TIPOVOU, "TipVou"
    
'    colSelecao.Add "NumFat", OP_IGUAL, 0
'    colSelecao.Add "DataVigenciaDe", OP_MAIOR, Date
    colSelecao.Add "Status", OP_DIFERENTE, STATUS_TRP_VOU_CANCELADO
    colSelecao.Add "GeraComissao", OP_IGUAL, DESMARCADO

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 200462

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200463)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objVou As New ClassTRPVouchers

On Error GoTo Erro_Tela_Preenche

    objVou.lNumVou = colCampoValor.Item("NumVou").vValor
    objVou.sSerie = colCampoValor.Item("Serie").vValor
    objVou.sTipVou = colCampoValor.Item("TipVou").vValor

    If objVou.lNumVou <> 0 And Len(Trim(objVou.sSerie)) > 0 And Len(Trim(objVou.sTipVou)) > 0 Then
        
        lErro = Traz_TRPVouchers_Tela(objVou)
        If lErro <> SUCESSO Then gError 200464
        
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 200464

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200465)

    End Select

    Exit Function

End Function

Function Limpa_Tela_TRPVouchers() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_TRPVouchers

    gbTrazendodados = False
    gbVoucherBloqueado = False
    gbVoucherValorBloqueado = False

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    sClienteAnt = ""
    sProdutoAnt = ""
    iVigenciaAnt = 0
    iMoedaAnt = 0
    iDestinoAnt = 0
    
    EmVigencia.ForeColor = vbBlack
    EmVigencia.Caption = "NÃO"
    Faturado.ForeColor = vbBlack
    Faturado.Caption = "NÃO"

    TipoVou.Caption = ""
    SerieVou.Caption = ""
    NumeroVou.Caption = ""
    DataEmissaoVou.Caption = ""
    Vigencia.Caption = ""
    TarifaFolheto.Caption = ""
    BrutoUSSPaxEmi.Caption = ""
    BrutoUSSPax.Caption = ""
    BrutoPax.Caption = ""
    ClienteComissao.Caption = ""
    EmissorComissao.Caption = ""
    BrutoNaMoeda.Caption = ""
    BrutoRS.Caption = ""
    ValorFaturavelRS.Caption = ""
    VlrComiAge.Caption = ""
    VlrComiRep.Caption = ""
    VlrComiCor.Caption = ""
    VlrComiEmi.Caption = ""
    Controle.Caption = ""
    CondPagto.Caption = ""
    
    UsuarioLib.Caption = ""
    DataLib.Caption = ""
    HoraLib.Caption = ""
    
    ImprimirValor.Value = vbChecked
    
    VigenciaAte.Enabled = True
    UpDownVigenciaAte.Enabled = True
    
    'QtdePax.Enabled = True
    
    Call Bloqueia_Funcionalidades(False)
    
    OptCCNao.Value = vbChecked
    
    Emissor.Clear
       
    Set gobjVoucher = Nothing
    
    Call Grid_Limpa(objGridContatos)
    Call Grid_Limpa(objGridVendedores)
    Call Grid_Limpa(objGridPax)

    iAlterado = 0

    Limpa_Tela_TRPVouchers = SUCESSO

    Exit Function

Erro_Limpa_Tela_TRPVouchers:

    Limpa_Tela_TRPVouchers = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200466)

    End Select

    Exit Function

End Function

Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200467)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    Call Limpa_Tela_TRPVouchers

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200468)

    End Select

    Exit Sub

End Sub

Sub BotaoCancelar_Click()

Dim lErro As Long
Dim objVou As New ClassTRPVouchers
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoCancelar_Click

    GL_objMDIForm.MousePointer = vbHourglass

    lErro = Move_Tela_Memoria(objVou)
    If lErro <> SUCESSO Then gError 200469
    
    lErro = CF("TRPVouchers_Le", objVou)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 198200
    
    If lErro <> SUCESSO Then gError 198201
    
    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_CANCELAMENTO_TRPVOUCHERS")

    If vbMsgRes = vbYes Then

        If objVou.iStatus = STATUS_TRP_OCR_CANCELADO Then gError 192650
    
        'Cancela o voucher
        lErro = CF("TRPVoucher_Exclui", objVou)
        If lErro <> SUCESSO Then gError 200470

        'Limpa Tela
        Call Limpa_Tela_TRPVouchers
        
    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoCancelar_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 200469, 200470, 198200
        
        Case 192650
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_JA_CANCELADO", gErr)

        Case 198201
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_NAO_CADASTRADO", gErr, objVou.lNumVou, objVou.sSerie, objVou.sTipVou)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200471)

    End Select

    Exit Sub

End Sub

Private Sub objEventoVoucher_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objVou As ClassTRPVouchers

On Error GoTo Erro_objEventoVoucher_evSelecao

    Set objVou = obj1

    'Mostra os dados do TRPVouchers na tela
    lErro = Traz_TRPVouchers_Tela(objVou)
    If lErro <> SUCESSO Then gError 200472

    Me.Show

    Exit Sub

Erro_objEventoVoucher_evSelecao:

    Select Case gErr

        Case 200472

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200473)

    End Select

    Exit Sub

End Sub

Private Sub BotaoVou_Click()

Dim lErro As Long
Dim objVoucher As New ClassTRPVouchers
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoVou_Click

    objVoucher.lNumVou = StrParaLong(NumeroVou.Caption)
    objVoucher.sSerie = SerieVou.Caption
    objVoucher.sTipVou = TipoVou.Caption

    'Call Chama_Tela("TRPVoucherRapidoLista", colSelecao, objVoucher, objEventoVoucher, "CodCli <> 0 AND Fatura = 0 AND Cancelado = 'Não' AND DataVigIni > ?")
    Call Chama_Tela("TRPVoucherRapidoLista", colSelecao, objVoucher, objEventoVoucher, "CodCli <> 0 AND Cancelado = 'Não'")

    Exit Sub

Erro_BotaoVou_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200474)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long
Dim iLinha As Integer
Dim iFrameAnterior

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index = iFrameAtual Then Exit Sub

    If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

    'Torna Frame correspondente ao Tab selecionado visivel
    Frame1(TabStrip1.SelectedItem.Index).Visible = True
    'Torna Frame atual invisivel
    Frame1(iFrameAtual).Visible = False
    'Armazena novo valor de iFrameAtual
    iFrameAtual = TabStrip1.SelectedItem.Index

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200475)

    End Select

    Exit Sub

End Sub

Private Sub BotaoTrazerVou_Click()

Dim lErro As Long
Dim objVoucher As New ClassTRPVouchers

On Error GoTo Erro_BotaoTrazerVou_Click

    objVoucher.lNumVou = StrParaLong(NumeroVouP.Text)
    objVoucher.sSerie = SerieVouP.Text
    objVoucher.sTipVou = TipoVouP.Text
    
    lErro = Trata_Parametros(objVoucher)
    If lErro <> SUCESSO Then gError 200476

    Exit Sub

Erro_BotaoTrazerVou_Click:

    Select Case gErr
    
        Case 200476
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200477)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoHist_Click()

Dim lErro As Long
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoHist_Click

    colSelecao.Add StrParaLong(NumeroVou.Caption)
    colSelecao.Add TipoVou.Caption
    colSelecao.Add SerieVou.Caption

    Call Chama_Tela("TRPVoucherInfoLista", colSelecao, Nothing, Nothing, "NumVou= ? AND TipVou = ? AND Serie = ?")

    Exit Sub

Erro_BotaoHist_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200478)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Grid_Pax(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("T")
    objGridInt.colColuna.Add ("C")
    objGridInt.colColuna.Add ("SobreNome")
    objGridInt.colColuna.Add ("Nome")
    objGridInt.colColuna.Add ("Data Nasc")
    objGridInt.colColuna.Add ("Tipo Doc")
    objGridInt.colColuna.Add ("Núm. Doc")
    objGridInt.colColuna.Add ("Sexo")
    objGridInt.colColuna.Add ("Valor Emi.")
    objGridInt.colColuna.Add ("Valor Atu.")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (PaxTitular.Name)
    objGridInt.colCampo.Add (PaxCancelado.Name)
    objGridInt.colCampo.Add (PaxSobreNome.Name)
    objGridInt.colCampo.Add (PaxNome.Name)
    objGridInt.colCampo.Add (PaxDataNasc.Name)
    objGridInt.colCampo.Add (PaxTipoDoc.Name)
    objGridInt.colCampo.Add (PaxNumDoc.Name)
    objGridInt.colCampo.Add (PaxSexo.Name)
    objGridInt.colCampo.Add (PaxValorEmi.Name)
    objGridInt.colCampo.Add (PaxValor.Name)

    'Colunas do GridRepr
    iGrid_PaxTitular_Col = 1
    iGrid_PaxCancelado_Col = 2
    iGrid_PaxSobreNome_Col = 3
    iGrid_PaxNome_Col = 4
    iGrid_PaxDataNasc_Col = 5
    iGrid_PaxTipoDoc_Col = 6
    iGrid_PaxNumDoc_Col = 7
    iGrid_PaxSexo_Col = 8
    iGrid_PaxValorEmi_Col = 9
    iGrid_PaxValor_Col = 10

    'Grid do GridInterno
    objGridInt.objGrid = GridPax

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
    
    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 201

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 4
    
    'Largura da primeira coluna
    GridPax.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Pax = SUCESSO

    Exit Function

End Function

Public Sub GridPax_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridPax, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPax, iAlterado)
    End If
    
End Sub

Public Sub GridPax_EnterCell()
    Call Grid_Entrada_Celula(objGridPax, iAlterado)
End Sub

Public Sub GridPax_GotFocus()
    Call Grid_Recebe_Foco(objGridPax)
End Sub

Public Sub GridPax_KeyDown(KeyCode As Integer, Shift As Integer)
    
On Error GoTo Erro_GridPax_KeyDown
    
    If KeyCode = vbKeyDelete Then
        If Abs(StrParaDbl(GridPax.TextMatrix(GridPax.Row, iGrid_PaxValorEmi_Col))) > DELTA_VALORMONETARIO Then gError 200479
        If gbVoucherBloqueado Or gbVoucherValorBloqueado Then gError 200628
    End If
        
    Call Grid_Trata_Tecla1(KeyCode, objGridPax)
    
    If KeyCode = vbKeyDelete Then
        Call BotaoRecalcularPax_Click
    End If

    Exit Sub
    
Erro_GridPax_KeyDown:

    Select Case gErr
    
        Case 200479
            Call Rotina_Erro(vbOKOnly, "ERRO_EXC_PAX_EMITIDO", gErr)
            
        Case 200628
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_BLOQUEADA_VOU", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200480)
    
    End Select

    Exit Sub

End Sub

Public Sub GridPax_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Trata_Tecla(KeyAscii, objGridPax, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPax, iAlterado)
    End If
End Sub

Public Sub GridPax_LeaveCell()
    Call Saida_Celula(objGridPax)
End Sub

Public Sub GridPax_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridPax)
End Sub

Public Sub GridPax_RowColChange()
    Call Grid_RowColChange(objGridPax)
End Sub

Public Sub GridPax_Scroll()
    Call Grid_Scroll(objGridPax)
End Sub

Private Function Inicializa_Grid_Contatos(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Nome")
    objGridInt.colColuna.Add ("Telefone")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ContatoNome.Name)
    objGridInt.colCampo.Add (ContatoTelefone.Name)

    'Colunas do GridRepr
    iGrid_ContatoNome_Col = 1
    iGrid_ContatoTelefone_Col = 2

    'Grid do GridInterno
    objGridInt.objGrid = GridContatos

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 100

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 8

    'Largura da primeira coluna
    GridContatos.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Contatos = SUCESSO

    Exit Function

End Function

Public Sub GridContatos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridContatos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridContatos, iAlterado)
    End If
    
End Sub

Public Sub GridContatos_EnterCell()
    Call Grid_Entrada_Celula(objGridContatos, iAlterado)
End Sub

Public Sub GridContatos_GotFocus()
    Call Grid_Recebe_Foco(objGridContatos)
End Sub

Public Sub GridContatos_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call Grid_Trata_Tecla1(KeyCode, objGridContatos)
    
End Sub

Public Sub GridContatos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridContatos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridContatos, iAlterado)
    End If
    
End Sub

Public Sub GridContatos_LeaveCell()
    Call Saida_Celula(objGridContatos)
End Sub

Public Sub GridContatos_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridContatos)
End Sub

Public Sub GridContatos_RowColChange()
    Call Grid_RowColChange(objGridContatos)
End Sub

Public Sub GridContatos_Scroll()
    Call Grid_Scroll(objGridContatos)
End Sub

Private Function Inicializa_Grid_Vendedores(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Nome")
    objGridInt.colColuna.Add ("% Comissão")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Vendedor.Name)
    objGridInt.colCampo.Add (VendedorPerc.Name)

    'Colunas do GridRepr
    iGrid_Vendedor_Col = 1
    iGrid_VendedorPerc_Col = 2

    'Grid do GridInterno
    objGridInt.objGrid = GridVendedores
    
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
    
    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 100

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 6

    'Largura da primeira coluna
    GridVendedores.ColWidth(0) = 400

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

Public Sub LabelCliente_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    objCliente.sNomeReduzido = Cliente.Text

    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    'Preenche campo Cliente
    Cliente.Text = objCliente.sNomeReduzido

    'Executa o Validate
    Call Cliente_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome
Dim iCodFilial As Integer

On Error GoTo Erro_Cliente_Validate


    'Verifica se o Cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        'Busca o Cliente no BD
        lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
        If lErro <> SUCESSO Then gError 200481
    
        If sClienteAnt <> Cliente.Text Then
        
            'Lê cada código e descrição da tabela Paises
            lErro = CF("Cod_Nomes_Le", "WebEmissores", "Codigo", "NomeReduzido", STRING_FORNECEDOR_NOME_REDUZIDO, colCodigoDescricao, "Cliente = " & CStr(objCliente.lCodigo))
            If lErro <> SUCESSO Then gError 200482
        
            'Preenche cada ComboBox País com os objetos da coleção colCodigoDescricao
            Emissor.Clear
            For Each objCodigoDescricao In colCodigoDescricao
        
                Emissor.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
                Emissor.ItemData(Emissor.NewIndex) = objCodigoDescricao.iCodigo
        
            Next
            
        End If
        
        ClienteComissao.Caption = objCliente.lCodigo & SEPARADOR & objCliente.sNomeReduzido
        
        Call Calcula_Tarifa
        
    Else
        Emissor.Clear
    End If
    
    sClienteAnt = Cliente.Text
    
    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 200481, 200482

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200483)

    End Select

    Exit Sub

End Sub

Private Sub Cliente_GotFocus()
    Call MaskEdBox_TrataGotFocus(Cliente, iAlterado)
End Sub

Private Sub Cliente_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Function Carrega_Moeda() As Long

Dim lErro As Long
Dim objMoeda As ClassMoedas
Dim colMoedas As New Collection

On Error GoTo Erro_Carrega_Moeda
    
    lErro = CF("Moedas_Le_Todas", colMoedas)
    If lErro <> SUCESSO Then gError 200484
    
    'se não existem Moedas cadastradas
    If colMoedas.Count = 0 Then gError 200485
    
    For Each objMoeda In colMoedas
    
        Moeda.AddItem objMoeda.iCodigo & SEPARADOR & objMoeda.sNome
        Moeda.ItemData(Moeda.NewIndex) = objMoeda.iCodigo
    
    Next

    Carrega_Moeda = SUCESSO
    
    Exit Function
    
Erro_Carrega_Moeda:

    Carrega_Moeda = gErr
    
    Select Case gErr
    
        Case 200484
        
        Case 200485
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDAS_NAO_CADASTRADAS", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200486)
    
    End Select

End Function

Function Carrega_NumParcelas() As Long

Dim lErro As Long
Dim objAdmMeioPagto As ClassAdmMeioPagto
Dim objCondPagto As ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_Carrega_NumParcelas

    NumParcelasCC.Clear

    If BandeiraCC.ListIndex <> -1 Then

        For Each objAdmMeioPagto In gcolAdmMeioPagto
        
            If objAdmMeioPagto.iCodigo = BandeiraCC.ItemData(BandeiraCC.ListIndex) Then
            
                For Each objCondPagto In objAdmMeioPagto.colCondPagtoLoja
                    NumParcelasCC.AddItem objCondPagto.iNumParcelas
                    NumParcelasCC.ItemData(NumParcelasCC.NewIndex) = objCondPagto.iParcelamento
                Next
                Exit For
            End If
        
        Next
        
    End If

    Carrega_NumParcelas = SUCESSO
    
    Exit Function
    
Erro_Carrega_NumParcelas:

    Carrega_NumParcelas = gErr
    
    Select Case gErr
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200487)
    
    End Select

End Function

Function Move_Tela_Memoria(objVou As ClassTRPVouchers, Optional bValida As Boolean = True, Optional bGravando As Boolean = False) As Long

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objEmi As New ClassCliEmissoresTRP
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iIndice As Integer
Dim objPax As ClassTRPVouPassageiros
Dim objContato As ClassTRPVouContatos
Dim objVendedor As ClassVendedor
Dim iSeq As Integer

On Error GoTo Erro_Move_Tela_Memoria

    If bValida Then
    
        If Len(Trim(NumeroVou.Caption)) = 0 Then gError 200488
        If Len(Trim(SerieVou.Caption)) = 0 Then gError 200489
        If Len(Trim(TipoVou.Caption)) = 0 Then gError 200490
        
    End If

    objVou.lNumVou = StrParaLong(NumeroVou.Caption)
    objVou.sTipVou = TipoVou.Caption
    objVou.sSerie = SerieVou.Caption
    objVou.dtData = StrParaDate(DataEmissaoVou.Caption)
    
    If bGravando Then
    
        lErro = CF("TRPVouchers_Le", objVou)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 200491
        
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 200496
    
        Set objVou.colContatos = New Collection
        Set objVou.colPassageiros = New Collection
        Set objVou.colVendedores = New Collection
    
        objCliente.sNomeReduzido = Cliente.Text
    
        'Lê o Cliente através do Nome Reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 200492
        
        objVou.lClienteVou = objCliente.lCodigo
        objVou.lClienteComissao = objCliente.lCodigo
        objVou.iGrupo = Grupo.ItemData(Grupo.ListIndex)
        
        objEmi.lCodigo = Emissor.ItemData(Emissor.ListIndex)
        
        lErro = CF("TRPEmissores_Le", objEmi)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 200493
        
        objVou.lEmissor = objEmi.lFornEmissor
        objVou.iDestino = Codigo_Extrai(Destino.Text)
        
        lErro = CF("Produto_Formata", SCodigo_Extrai(Produto.Text), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 200494
    
        objVou.sProduto = sProdutoFormatado
        
        objVou.dtDataVigenciaDe = StrParaDate(VigenciaDe.Text)
        objVou.dtDataVigenciaAte = StrParaDate(VigenciaAte.Text)
        objVou.iVigencia = StrParaInt(Vigencia.Caption)
        objVou.iMoeda = Codigo_Extrai(Moeda.Text)
        'objVou.iPax = StrParaInt(QtdePax.Text)
        objVou.dTarifaUnitaria = StrParaDbl(TarifaUnitaria.Text)
        objVou.dTarifaUnitariaFolheto = StrParaDbl(TarifaFolheto.Caption)
        objVou.dCambio = StrParaDbl(Cambio.Text)
        objVou.sObservacao = Observacao.Text
        
        objVou.sUsuarioLibManut = UsuarioLib.Caption
        objVou.dtDataLibManut = StrParaDate(DataLib.Caption)
        If HoraLib.Caption <> "" Then objVou.dHoraLibManut = CDbl(StrParaDate(HoraLib.Caption))
        
        objVou.iPax = 0
        iSeq = 1
        For iIndice = 1 To objGridPax.iLinhasExistentes
        
            Set objPax = New ClassTRPVouPassageiros
            
            objPax.iTitular = StrParaInt(GridPax.TextMatrix(iIndice, iGrid_PaxTitular_Col))
            
            If StrParaInt(GridPax.TextMatrix(iIndice, iGrid_PaxCancelado_Col)) = MARCADO Then
                objPax.iStatus = STATUS_TRP_VOU_CANCELADO
            Else
                objPax.iStatus = STATUS_TRP_VOU_ABERTO
                objVou.iPax = objVou.iPax + 1
            End If
            
            objPax.dtDataNascimento = StrParaDate(GridPax.TextMatrix(iIndice, iGrid_PaxDataNasc_Col))
            objPax.sPrimeiroNome = GridPax.TextMatrix(iIndice, iGrid_PaxNome_Col)
            objPax.sSobreNome = GridPax.TextMatrix(iIndice, iGrid_PaxSobreNome_Col)
            objPax.sNome = objPax.sPrimeiroNome & " " & objPax.sSobreNome
            objPax.sNumeroDocumento = GridPax.TextMatrix(iIndice, iGrid_PaxNumDoc_Col)
            objPax.sSexo = GridPax.TextMatrix(iIndice, iGrid_PaxSexo_Col)
            objPax.sTipoDocumento = GridPax.TextMatrix(iIndice, iGrid_PaxTipoDoc_Col)
            objPax.dValorPago = StrParaDbl(GridPax.TextMatrix(iIndice, iGrid_PaxValor_Col))
            objPax.dValorPagoEmi = StrParaDbl(GridPax.TextMatrix(iIndice, iGrid_PaxValorEmi_Col))
            
            If objPax.iTitular = MARCADO Then
                If objPax.sTipoDocumento = "CPF" Then
                    objVou.sPassageiroCGC = objPax.sNumeroDocumento
                Else
                    objVou.sPassageiroCGC = ""
                End If
                objVou.sPassageiroNome = objPax.sPrimeiroNome
                objVou.sPassageiroSobreNome = objPax.sSobreNome
                objVou.dtPassageiroDataNasc = objPax.dtDataNascimento
                
                If objPax.iStatus = STATUS_TRP_VOU_CANCELADO Then gError 209070
                
                objPax.iSeq = 1
            Else
                iSeq = iSeq + 1
                objPax.iSeq = iSeq
            End If
        
            objVou.colPassageiros.Add objPax
        
        Next
        
        objVou.objEnderecoPax.sBairro = Bairro.Text
        objVou.objEnderecoPax.sCEP = CEP.Text
        objVou.objEnderecoPax.sCidade = Cidade.Text
        objVou.objEnderecoPax.sContato = Contato.Text
        objVou.objEnderecoPax.sEmail = Email.Text
        objVou.objEnderecoPax.sEndereco = Endereco.Text
        objVou.objEnderecoPax.sSiglaEstado = Estado.Text
        objVou.objEnderecoPax.sTelefone1 = Telefone1.Text
        objVou.objEnderecoPax.sTelefone2 = Telefone2.Text
    
        For iIndice = 1 To objGridContatos.iLinhasExistentes
        
            Set objContato = New ClassTRPVouContatos
            
            objContato.sNome = GridContatos.TextMatrix(iIndice, iGrid_ContatoNome_Col)
            objContato.sTelefone = GridContatos.TextMatrix(iIndice, iGrid_ContatoTelefone_Col)
        
            objContato.iSeq = iIndice
        
            objVou.colContatos.Add objContato
        
        Next
        
        If OptCCSim.Value Then
            objVou.iCartao = MARCADO
            objVou.sCiaCart = BandeiraCC.Text
            objVou.iQuantParc = StrParaInt(NumParcelasCC.Text)
            objVou.sNumCCred = NumeroCC.Text
            objVou.sTitular = TitularCC.Text
            objVou.sTitularCPF = Trim(CPFTitularCC.ClipText)
            objVou.sValidadeCC = ValidadeCC.Text
            objVou.lNumAuto = StrParaLong(NumAutoCC.Text)
            objVou.dtDataAutoCC = StrParaDate(DataAutoCC.Text)
            objVou.iCodSegurancaCC = StrParaInt(CodSegurancaCC.Text)
        Else
            objVou.iCartao = DESMARCADO
            objVou.sCiaCart = ""
            objVou.iQuantParc = 0
            objVou.sNumCCred = ""
            objVou.sTitular = ""
            objVou.sTitularCPF = ""
            objVou.sValidadeCC = ""
            objVou.lNumAuto = 0
            objVou.dtDataAutoCC = DATA_NULA
            objVou.iCodSegurancaCC = 0
        End If
        
        objVou.lRepresentante = LCodigo_Extrai(RepresentanteComissao.Text)
        objVou.lCorrentista = LCodigo_Extrai(CorrentistaComissao.Text)
        objVou.iPromotor = Codigo_Extrai(PromotorComissao.Text)

        For iIndice = 1 To objGridVendedores.iLinhasExistentes
        
            Set objVendedor = New ClassVendedor
            
            objVendedor.iCodigo = Codigo_Extrai(GridVendedores.TextMatrix(iIndice, iGrid_Vendedor_Col))
            objVendedor.dPercComissao = PercentParaDbl(GridVendedores.TextMatrix(iIndice, iGrid_VendedorPerc_Col))
        
            objVou.colVendedores.Add objVendedor
        
        Next
        
        objVou.dComissaoAg = StrParaDbl(PercComiAgeNovo.Text) / 100
        objVou.dComissaoRep = StrParaDbl(PercComiRepNovo.Text) / 100
        objVou.dComissaoCorr = StrParaDbl(PercComiCorNovo.Text) / 100
        objVou.dComissaoEmissor = StrParaDbl(PercComiEmiNovo.Text) / 100
        objVou.dComissaoProm = StrParaDbl(PercComiProNovo.Text) / 100
        
        objVou.dValorBruto = StrParaDbl(BrutoRS.Caption)
        objVou.dValorCambio = StrParaDbl(BrutoNaMoeda.Caption)
        objVou.dValor = StrParaDbl(ValorFaturavelRS.Caption)
        
        If Idioma.ListIndex <> -1 Then
            objVou.iIdioma = Idioma.ItemData(Idioma.ListIndex)
        End If
        
        If ImprimirValor.Value = vbChecked Then
            objVou.iImprimirValor = MARCADO
        Else
            objVou.iImprimirValor = DESMARCADO
        End If
           
    End If

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 200488 To 200490 'ERRO_TRP_NUMVOU_NAO_PREENCHIDO
            Call Rotina_Erro(vbOKOnly, "ERRO_TRP_NUMVOU_NAO_PREENCHIDO", gErr)

        Case 200491 To 200494
        
        Case 200496
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_NAO_CADASTRADO", gErr, objVou.lNumVou, objVou.sSerie, objVou.sTipVou)

        Case 209070
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_TITULAR_CANCELADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200495)

    End Select

    Exit Function

End Function

Function Traz_TRPVouchers_Tela(objVou As ClassTRPVouchers) As Long

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objFornecedor As New ClassFornecedor
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim objVendedor As New ClassVendedor
Dim sSiglaDest As String
Dim sDescricaoDest As String
Dim objEmi As New ClassCliEmissoresTRP
Dim iLinha As Integer
Dim objPax As ClassTRPVouPassageiros
Dim objContato As ClassTRPVouContatos
Dim iIndice As Integer

On Error GoTo Erro_Traz_TRPVouchers_Tela

    Call Limpa_Tela_TRPVouchers

    gbTrazendodados = True

    NumeroVouP.PromptInclude = False
    NumeroVouP.Text = CStr(objVou.lNumVou)
    NumeroVouP.PromptInclude = True
    SerieVouP.Text = objVou.sSerie
    TipoVouP.Text = objVou.sTipVou
    
    'Lê o TRPVouchers que está sendo Passado
    lErro = CF("TRPVouchers_Le", objVou)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 200497

    If lErro <> SUCESSO Then gError 200498 'Voucher não cadastrado
    
    'If objVou.dtDataVigenciaDe <= Date Then gError 200499 'Em Vigência
    If objVou.iStatus = STATUS_TRP_VOU_CANCELADO Then gError 200500 'Voucher cancelado
    If objVou.iGeraComissao = MARCADO Then gError 200501 'Voucher sem comissão gerada
    'If objVou.lNumFat <> 0 Then gError 200502  'Voucher Faturado
        
    TipoVou.Caption = objVou.sTipVou
    SerieVou.Caption = objVou.sSerie
    NumeroVou.Caption = objVou.lNumVou
    DataEmissaoVou.Caption = Format(objVou.dtData, "dd/mm/yyyy")
    
    Call Combo_Seleciona_ItemData(Grupo, objVou.iGrupo)
    Call Grupo_Click
    
    objCliente.lCodigo = objVou.lClienteVou
    
    lErro = CF("Cliente_Le", objCliente)
    If lErro <> SUCESSO And lErro <> 12293 Then gError 200503
     
    Cliente.Text = objVou.lClienteVou
    Call Cliente_Validate(bSGECancelDummy)
    
    ClienteComissao.Caption = Cliente.Text
    
    objEmi.lFornEmissor = objVou.lEmissor
    
    lErro = CF("TRPEmissores_Le_Forn", objEmi)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 200504
       
'    lErro = LCombo_Seleciona(Emissor, objEmi.lCodigo)
'    If lErro <> SUCESSO Then gError 200505

    For iIndice = 0 To Emissor.ListCount
        If LCodigo_Extrai(Emissor.List(iIndice)) = objEmi.lCodigo Then
            Emissor.ListIndex = iIndice
            Exit For
        End If
    Next
    
    EmissorComissao.Caption = Emissor.Text
    
    Call Combo_Seleciona_ItemData(Destino, objVou.iDestino)
    iDestinoAnt = 0
    Call Destino_Change
     
    lErro = CF("SCombo_Seleciona2", Produto, Trim(objVou.sProduto))
    If lErro <> SUCESSO Then gError 200506
    Call Produto_Click
     
    VigenciaDe.PromptInclude = False
    VigenciaDe.Text = Format(objVou.dtDataVigenciaDe, "dd/mm/yy")
    VigenciaDe.PromptInclude = True
    Call VigenciaDe_Validate(bSGECancelDummy)
     
    VigenciaAte.PromptInclude = False
    VigenciaAte.Text = Format(objVou.dtDataVigenciaAte, "dd/mm/yy")
    VigenciaAte.PromptInclude = True
    Call VigenciaAte_Validate(bSGECancelDummy)
     
    Call Combo_Seleciona_ItemData(Moeda, objVou.iMoeda)
     
    QtdePax.PromptInclude = False
    If objVou.iPax <> 0 Then
        QtdePax.Text = objVou.iPax
    Else
        QtdePax.Text = ""
    End If
    QtdePax.PromptInclude = True
    
    TarifaFolheto.Caption = Format(objVou.dTarifaUnitariaFolheto, "STANDARD")
    
    TarifaUnitaria.Text = Format(objVou.dTarifaUnitaria, "STANDARD")
    
    Cambio.PromptInclude = False
    If objVou.dCambio <> 0 Then
        Cambio.Text = Format(objVou.dCambio, Cambio.Format)
    Else
        Cambio.Text = ""
    End If
    Cambio.PromptInclude = True
    
    Call Grid_Limpa(objGridPax)
    iLinha = 0
    For Each objPax In objVou.colPassageiros
    
        iLinha = iLinha + 1
        
        If objPax.iStatus = STATUS_TRP_VOU_CANCELADO Then
            GridPax.TextMatrix(iLinha, iGrid_PaxCancelado_Col) = MARCADO
        Else
            GridPax.TextMatrix(iLinha, iGrid_PaxCancelado_Col) = DESMARCADO
        End If
        
        GridPax.TextMatrix(iLinha, iGrid_PaxTitular_Col) = objPax.iTitular
        GridPax.TextMatrix(iLinha, iGrid_PaxNome_Col) = objPax.sPrimeiroNome
        GridPax.TextMatrix(iLinha, iGrid_PaxSobreNome_Col) = objPax.sSobreNome
        GridPax.TextMatrix(iLinha, iGrid_PaxDataNasc_Col) = Format(objPax.dtDataNascimento, "dd/mm/yyyy")
        GridPax.TextMatrix(iLinha, iGrid_PaxTipoDoc_Col) = objPax.sTipoDocumento
        GridPax.TextMatrix(iLinha, iGrid_PaxNumDoc_Col) = objPax.sNumeroDocumento
        GridPax.TextMatrix(iLinha, iGrid_PaxSexo_Col) = objPax.sSexo
        GridPax.TextMatrix(iLinha, iGrid_PaxValor_Col) = Format(objPax.dValorPago, "STANDARD")
        GridPax.TextMatrix(iLinha, iGrid_PaxValorEmi_Col) = Format(objPax.dValorPagoEmi, "STANDARD")
       
    Next
    
    objGridPax.iLinhasExistentes = iLinha
    
    Call Grid_Refresh_Checkbox(objGridPax)
    
    Call Calcula_SomaPax
    
    BrutoPax.Caption = Format(objVou.dValorBruto, "STANDARD")
    
    Endereco.Text = objVou.objEnderecoPax.sEndereco
    Bairro.Text = objVou.objEnderecoPax.sBairro
    Cidade.Text = objVou.objEnderecoPax.sCidade
    CEP.Text = objVou.objEnderecoPax.sCEP
    Estado.Text = objVou.objEnderecoPax.sSiglaEstado
    Call Estado_Validate(bSGECancelDummy)

    If objVou.objEnderecoPax.iCodigoPais = 0 Then
        Pais.Text = ""
    Else
        Pais.Text = objVou.objEnderecoPax.iCodigoPais
        Call Pais_Validate(bSGECancelDummy)
    End If
    
    Telefone1.Text = objVou.objEnderecoPax.sTelefone1
    Telefone2.Text = objVou.objEnderecoPax.sTelefone2
    Email.Text = objVou.objEnderecoPax.sEmail
    Contato.Text = objVou.objEnderecoPax.sContato
    
    Call Grid_Limpa(objGridContatos)
    iLinha = 0
    For Each objContato In objVou.colContatos
    
        iLinha = iLinha + 1
        
        GridContatos.TextMatrix(iLinha, iGrid_ContatoNome_Col) = objContato.sNome
        GridContatos.TextMatrix(iLinha, iGrid_ContatoTelefone_Col) = objContato.sTelefone
       
    Next
    
    objGridContatos.iLinhasExistentes = iLinha
    
    If objVou.iCartao = MARCADO Then
        OptCCSim.Value = True
        FrameCC.Enabled = True
    
        TitularCC.Text = objVou.sTitular
    
        CPFTitularCC.PromptInclude = False
        CPFTitularCC.Text = objVou.sTitularCPF
        CPFTitularCC.PromptInclude = True
        Call CPFTitularCC_Validate(bSGECancelDummy)
        
        Call CF("SCombo_Seleciona2", BandeiraCC, objVou.sCiaCart)
        'If lErro <> SUCESSO Then gError 200507
        
        Call Combo_Seleciona_ItemData(NumParcelasCC, objVou.iQuantParc)
        
        NumeroCC.PromptInclude = False
        NumeroCC.Text = Format(objVou.sNumCCred, NumeroCC.Format)
        NumeroCC.PromptInclude = True
    
        ValidadeCC.PromptInclude = False
        ValidadeCC.Text = objVou.sValidadeCC
        ValidadeCC.PromptInclude = True
    
        NumAutoCC.PromptInclude = False
        If objVou.lNumAuto <> 0 Then
            NumAutoCC.Text = CStr(objVou.lNumAuto)
        Else
            NumAutoCC.Text = ""
        End If
        NumAutoCC.PromptInclude = True
    
        DataAutoCC.PromptInclude = False
        If objVou.dtDataAutoCC <> DATA_NULA Then
            DataAutoCC.Text = Format(objVou.dtDataAutoCC, "dd/mm/yy")
        Else
            DataAutoCC.Text = ""
        End If
        DataAutoCC.PromptInclude = True
    
        If objVou.iCodSegurancaCC <> 0 Then
            CodSegurancaCC.PromptInclude = False
            CodSegurancaCC.Text = CStr(objVou.iCodSegurancaCC)
            CodSegurancaCC.PromptInclude = True
        End If
    Else
        Call Limpa_FrameCC
    End If
    
    If objVou.lRepresentante <> 0 Then
        RepresentanteComissao.Text = objVou.lRepresentante
        Call RepresentanteComissao_Validate(bSGECancelDummy)
    Else
        RepresentanteComissao.Text = ""
        Call RepresentanteComissao_Validate(bSGECancelDummy)
    End If
    
    If objVou.lCorrentista <> 0 Then
        CorrentistaComissao.Text = objVou.lCorrentista
        Call CorrentistaComissao_Validate(bSGECancelDummy)
    Else
        CorrentistaComissao.Text = ""
        Call CorrentistaComissao_Validate(bSGECancelDummy)
    End If
       
    PercComiRepNovo.Text = objVou.dComissaoRep * 100
    Call PercComiRepNovo_Validate(bSGECancelDummy)
        
    PercComiCorNovo.Text = objVou.dComissaoCorr * 100
    Call PercComiCorNovo_Validate(bSGECancelDummy)
        
    PercComiEmiNovo.Text = objVou.dComissaoEmissor * 100
    Call PercComiEmiNovo_Validate(bSGECancelDummy)
    
    PercComiAgeNovo.Text = objVou.dComissaoAg * 100
    Call PercComiAgeNovo_Validate(bSGECancelDummy)
    
    PercComiProNovo.Text = objVou.dComissaoProm * 100
    Call PercComiProNovo_Validate(bSGECancelDummy)
    
    BrutoNaMoeda.Caption = Format(objVou.dValorCambio, "STANDARD")
    BrutoRS.Caption = Format(objVou.dValorBruto, "STANDARD")
    ValorFaturavelRS.Caption = Format(objVou.dValor, "STANDARD")
    
    Controle.Caption = objVou.sControle
     
    objCondicaoPagto.iCodigo = objCliente.iCondicaoPagto

    'Lê Condição Pagamento no BD
    lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
    If lErro <> SUCESSO And lErro <> 19205 Then gError 200508
     
    CondPagto.Caption = objCliente.iCondicaoPagto & SEPARADOR & objCondicaoPagto.sDescReduzida
     
    Call Combo_Seleciona_ItemData(Idioma, objVou.iIdioma)
     
    Call Grid_Limpa(objGridVendedores)
    iLinha = 0
    For Each objVendedor In objVou.colVendedores
    
        iLinha = iLinha + 1
        
        PromotorComissao.Text = objVendedor.iCodigo
        Call PromotorComissao_Validate(bSGECancelDummy)
        
        GridVendedores.TextMatrix(iLinha, iGrid_Vendedor_Col) = PromotorComissao.Text
        GridVendedores.TextMatrix(iLinha, iGrid_VendedorPerc_Col) = Format(objVendedor.dPercComissao, "Percent")
       
    Next

    objGridVendedores.iLinhasExistentes = iLinha

    If objVou.iPromotor <> 0 Then
        PromotorComissao.Text = objVou.iPromotor
        Call PromotorComissao_Validate(bSGECancelDummy)
    Else
        PromotorComissao.Text = ""
        Call PromotorComissao_Validate(bSGECancelDummy)
    End If
    
    If objVou.iImprimirValor = MARCADO Then
        ImprimirValor.Value = vbChecked
    Else
        ImprimirValor.Value = vbUnchecked
    End If
       
    Call Calcula_Valores
    
    Observacao.Text = objVou.sObservacao
       
    Set gobjVoucher = objVou
    
    If objVou.lNumFat = 0 Then
        Faturado.ForeColor = vbBlack
        Faturado.Caption = "NÃO"
    Else
        Faturado.ForeColor = vbRed
        Faturado.Caption = "SIM"
    End If
       
    If Date <= objVou.dtDataVigenciaDe Then
        EmVigencia.ForeColor = vbBlack
        EmVigencia.Caption = "NÃO"
    Else
        EmVigencia.ForeColor = vbRed
        EmVigencia.Caption = "SIM"
    End If
    
    If objVou.lNumFat <> 0 Or objVou.dtDataVigenciaDe < Date Then
    
        'Call Rotina_Aviso(vbOKOnly, "AVISO_MANU_VOU_FAT_OU_EM_VIG")
    
        objVou.sUsuarioLibManut = ""
        objVou.dtDataLibManut = DATA_NULA
        objVou.dHoraLibManut = 0
    
        Load TRPAcessoLiberaVou
    
        lErro = TRPAcessoLiberaVou.Trata_Parametros(objVou)
        If lErro <> SUCESSO Then gError 129290
    
        TRPAcessoLiberaVou.Show vbModal
    
        If objVou.dtDataLibManut = DATA_NULA Then
            gbVoucherBloqueado = True
            gbVoucherValorBloqueado = True
            Call Bloqueia_Funcionalidades(True)
        Else
            DataLib.Caption = Format(objVou.dtDataLibManut, "dd/mm/yyyy")
            UsuarioLib.Caption = objVou.sUsuarioLibManut
            HoraLib.Caption = Format(objVou.dHoraLibManut, "hh:mm:ss")
            If objVou.lNumFat <> 0 Then
                gbVoucherValorBloqueado = True
            Else
                gbVoucherValorBloqueado = False
            End If
            gbVoucherBloqueado = False
            Call Bloqueia_Funcionalidades(False)
        End If
    Else
        gbVoucherValorBloqueado = False
        gbVoucherBloqueado = False
        Call Bloqueia_Funcionalidades(False)
    End If
    
    gbTrazendodados = False

    iAlterado = 0

    Traz_TRPVouchers_Tela = SUCESSO

    Exit Function

Erro_Traz_TRPVouchers_Tela:

    gbTrazendodados = False

    Traz_TRPVouchers_Tela = gErr

    Select Case gErr

        Case 200497
        
        Case 200498 'Não cadastrado
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_NAO_CADASTRADO", gErr)
        
        Case 200499 'Em vigência
            Call Rotina_Erro(vbOKOnly, "ERRO_MANUT_VOU_EM_VIG", gErr)
        
        Case 200500 'Cancelado
            Call Rotina_Erro(vbOKOnly, "ERRO_MANUT_VOU_CANCELADO", gErr)
        
        Case 200501 'Sem comissão
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_SEM_COMISSAO", gErr)

        Case 200502 'Faturado
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_JA_FATURADO", gErr)
            
        Case 200503 To 200508

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200509)

    End Select

    Exit Function

End Function

Private Function Calcula_Vigencia() As Long

On Error GoTo Erro_Calcula_Vigencia

    If StrParaDate(VigenciaDe.Text) <> DATA_NULA And StrParaDate(VigenciaAte.Text) <> DATA_NULA Then
        Vigencia.Caption = DateDiff("d", StrParaDate(VigenciaDe.Text), StrParaDate(VigenciaAte.Text)) + 1
    
        Call Calcula_Tarifa
    End If

    Calcula_Vigencia = SUCESSO

    Exit Function

Erro_Calcula_Vigencia:

    Calcula_Vigencia = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200510)

    End Select

    Exit Function
    
End Function

Private Function Calcula_Tarifa() As Long

Dim lErro As Long
Dim sProduto As String
Dim iVigencia As Integer
Dim objProdCalc As New ClassTRPProdTafCalc
Dim colTarifa As New Collection
Dim colTarifaAcordo As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim dTarifa As Double
Dim iDestino As Integer
Dim objCliente As New ClassCliente
Dim objProd As New ClassProduto
Dim objProdTRP As ClassProdutoTRP

On Error GoTo Erro_Calcula_Tarifa

    If Not gbVoucherValorBloqueado Then

        sProduto = SCodigo_Extrai(Produto.Text)
        If StrParaDbl(Vigencia.Caption) > 0 Then
            iVigencia = StrParaInt(Vigencia.Caption)
        End If
        iDestino = Codigo_Extrai(Destino.Text)
        
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 200511
        
        objProd.sCodigo = sProdutoFormatado
    
        lErro = CF("Produto_Le_Customizado", objProd)
        If lErro <> SUCESSO Then gError 200511
        
        Set objProdTRP = objProd.objInfoUsu
        
        If sProdutoAnt <> sProduto Then
            If objProdTRP.iVigenciaFixa > 0 Then
                If Not gbTrazendodados And Not gbVoucherBloqueado Then
                    VigenciaAte.PromptInclude = False
                    VigenciaAte.Text = Format(DateAdd("d", objProdTRP.iVigenciaFixa - 1, StrParaDate(VigenciaDe.Text)), "dd/mm/yy")
                    VigenciaAte.PromptInclude = True
                    Vigencia.Caption = CStr(objProdTRP.iVigenciaFixa)
                    iVigencia = objProdTRP.iVigenciaFixa
                End If
                VigenciaAte.Enabled = False
                UpDownVigenciaAte.Enabled = False
            Else
                VigenciaAte.Enabled = True
                UpDownVigenciaAte.Enabled = True
            End If
        End If
        
        If Not gbTrazendodados And Not gbVoucherBloqueado Then
            
            If sProdutoAnt <> sProduto Or iVigencia <> iVigenciaAnt And iVigencia > 0 Or iDestinoAnt <> iDestino Or sClienteAnt <> Cliente.Text Then
            
            
                objCliente.sNomeReduzido = Cliente.Text
            
                'Lê o Cliente através do Nome Reduzido
                lErro = CF("Cliente_Le_NomeReduzido", objCliente)
                If lErro <> SUCESSO And lErro <> 12348 Then gError 200511
            
                lErro = CF("TRPProdTarifa_Le", colTarifa, sProdutoFormatado)
                If lErro <> SUCESSO Then gError 200512
                
                lErro = CF("TRPAcordoTarifa_Le", colTarifaAcordo, sProdutoFormatado, iDestino, objCliente.lCodigo, StrParaDate(DataEmissaoVou.Caption))
                If lErro <> SUCESSO Then gError 200512
                
                If colTarifa.Count <> 0 Then
                
                    Set objProdCalc = colTarifa.Item(1)
                    
                End If
                    
                Call objProdCalc.Calcula_Tarifa(StrParaDate(DataEmissaoVou.Caption), iVigencia, dTarifa)
                    
                TarifaFolheto.Caption = Format(dTarifa, "STANDARD")
                    
                Set objProdCalc.colTarifaAcordo = colTarifaAcordo
                    
                Call objProdCalc.Calcula_Tarifa(StrParaDate(DataEmissaoVou.Caption), iVigencia, dTarifa)
            
                TarifaUnitaria.Text = Format(dTarifa, "STANDARD")
            
                sProdutoAnt = sProduto
                iVigenciaAnt = iVigencia
            End If
            
            Call BotaoRecalcularPax_Click
            
        End If
        
    End If

    Calcula_Tarifa = SUCESSO

    Exit Function

Erro_Calcula_Tarifa:

    Calcula_Tarifa = gErr

    Select Case gErr
    
        Case 200511, 200512

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200513)

    End Select

    Exit Function
    
End Function

Private Function Calcula_ValoresEmReal() As Long

On Error GoTo Erro_Calcula_ValoresEmReal

    If Not gbTrazendodados And Not gbVoucherBloqueado And Not gbVoucherValorBloqueado Then
    
        BrutoPax.Caption = Format(StrParaDbl(BrutoUSSPax.Caption) * StrParaDbl(Cambio.Text), "STANDARD")
        BrutoRS.Caption = Format(StrParaDbl(BrutoNaMoeda.Caption) * StrParaDbl(Cambio.Text), "STANDARD")
    
        Call Calcula_Valores
    
    End If

    Calcula_ValoresEmReal = SUCESSO

    Exit Function

Erro_Calcula_ValoresEmReal:

    Calcula_ValoresEmReal = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200514)

    End Select

    Exit Function
    
End Function

Private Function Calcula_Valores() As Long

Dim dBruto As Double

On Error GoTo Erro_Calcula_Valores

    dBruto = StrParaDbl(BrutoRS.Caption)

    VlrComiCor.Caption = Format(dBruto * StrParaDbl(PercComiCorNovo.Text) / 100, "STANDARD")
    VlrComiAge.Caption = Format(dBruto * StrParaDbl(PercComiAgeNovo.Text) / 100, "STANDARD")
    VlrComiRep.Caption = Format(dBruto * StrParaDbl(PercComiRepNovo.Text) / 100, "STANDARD")
    VlrComiEmi.Caption = Format(dBruto * StrParaDbl(PercComiEmiNovo.Text) / 100, "STANDARD")
        
    If Not gbTrazendodados And Not gbVoucherBloqueado And Not gbVoucherValorBloqueado Then
        ValorFaturavelRS.Caption = Format(dBruto - IIf(OptCCSim.Value, 0, StrParaDbl(VlrComiAge.Caption)), "STANDARD")
    End If
    
    Calcula_Valores = SUCESSO

    Exit Function

Erro_Calcula_Valores:

    Calcula_Valores = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200515)

    End Select

    Exit Function
    
End Function

Private Function Calcula_Cambio() As Long

Dim lErro As Long
Dim dBruto As Double
Dim iMoeda As Integer
Dim dCambio As Double

On Error GoTo Erro_Calcula_Cambio

    iMoeda = Codigo_Extrai(Moeda.Text)

    If Not gbTrazendodados And Not gbVoucherBloqueado And Not gbVoucherValorBloqueado Then
        
        If iMoeda <> iMoedaAnt Then
        
            If iMoeda = MOEDA_REAL Then
                dCambio = 1
            Else
                lErro = CF("TRPCambio_Le", iMoeda, StrParaDate(DataEmissaoVou.Caption), dCambio)
                If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 200516
            End If
        End If
        
        Cambio.Text = Format(dCambio, "#0.0000")
        
        Call Calcula_ValoresEmReal
       
    End If
    
    iMoedaAnt = iMoeda
    
    Calcula_Cambio = SUCESSO

    Exit Function

Erro_Calcula_Cambio:

    Calcula_Cambio = gErr

    Select Case gErr
    
        Case 200516

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200517)

    End Select

    Exit Function
    
End Function

Private Function Calcula_SomaPax() As Long

Dim dBruto As Double
Dim dBrutoEmi As Double
Dim iIndice As Integer
Dim iPax As Integer

On Error GoTo Erro_Calcula_SomaPax

    iPax = 0
    For iIndice = 1 To objGridPax.iLinhasExistentes
        If StrParaInt(GridPax.TextMatrix(iIndice, iGrid_PaxCancelado_Col)) = DESMARCADO Then
            iPax = iPax + 1
        End If
    Next
    QtdePax.PromptInclude = False
    QtdePax.Text = CStr(iPax)
    QtdePax.PromptInclude = True
    
    If Not gbVoucherValorBloqueado And Not gbVoucherBloqueado Then
    
        For iIndice = 1 To objGridPax.iLinhasExistentes
         
            If StrParaInt(GridPax.TextMatrix(iIndice, iGrid_PaxCancelado_Col)) = DESMARCADO Then
               dBruto = dBruto + StrParaDbl(GridPax.TextMatrix(iIndice, iGrid_PaxValor_Col))
            End If
             
            dBrutoEmi = dBrutoEmi + StrParaDbl(GridPax.TextMatrix(iIndice, iGrid_PaxValorEmi_Col))
        
        Next
         
        BrutoUSSPaxEmi.Caption = Format(dBrutoEmi, "STANDARD")
        BrutoUSSPax.Caption = Format(dBruto, "STANDARD")
        BrutoNaMoeda.Caption = Format(dBruto, "STANDARD")
        
        Call Calcula_ValoresEmReal
        
    End If
        
    Calcula_SomaPax = SUCESSO
    
    Exit Function

Erro_Calcula_SomaPax:

    Calcula_SomaPax = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200518)

    End Select

    Exit Function
    
End Function

Private Sub BotaoRecalcularPax_Click()
    Call Calcula_Dados_Pax
End Sub

Private Function Calcula_Dados_Pax(Optional ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objProd As New ClassProduto
Dim objProdTRP As New ClassProdutoTRP
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iGrupo As Integer
Dim iIdade As Integer
Dim dValorPax As Double
Dim iNumFamiComDesc As Integer
Dim dtDataNasc As Date
Dim dtDataBase As Date
Dim dtDataNiver As Date

On Error GoTo Erro_Calcula_Dados_Pax

    If Not gbVoucherValorBloqueado And Not gbVoucherValorBloqueado Then

        lErro = CF("Produto_Formata", SCodigo_Extrai(Produto.Text), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 200519
        
        objProd.sCodigo = sProdutoFormatado
    
        lErro = CF("Produto_Le_Customizado", objProd)
        If lErro <> SUCESSO Then gError 200520
        
        If Grupo.ListIndex <> -1 Then
            iGrupo = Grupo.ItemData(Grupo.ListIndex)
        End If
        
        Set objProdTRP = objProd.objInfoUsu
    
        For iIndice = 1 To objGridPax.iLinhasExistentes
        
            If iLinha = 0 Or iLinha = iIndice Then
        
                dValorPax = StrParaDbl(TarifaUnitaria.Text)
                iNumFamiComDesc = 0
                
                dtDataNasc = StrParaDate(GridPax.TextMatrix(iIndice, iGrid_PaxDataNasc_Col))
                dtDataBase = StrParaDate(VigenciaDe.Text)
                dtDataNiver = DateAdd("yyyy", DateDiff("yyyy", dtDataNasc, dtDataBase), dtDataNasc)
                
                iIdade = Year(dtDataBase) - Year(dtDataNasc) - IIf(dtDataBase <= dtDataNiver, 1, 0)
            
                If iGrupo = TRP_GRUPO_FAMILIA Then
                
                    If StrParaInt(GridPax.TextMatrix(iIndice, iGrid_PaxTitular_Col)) = MARCADO Then
                        dValorPax = dValorPax * (1 - objProdTRP.dDescontoTitular)
                    ElseIf iNumFamiComDesc < objProdTRP.iMaxAcompanhante Then
                        iNumFamiComDesc = 1 + iNumFamiComDesc
                        dValorPax = dValorPax * (1 - objProdTRP.dDescontoAcompanhante)
                    Else
                        dValorPax = dValorPax
                    End If
                
                End If
                    
                If iIdade >= objProdTRP.iLimiteIdade Then
                    
                    If objProdTRP.iAceitaIdadeMaiorLimite = DESMARCADO Then gError 200521
                    
                    dValorPax = dValorPax * (1 + objProdTRP.dPercAcrescIdadeMaiorLimite)
                    
                End If
                    
                GridPax.TextMatrix(iIndice, iGrid_PaxValor_Col) = Format(dValorPax, "STANDARD")
                
            End If
                
        Next
        
    End If
        
    Call Calcula_SomaPax
    
    Calcula_Dados_Pax = SUCESSO
    
    Exit Function

Erro_Calcula_Dados_Pax:

    Calcula_Dados_Pax = gErr

    Select Case gErr
    
        Case 200519, 200520
        
        Case 200521
            Call Rotina_Erro(vbOKOnly, "ERRO_PAX_GRID_SUPERIOR_IDADE_LIMITE", gErr, iIdade, objProdTRP.iLimiteIdade)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200522)

    End Select

    Exit Function
    
End Function

Private Function Valida_Dados(ByVal objVou As ClassTRPVouchers) As Long

Dim lErro As Long
Dim objProd As New ClassProduto
Dim objProdTRP As ClassProdutoTRP
Dim dtDataNasc As Date
Dim dtDataBase As Date
Dim dtDataNiver As Date
Dim iIdade As Integer
Dim objPax As ClassTRPVouPassageiros

On Error GoTo Erro_Valida_Dados

    objProd.sCodigo = objVou.sProduto

    lErro = CF("Produto_Le_Customizado", objProd)
    If lErro <> SUCESSO Then gError 200523
    
    Set objProdTRP = objProd.objInfoUsu
    
    If objProdTRP.iAceitaIdadeMaiorLimite = DESMARCADO Then
    
        For Each objPax In objVou.colPassageiros
    
            dtDataNasc = objPax.dtDataNascimento
            dtDataBase = objVou.dtDataVigenciaDe
            dtDataNiver = DateAdd("yyyy", DateDiff("yyyy", dtDataNasc, dtDataBase), dtDataNasc)
            
            iIdade = Year(dtDataBase) - Year(dtDataNasc) - IIf(dtDataBase <= dtDataNiver, 1, 0)
        
            If iIdade > objProdTRP.iLimiteIdade Then gError 200524
        
        Next
    
    End If
    
    If objVou.iCartao = MARCADO Then
    
        If Len(Trim(objVou.sTitular)) = 0 Then gError 200525
        If Len(Trim(objVou.sTitularCPF)) = 0 Then gError 200526
        If Len(Trim(objVou.sValidadeCC)) = 0 Then gError 200527
        If Len(Trim(objVou.sNumCCred)) = 0 Then gError 200528
        If Len(Trim(objVou.sCiaCart)) = 0 Then gError 200529
        If objVou.iQuantParc = 0 Then gError 200530
        
        If objVou.lNumAuto <> 0 Then
            If objVou.dtDataAutoCC = DATA_NULA Then gError 200625
        End If
    
    End If
    
    If objVou.dComissaoCorr > 0 And objVou.lCorrentista = 0 Then gError 200645
    If objVou.dComissaoRep > 0 And objVou.lRepresentante = 0 Then gError 200646
    
    If objVou.dtDataVigenciaDe > objVou.dtDataVigenciaAte Then gError 200626
    
    Valida_Dados = SUCESSO
    
    Exit Function

Erro_Valida_Dados:

    Valida_Dados = gErr

    Select Case gErr
    
        Case 200523
        
        Case 200524
            Call Rotina_Erro(vbOKOnly, "ERRO_PAX_GRID_SUPERIOR_IDADE_LIMITE", gErr, iIdade, objProdTRP.iLimiteIdade)

        Case 200525
            Call Rotina_Erro(vbOKOnly, "ERRO_VOU_CC_SEM_TITULAR", gErr)

        Case 200526
            Call Rotina_Erro(vbOKOnly, "ERRO_VOU_CC_SEM_TITULARCPF", gErr)

        Case 200527
            Call Rotina_Erro(vbOKOnly, "ERRO_VOU_CC_SEM_VALIDADECC", gErr)

        Case 200528
            Call Rotina_Erro(vbOKOnly, "ERRO_VOU_CC_SEM_NUMCC", gErr)

        Case 200529
            Call Rotina_Erro(vbOKOnly, "ERRO_VOU_CC_SEM_BANDEIRA", gErr)

        Case 200530
            Call Rotina_Erro(vbOKOnly, "ERRO_VOU_CC_SEM_PARC", gErr)
            
        Case 200625
            Call Rotina_Erro(vbOKOnly, "ERRO_VOU_CC_SEM_DATAAUTO", gErr)
            
        Case 200626
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR_DATA_FINAL", gErr)
            
        Case 200645
            Call Rotina_Erro(vbOKOnly, "ERRO_COMISSAO_COR_SEM_COR", gErr)

        Case 200645
            Call Rotina_Erro(vbOKOnly, "ERRO_COMISSAO_REP_SEM_REP", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200531)

    End Select

    Exit Function
    
End Function

Private Sub ValidadeCC_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ValidadeCC_GotFocus()
    Call MaskEdBox_TrataGotFocus(ValidadeCC, iAlterado)
End Sub

Private Sub VigenciaDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub VigenciaAte_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub VigenciaDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(VigenciaDe, iAlterado)
    
End Sub

Private Sub VigenciaDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VigenciaDe_Validate

    If Len(Trim(VigenciaDe.ClipText)) <> 0 Then

        lErro = Data_Critica(VigenciaDe.Text)
        If lErro <> SUCESSO Then gError 200532

    End If
    
    Call Calcula_Vigencia

    Exit Sub

Erro_VigenciaDe_Validate:

    Cancel = True

    Select Case gErr

        Case 200532

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200533)

    End Select

    Exit Sub

End Sub

Private Sub VigenciaAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(VigenciaAte, iAlterado)
End Sub

Private Sub VigenciaAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VigenciaAte_Validate

    If Len(Trim(VigenciaAte.ClipText)) <> 0 Then

        lErro = Data_Critica(VigenciaAte.Text)
        If lErro <> SUCESSO Then gError 200534

    End If
    
    Call Calcula_Vigencia

    Exit Sub

Erro_VigenciaAte_Validate:

    Cancel = True

    Select Case gErr

        Case 200534

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200535)

    End Select

    Exit Sub

End Sub

Private Sub UpDownVigenciaDe_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownVigenciaDe_DownClick

    VigenciaDe.SetFocus

    If Len(VigenciaDe.ClipText) > 0 Then

        sData = VigenciaDe.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 200536

        VigenciaDe.Text = sData

    End If

    Call Calcula_Vigencia

    Exit Sub

Erro_UpDownVigenciaDe_DownClick:

    Select Case gErr

        Case 200536

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200537)

    End Select

    Exit Sub

End Sub

Private Sub UpDownVigenciaDe_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownVigenciaDe_UpClick

    VigenciaDe.SetFocus

    If Len(Trim(VigenciaDe.ClipText)) > 0 Then

        sData = VigenciaDe.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 200538

        VigenciaDe.Text = sData

    End If

    Call Calcula_Vigencia

    Exit Sub

Erro_UpDownVigenciaDe_UpClick:

    Select Case gErr

        Case 200538

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200539)

    End Select

    Exit Sub

End Sub

Private Sub UpDownVigenciaAte_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownVigenciaAte_DownClick

    VigenciaAte.SetFocus

    If Len(VigenciaAte.ClipText) > 0 Then

        sData = VigenciaAte.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 200540

        VigenciaAte.Text = sData

    End If

    Call Calcula_Vigencia

    Exit Sub

Erro_UpDownVigenciaAte_DownClick:

    Select Case gErr

        Case 200540

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200541)

    End Select

    Exit Sub

End Sub

Private Sub UpDownVigenciaAte_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownVigenciaAte_UpClick

    VigenciaAte.SetFocus

    If Len(Trim(VigenciaAte.ClipText)) > 0 Then

        sData = VigenciaAte.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 200542

        VigenciaAte.Text = sData

    End If

    Call Calcula_Vigencia

    Exit Sub

Erro_UpDownVigenciaAte_UpClick:

    Select Case gErr

        Case 200542

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200543)

    End Select

    Exit Sub

End Sub

Public Sub PaxNome_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PaxNome_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPax)
End Sub

Public Sub PaxNome_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPax)
End Sub

Public Sub PaxNome_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridPax.objControle = PaxNome
    lErro = Grid_Campo_Libera_Foco(objGridPax)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub PaxDataNasc_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PaxDataNasc_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPax)
End Sub

Public Sub PaxDataNasc_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPax)
End Sub

Public Sub PaxDataNasc_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridPax.objControle = PaxDataNasc
    lErro = Grid_Campo_Libera_Foco(objGridPax)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub PaxTipoDoc_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PaxTipoDoc_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPax)
End Sub

Public Sub PaxTipoDoc_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPax)
End Sub

Public Sub PaxTipoDoc_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridPax.objControle = PaxTipoDoc
    lErro = Grid_Campo_Libera_Foco(objGridPax)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub PaxNumDoc_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PaxNumDoc_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPax)
End Sub

Public Sub PaxNumDoc_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPax)
End Sub

Public Sub PaxNumDoc_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridPax.objControle = PaxNumDoc
    lErro = Grid_Campo_Libera_Foco(objGridPax)
    If lErro <> SUCESSO Then Cancel = True
End Sub


Public Sub PaxSexo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PaxSexo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPax)
End Sub

Public Sub PaxSexo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPax)
End Sub

Public Sub PaxSexo_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridPax.objControle = PaxSexo
    lErro = Grid_Campo_Libera_Foco(objGridPax)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub PaxValorEmi_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PaxValorEmi_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPax)
End Sub

Public Sub PaxValorEmi_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPax)
End Sub

Public Sub PaxValorEmi_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridPax.objControle = PaxValorEmi
    lErro = Grid_Campo_Libera_Foco(objGridPax)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub PaxValor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PaxValor_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPax)
End Sub

Public Sub PaxValor_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPax)
End Sub

Public Sub PaxValor_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridPax.objControle = PaxValor
    lErro = Grid_Campo_Libera_Foco(objGridPax)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub PaxCancelado_Click()
    iAlterado = REGISTRO_ALTERADO
    Call Calcula_SomaPax
End Sub

Public Sub PaxCancelado_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPax)
End Sub

Public Sub PaxCancelado_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPax)
End Sub

Public Sub PaxCancelado_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridPax.objControle = PaxCancelado
    lErro = Grid_Campo_Libera_Foco(objGridPax)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Function Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    Select Case objControl.Name

        Case PaxTitular.Name, PaxValorEmi.Name
            objControl.Enabled = False
            
        Case PaxDataNasc.Name
            If gbVoucherBloqueado Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case PaxValor.Name, PaxCancelado.Name
            If gbVoucherBloqueado Or gbVoucherValorBloqueado Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
 
        Case Else
            objControl.Enabled = True
 
    End Select
    
    Rotina_Grid_Enable = SUCESSO

    Exit Function

Erro_Rotina_Grid_Enable:

    Rotina_Grid_Enable = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200544)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Data(objGridInt As AdmGrid, ByVal objControle As Object, Optional ByVal bBloqueiaInsercao As Boolean = False) As Long
'Faz a crítica da célula Data que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Data

    Set objGridInt.objControle = objControle

    If Len(Trim(objControle.ClipText)) > 0 Then
    
        'Critica a Data informada
        lErro = Data_Critica(objControle.Text)
        If lErro <> SUCESSO Then gError 200545
               
        If Not bBloqueiaInsercao Then
            'verifica se precisa preencher o grid com uma nova linha
            If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
                objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            End If
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 200546

    If Len(Trim(objControle.ClipText)) > 0 Then
        If objControle.Name = PaxDataNasc.Name Then
            Call Calcula_Dados_Pax(objGridInt.objGrid.Row)
        End If
    End If
        
    Saida_Celula_Data = SUCESSO

    Exit Function

Erro_Saida_Celula_Data:

    Saida_Celula_Data = gErr

    Select Case gErr

        Case 200545 To 200546
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200547)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Padrao(objGridInt As AdmGrid, ByVal objControle As Object, Optional ByVal bBloqueiaInsercao As Boolean = False) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Padrao

    Set objGridInt.objControle = objControle
    
    If Len(Trim(objControle.Text)) > 0 Then
    
        If Not bBloqueiaInsercao Then
            'verifica se precisa preencher o grid com uma nova linha
            If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
                objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            End If
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 200548

    Saida_Celula_Padrao = SUCESSO

    Exit Function

Erro_Saida_Celula_Padrao:

    Saida_Celula_Padrao = gErr

    Select Case gErr

        Case 200548
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200549)

    End Select

    Exit Function

End Function

Function Saida_Celula_Valor(objGridInt As AdmGrid, ByVal objControle As Object, Optional ByVal bBloqueiaInsercao As Boolean = False) As Long
'Faz a crítica da célula que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridInt.objControle = objControle
    
    If Len(Trim(objControle.Text)) > 0 Then
    
        'Critica o valor informado
        lErro = Valor_NaoNegativo_Critica(objControle.Text)
        If lErro <> SUCESSO Then gError 200550

        objControle.Text = Format(objControle.Text, "STANDARD")
               
        If Not bBloqueiaInsercao Then
            'verifica se precisa preencher o grid com uma nova linha
            If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
                objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            End If
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 200551
        
    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = gErr

    Select Case gErr

        Case 200550, 200551
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200552)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

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

        'GridPAX
        If objGridInt.objGrid.Name = GridPax.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col

                Case iGrid_PaxValor_Col
                
                    lErro = Saida_Celula_Valor(objGridInt, Me.ActiveControl, gbVoucherBloqueado)
                    If lErro <> SUCESSO Then gError 200553
                    
                    Call Calcula_SomaPax
                
                Case iGrid_PaxNome_Col, iGrid_PaxSobreNome_Col, iGrid_PaxTipoDoc_Col, iGrid_PaxNumDoc_Col, iGrid_PaxSexo_Col, iGrid_PaxCancelado_Col

                    lErro = Saida_Celula_Padrao(objGridInt, Me.ActiveControl, gbVoucherBloqueado)
                    If lErro <> SUCESSO Then gError 200554
                
                Case iGrid_PaxDataNasc_Col

                    lErro = Saida_Celula_Data(objGridInt, Me.ActiveControl, gbVoucherBloqueado)
                    If lErro <> SUCESSO Then gError 200555

            End Select
                    
        End If

        'GridContatos
        If objGridInt.objGrid.Name = GridContatos.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col

                Case iGrid_ContatoNome_Col, iGrid_ContatoTelefone_Col

                    lErro = Saida_Celula_Padrao(objGridInt, Me.ActiveControl)
                    If lErro <> SUCESSO Then gError 200556
                 
            End Select
                    
        End If

        'GridVendedores
        If objGridInt.objGrid.Name = GridVendedores.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col

                 Case iGrid_Vendedor_Col

                    lErro = Saida_Celula_Vendedor(objGridInt)
                    If lErro <> SUCESSO Then gError 200557
                
                 Case iGrid_VendedorPerc_Col

                    lErro = Saida_Celula_Percentual(objGridInt, Me.ActiveControl)
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

Private Sub Contato_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Email_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Endereco_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Estado_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Fax_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Telefone1_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Telefone2_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PaisLabel_Click()

Dim objPais As New ClassPais
Dim colSelecao As Collection

    'Chama a Tela de PaisesLista
    Call Chama_Tela("PaisesLista", colSelecao, objPais, objEventoPais)

End Sub

Public Sub LabelCidade_Click()

Dim objCidade As New ClassCidades
Dim colSelecao As Collection

    objCidade.sDescricao = Cidade.Text
    
    'Chama a Tela de browse
    Call Chama_Tela("CidadeLista", colSelecao, objCidade, objEventoCidade)

End Sub

Private Sub objEventoCidade_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCidade As ClassCidades

On Error GoTo Erro_objEventoCidade_evSelecao

    Set objCidade = obj1

    If objCidade Is Nothing Then
        Cidade.Text = ""
    Else
        Cidade.Text = CStr(objCidade.sDescricao)
    End If

    Me.Show

    Exit Sub

Erro_objEventoCidade_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200560)

    End Select

    Exit Sub

End Sub

Private Sub objEventoPais_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPais As New ClassPais

On Error GoTo Erro_objEventoPais_evSelecao

    Set objPais = obj1

    If objPais Is Nothing Then
        Pais.Text = ""
    Else
        Pais.Text = CStr(objPais.iCodigo)
        lErro = Combo_Item_Seleciona(Pais)
        If lErro <> SUCESSO Then gError 200561
    End If

    Me.Show

    Exit Sub

Erro_objEventoPais_evSelecao:

    Select Case gErr

        Case 200561

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200562)

    End Select

    Exit Sub

End Sub

Public Sub Cidade_Validate(Cancel As Boolean)

Dim lErro As Long, objCidade As New ClassCidades
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Cidade_Validate

    If Len(Trim(Cidade.Text)) = 0 Then Exit Sub
    
    objCidade.sDescricao = Cidade.Text
    lErro = CF("Cidade_Le_Nome", objCidade)
    If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 200563
    
    If lErro <> SUCESSO Then gError 200564
    
    Exit Sub
     
Erro_Cidade_Validate:

    Cancel = True
    
    Select Case gErr
          
        Case 200563
        
        Case 200564
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CIDADE")

            If vbMsgRes = vbYes Then
    
                 Call Chama_Tela("CidadeCadastro", objCidade)
    
            End If
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200565)
     
    End Select
     
    Exit Sub

End Sub

Private Sub CEP_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CEP_GotFocus()
    Call MaskEdBox_TrataGotFocus(CEP, iAlterado)
End Sub

Private Sub Cidade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CPFTitularCC_GotFocus()
    Call MaskEdBox_TrataGotFocus(CPFTitularCC, iAlterado)
End Sub

Private Sub CPFTitularCC_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CPFTitularCC_Validate

    If Len(Trim(CPFTitularCC.Text)) = 0 Then Exit Sub

    Select Case Len(Trim(CPFTitularCC.Text))

        Case STRING_CPF 'CPF
    
            lErro = Cpf_Critica(CPFTitularCC.Text)
            If lErro <> SUCESSO Then gError 200566
    
            CPFTitularCC.Format = "000\.000\.000-00; ; ; "
            CPFTitularCC.Text = CPFTitularCC.Text
    
    Case Else

        gError 200567

    End Select

    Exit Sub

Erro_CPFTitularCC_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 200566

        Case 200567
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC_CPF", Err)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200568)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAutoCC_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataAutoCC_DownClick

    DataAutoCC.SetFocus

    If Len(DataAutoCC.ClipText) > 0 Then

        sData = DataAutoCC.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 200569

        DataAutoCC.Text = sData

    End If

    Exit Sub

Erro_UpDownDataAutoCC_DownClick:

    Select Case gErr

        Case 200569

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200570)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAutoCC_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataAutoCC_UpClick

    DataAutoCC.SetFocus

    If Len(Trim(DataAutoCC.ClipText)) > 0 Then

        sData = DataAutoCC.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 200571

        DataAutoCC.Text = sData

    End If

    Exit Sub

Erro_UpDownDataAutoCC_UpClick:

    Select Case gErr

        Case 200571

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200572)

    End Select

    Exit Sub

End Sub

Private Sub DataAutoCC_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataAutoCC, iAlterado)
    
End Sub

Private Sub DataAutoCC_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAutoCC_Validate

    If Len(Trim(DataAutoCC.ClipText)) <> 0 Then

        lErro = Data_Critica(DataAutoCC.Text)
        If lErro <> SUCESSO Then gError 200573

    End If

    Exit Sub

Erro_DataAutoCC_Validate:

    Cancel = True

    Select Case gErr

        Case 200573

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200574)

    End Select

    Exit Sub

End Sub

Private Sub DataAutoCC_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ContatoNome_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ContatoNome_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridContatos)
End Sub

Public Sub ContatoNome_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContatos)
End Sub

Public Sub ContatoNome_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridContatos.objControle = ContatoNome
    lErro = Grid_Campo_Libera_Foco(objGridContatos)
    If lErro <> SUCESSO Then Cancel = True
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

Public Sub VendedorPerc_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub VendedorPerc_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridVendedores)
End Sub

Public Sub VendedorPerc_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridVendedores)
End Sub

Public Sub VendedorPerc_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridVendedores.objControle = VendedorPerc
    lErro = Grid_Campo_Libera_Foco(objGridVendedores)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub ContatoTelefone_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ContatoTelefone_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridContatos)
End Sub

Public Sub ContatoTelefone_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContatos)
End Sub

Public Sub ContatoTelefone_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridContatos.objControle = ContatoTelefone
    lErro = Grid_Campo_Libera_Foco(objGridContatos)
    If lErro <> SUCESSO Then Cancel = True
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

Private Sub PercComiAgeNovo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercComiAgeNovo_Validate

    'Verifica se foi preenchido a Comissao de Venda
    If Len(Trim(PercComiAgeNovo.Text)) = 0 Then Exit Sub

    'Critica se é porcentagem
    lErro = Porcentagem_Critica(PercComiAgeNovo.Text)
    If lErro <> SUCESSO Then gError 200575

    'Formata
    PercComiAgeNovo.Text = Format(PercComiAgeNovo.Text, "Fixed")

    VlrComiAge.Caption = Format(StrParaDbl(BrutoRS.Caption) * StrParaDbl(PercComiAgeNovo.Text) / 100, "STANDARD")
    
    If Not OptCCSim.Value Then
        ValorFaturavelRS.Caption = Format(StrParaDbl(BrutoRS.Caption) - StrParaDbl(VlrComiAge.Caption), "STANDARD")
    End If

    Exit Sub

Erro_PercComiAgeNovo_Validate:

    Cancel = True

    Select Case gErr

        Case 200575
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200576)
    
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
    If lErro <> SUCESSO Then gError 200577

    'Formata
    PercComiCorNovo.Text = Format(PercComiCorNovo.Text, "Fixed")

    VlrComiCor.Caption = Format(StrParaDbl(BrutoRS.Caption) * StrParaDbl(PercComiCorNovo.Text) / 100, "STANDARD")

    Exit Sub

Erro_PercComiCorNovo_Validate:

    Cancel = True

    Select Case gErr

        Case 200577
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200578)
    
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
    If lErro <> SUCESSO Then gError 200579

    'Formata
    PercComiRepNovo.Text = Format(PercComiRepNovo.Text, "Fixed")

    VlrComiRep.Caption = Format(StrParaDbl(BrutoRS.Caption) * StrParaDbl(PercComiRepNovo.Text) / 100, "STANDARD")

    Exit Sub

Erro_PercComiRepNovo_Validate:

    Cancel = True

    Select Case gErr

        Case 200579
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200580)
    
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
    If lErro <> SUCESSO Then gError 200581

    'Formata
    PercComiEmiNovo.Text = Format(PercComiEmiNovo.Text, "Fixed")

    VlrComiEmi.Caption = Format(StrParaDbl(BrutoRS.Caption) * StrParaDbl(PercComiEmiNovo.Text) / 100, "STANDARD")

    Exit Sub

Erro_PercComiEmiNovo_Validate:

    Cancel = True

    Select Case gErr

        Case 200581
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200582)
    
    End Select

    Exit Sub

End Sub

Public Sub LabelRepresentante_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection
Dim sNomeBrowse As String

    'Preenche NomeReduzido com o cliente da tela
    If Len(Trim(RepresentanteComissao.Text)) > 0 Then
        objCliente.lCodigo = LCodigo_Extrai(RepresentanteComissao.Text)
        objCliente.sNomeReduzido = RepresentanteComissao.Text
    End If

    sNomeBrowse = "ClientesLista"

    Call CF("Cliente_Obtem_NomeBrowse", sNomeBrowse)
    
    'Chama Tela ClienteLista
    Call Chama_Tela(sNomeBrowse, colSelecao, objCliente, objEventoRepresentante)

End Sub

Public Sub objEventoRepresentante_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    RepresentanteComissao.Text = objCliente.lCodigo
    Call RepresentanteComissao_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Public Sub LabelCorrentista_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection
Dim sNomeBrowse As String

    'Preenche NomeReduzido com o cliente da tela
    If Len(Trim(CorrentistaComissao.Text)) > 0 Then
        objCliente.lCodigo = LCodigo_Extrai(CorrentistaComissao.Text)
        objCliente.sNomeReduzido = CorrentistaComissao.Text
    End If

    sNomeBrowse = "ClientesLista"

    Call CF("Cliente_Obtem_NomeBrowse", sNomeBrowse)

    'Chama Tela ClienteLista
    Call Chama_Tela(sNomeBrowse, colSelecao, objCliente, objEventoCorrentista)

End Sub

Public Sub objEventoCorrentista_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    CorrentistaComissao.Text = objCliente.lCodigo
    Call CorrentistaComissao_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub RepresentanteComissao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_RepresentanteComissao_Validate

    If Len(Trim(RepresentanteComissao.Text)) > 0 Then
    
        RepresentanteComissao.Text = LCodigo_Extrai(RepresentanteComissao.Text)

        'Tenta ler o Vendedor (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(RepresentanteComissao, objCliente)
        If lErro <> SUCESSO Then gError 200583
        
    End If
    
    Exit Sub

Erro_RepresentanteComissao_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 200583

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200584)
    
    End Select

End Sub

Private Sub CorrentistaComissao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_CorrentistaComissao_Validate

    If Len(Trim(CorrentistaComissao.Text)) > 0 Then
    
        CorrentistaComissao.Text = LCodigo_Extrai(CorrentistaComissao.Text)

        'Tenta ler o Vendedor (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(CorrentistaComissao, objCliente)
        If lErro <> SUCESSO Then gError 200585
        
    End If
    
    Exit Sub

Erro_CorrentistaComissao_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 200585

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200586)
    
    End Select

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
        lErro = TP_Vendedor_Le2(Vendedor, objVendedor)
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

        Case 200587, 25689, 200590
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 200588 'Não encontrou nome reduzido de vendedor no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_VENDEDOR")

            If vbMsgRes = vbYes Then

                'Preenche objVendedor com nome reduzido
                objVendedor.sNomeReduzido = Vendedor.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                'Chama a tela de Vendedores
                Call Chama_Tela("Vendedores", objVendedor)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 200589 'Não encontrou codigo do vendedor no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_VENDEDOR")

            If vbMsgRes = vbYes Then

                'Prenche objVendedor com codigo
                objVendedor.iCodigo = Codigo_Extrai(Vendedor.Text)

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                'Chama a tela de Vendedores
                Call Chama_Tela("Vendedores", objVendedor)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200591)

    End Select

    Exit Function

End Function

Private Sub PromotorComissao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_PromotorComissao_Validate

    If Len(Trim(PromotorComissao.Text)) > 0 Then
    
        PromotorComissao.Text = LCodigo_Extrai(PromotorComissao.Text)

        'Tenta ler o Vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le2(PromotorComissao, objVendedor)
        If lErro <> SUCESSO Then gError 200592
        
    End If
    
    Exit Sub

Erro_PromotorComissao_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 200592

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200593)
    
    End Select

End Sub

Public Sub LabelPromotor_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection

    'Preenche NomeReduzido com o cliente da tela
    If Len(Trim(PromotorComissao.Text)) > 0 Then
        objVendedor.iCodigo = Codigo_Extrai(PromotorComissao.Text)
        objVendedor.sNomeReduzido = PromotorComissao.Text
    End If

    'Chama Tela ClienteLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoPromotor)

End Sub

Public Sub objEventoPromotor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor
Dim bCancel As Boolean

    Set objVendedor = obj1

    PromotorComissao.Text = objVendedor.iCodigo
    Call PromotorComissao_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is RepresentanteComissao Then Call LabelRepresentante_Click
        If Me.ActiveControl Is CorrentistaComissao Then Call LabelCorrentista_Click
        If Me.ActiveControl Is PromotorComissao Then Call LabelPromotor_Click
        If Me.ActiveControl Is Cidade Then Call LabelCidade_Click
        If Me.ActiveControl Is Pais Then Call PaisLabel_Click
        If Me.ActiveControl Is Cliente Then Call LabelCliente_Click
        If Me.ActiveControl Is Produto Then Call LabelProduto_Click
        If Me.ActiveControl Is NumeroVouP Then Call BotaoVou_Click
    
    End If
    
End Sub

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 200594

    'Limpa Tela
    Call Limpa_Tela_TRPVouchers

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 200594

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200595)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objVou As New ClassTRPVouchers
Dim bSimulaImport As Boolean

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(NumeroVou.Caption)) = 0 Then gError 200596
    If Len(Trim(TipoVou.Caption)) = 0 Then gError 200597
    If Len(Trim(SerieVou.Caption)) = 0 Then gError 200598
    If Grupo.ListIndex = -1 Then gError 200599
    If Len(Trim(Cliente.Text)) = 0 Then gError 200600
    If Len(Trim(Emissor.Text)) = 0 Then gError 200601
    If Len(Trim(Produto.Text)) = 0 Then gError 200602
    If StrParaDate(VigenciaDe.Text) = DATA_NULA Then gError 200603
    If StrParaDate(VigenciaAte.Text) = DATA_NULA Then gError 200604
    If Moeda.ListIndex = -1 Then gError 200605
    If StrParaInt(QtdePax.Text) = 0 Then gError 200606
    If StrParaInt(Cambio.Text) = 0 Then gError 200607
'    If StrParaInt(QtdePax.Text) <> objGridPax.iLinhasExistentes Then gError 200608
    '#####################

    'Preenche o objTRPTiposOcorrencia
    lErro = Move_Tela_Memoria(objVou, True, True)
    If lErro <> SUCESSO Then gError 200609
    
    lErro = Valida_Dados(objVou)
    If lErro <> SUCESSO Then gError 200610

    lErro = Trata_Alteracao(objVou, objVou.sTipVou, objVou.sSerie, objVou.lNumVou)
    If lErro <> SUCESSO Then gError 200611

    'Grava o voucher no Banco de Dados
    lErro = CF("TRPVouManu_Grava", objVou)
    If lErro <> SUCESSO Then gError 200612
 
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 200596
            Call Rotina_Erro(vbOKOnly, "ERRO_TRP_NUMVOU_NAO_PREENCHIDO", gErr)

        Case 200597
            Call Rotina_Erro(vbOKOnly, "ERRO_TRP_TIPO_NAO_PREENCHIDO", gErr)

        Case 200598
            Call Rotina_Erro(vbOKOnly, "ERRO_TRP_SERIE_NAO_PREENCHIDO", gErr)

        Case 200599
            Call Rotina_Erro(vbOKOnly, "ERRO_TRP_PLANO_NAO_PREENCHIDO", gErr)

        Case 200600
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 200601
            Call Rotina_Erro(vbOKOnly, "ERRO_TRP_EMISSOR_NAO_PREENCHIDO", gErr)

        Case 200602
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 200603
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVIGENCIA_NAO_PREENCHIDA", gErr)

        Case 200604
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVIGENCIA_NAO_PREENCHIDA", gErr)

        Case 200605
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDA_NAO_PREENCHIDA", gErr)

        Case 200606
            Call Rotina_Erro(vbOKOnly, "ERRO_TRP_QTDPAX_NAO_PREENCHIDO", gErr)

        Case 200607
            Call Rotina_Erro(vbOKOnly, "ERRO_TRP_CAMBIO_NAO_PREENCHIDO", gErr)

        Case 200608
            Call Rotina_Erro(vbOKOnly, "ERRO_TRP_QTDPAX_DIF_NUMPAX", gErr)

        Case 200609 To 200612

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200613)

    End Select

    Exit Function

End Function

Private Sub LabelProduto_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection
Dim sFiltro As String

On Error GoTo Erro_LabelProduto_Click

    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 200614

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""
    
    objProduto.sCodigo = sProdutoFormatado
    
    sFiltro = "Codigo IN (SELECT Codigo FROM TRPProdutos WHERE ExibeWeb = 1)"
            
    'Lista de produtos
    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProduto, sFiltro)
    
    Exit Sub

Erro_LabelProduto_Click:

    Select Case gErr

        Case 200614

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200615)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    Call CF("SCombo_Seleciona2", Produto, Trim(objProduto.sCodigo))
    
    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200616)

    End Select

    Exit Sub

End Sub

Private Sub Pais_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objPais As New ClassPais

On Error GoTo Erro_Pais_Validate

    'Verifica se foi preenchida a Combo Pais
    If Len(Trim(Pais.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o item selecionado na ComboBox Pais
    If Pais.Text = Pais.List(Pais.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Pais, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 200617

    'Não existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objPais.iCodigo = iCodigo
        
        'Tenta ler Pais com esse codigo no BD
        lErro = CF("Paises_Le", objPais)
        If lErro <> SUCESSO And lErro <> 47876 Then gError 200618
        If lErro <> SUCESSO Then gError 200619

        Pais.Text = CStr(iCodigo) & SEPARADOR & objPais.sNome
        
    End If
    
    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 200620

    Exit Sub

Erro_Pais_Validate:

    Cancel = True

    Select Case gErr

    Case 200617, 200618
    
    Case 200619
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PAIS", objPais.iCodigo)

        If vbMsgRes = vbYes Then
            Call Chama_Tela("Paises", objPais)
        Else
            'Segura o foco
        End If

    Case 200620
        Call Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_CADASTRADO1", gErr, Trim(Pais.Text))
    
    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200621)

    End Select

    Exit Sub

End Sub

Private Sub Pais_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Estado_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Estado_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Estado_Validate

    'Verifica se foi preenchido o Estado
    If Len(Trim(Estado.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o item selecionado na ComboBox Estado
    If Estado.Text = Estado.List(Estado.ListIndex) Then Exit Sub

    'Verifica se existe o item no Estado, se existir seleciona o item
    lErro = Combo_Item_Igual_CI(Estado)
    If lErro <> SUCESSO And lErro <> 58583 Then gError 200622

    'Nao existe o item na ComboBox Estado
    If lErro <> SUCESSO Then gError 200623

    Exit Sub

Erro_Estado_Validate:

    Cancel = True

    Select Case gErr

    Case 200622

    Case 200623
        Call Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_CADASTRADO", gErr, Estado.Text)
        
    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200624)

    End Select

    Exit Sub

End Sub

Private Sub Bloqueia_Funcionalidades(ByVal bFlag As Boolean)
Dim iIndice As Integer
    If bFlag Then
        FrameDadosBasicos.Enabled = False
        If gobjVoucher.lNumFat <> 0 Then
            FramePagoCC.Enabled = False
            FrameCC.Enabled = False
        Else
            FramePagoCC.Enabled = True
            FrameCC.Enabled = True
        End If
        If Not OptCCSim.Value And Not (gobjVoucher Is Nothing) Then
            If gobjVoucher.lNumFat <> 0 Then
                PercComiAgeNovo.Enabled = False
            Else
                PercComiAgeNovo.Enabled = True
            End If
        Else
            PercComiAgeNovo.Enabled = True
        End If
        BotaoRecalcularPax.Enabled = False
    
        For iIndice = 0 To FrameDadosBasicosAux.UBound
            FrameDadosBasicosAux.Item(iIndice).Enabled = True
        Next
    Else
        If OptCCSim.Value Then
            FrameCC.Enabled = True
        Else
            FrameCC.Enabled = False
        End If
    
        FrameDadosBasicos.Enabled = True
        
        If gbVoucherValorBloqueado Then
        
            If Not OptCCSim.Value Then
                PercComiAgeNovo.Enabled = False
            Else
                PercComiAgeNovo.Enabled = True
            End If
            
            BotaoRecalcularPax.Enabled = False
            FramePagoCC.Enabled = False
            
            For iIndice = 0 To FrameDadosBasicosAux.UBound
                FrameDadosBasicosAux.Item(iIndice).Enabled = False
            Next
            FramePagoCC.Enabled = False
       
        Else
            PercComiAgeNovo.Enabled = True
            BotaoRecalcularPax.Enabled = True
            FramePagoCC.Enabled = True
        
            For iIndice = 0 To FrameDadosBasicosAux.UBound
                FrameDadosBasicosAux.Item(iIndice).Enabled = True
            Next
            FramePagoCC.Enabled = True
            
        End If
    
    End If

End Sub

Private Sub BotaoRecalcComis_Click()

Dim lErro As Long
Dim objVou As New ClassTRPVouchers

On Error GoTo Erro_BotaoRecalcComis_Click

    lErro = Move_Tela_Memoria(objVou, False, True)
    If lErro <> SUCESSO Then gError 200650
    
    lErro = CF("TRPVou_Le_Comissao", objVou)
    If lErro <> SUCESSO Then gError 200651
    
    If objVou.lCorrentista <> 0 Then
        CorrentistaComissao.Text = objVou.lCorrentista
        Call CorrentistaComissao_Validate(bSGECancelDummy)
    Else
        CorrentistaComissao.Text = ""
    End If

    If objVou.lRepresentante <> 0 Then
        RepresentanteComissao.Text = objVou.lRepresentante
        Call RepresentanteComissao_Validate(bSGECancelDummy)
    Else
        RepresentanteComissao.Text = ""
    End If

    PercComiRepNovo.Text = objVou.dComissaoRep * 100
    Call PercComiRepNovo_Validate(bSGECancelDummy)
        
    PercComiCorNovo.Text = objVou.dComissaoCorr * 100
    Call PercComiCorNovo_Validate(bSGECancelDummy)
        
    PercComiEmiNovo.Text = objVou.dComissaoEmissor * 100
    Call PercComiEmiNovo_Validate(bSGECancelDummy)
    
    If PercComiAgeNovo.Enabled Then
        PercComiAgeNovo.Text = objVou.dComissaoAg * 100
        Call PercComiAgeNovo_Validate(bSGECancelDummy)
    End If
    
    Exit Sub

Erro_BotaoRecalcComis_Click:

    Select Case gErr
    
        Case 200650, 200651

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200652)

    End Select

    Exit Sub


End Sub

Private Sub TarifaUnitaria_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TarifaUnitaria_Validate

    'Veifica se TarifaUnitaria está preenchida
    If Len(Trim(TarifaUnitaria.Text)) <> 0 Then

       'Critica a TarifaUnitaria
       lErro = Valor_NaoNegativo_Critica(TarifaUnitaria.Text)
       If lErro <> SUCESSO Then gError 190697
        
    End If
    
    Call BotaoRecalcularPax_Click
    
    Exit Sub

Erro_TarifaUnitaria_Validate:

    Cancel = True

    Select Case gErr

        Case 190697

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190698)

    End Select

    Exit Sub
    
End Sub

Private Sub Cambio_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Cambio_Validate

    'Veifica se Cambio está preenchida
    If Len(Trim(Cambio.Text)) <> 0 Then

       'Critica a Cambio
       lErro = Valor_Positivo_Critica(Cambio.Text)
       If lErro <> SUCESSO Then gError 190697
        
    End If
        
    Call Calcula_ValoresEmReal
        
    Exit Sub

Erro_Cambio_Validate:

    Cancel = True

    Select Case gErr

        Case 190697

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190698)

    End Select

    Exit Sub
    
End Sub

Function Saida_Celula_Percentual(objGridInt As AdmGrid, ByVal objControle As Object) As Long
'Faz a crítica da célula Percentual Desconto que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentDesc As Double

On Error GoTo Erro_Saida_Celula_Percentual

    Set objGridInt.objControle = objControle

    If Len(objControle) > 0 Then
        
        'Critica a porcentagem
        lErro = Porcentagem_Critica(objControle.Text)
        If lErro <> SUCESSO Then gError 195929

        dPercentDesc = CDbl(objControle.Text)
        
        objControle.Text = Format(dPercentDesc, "Fixed")
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195930

    Saida_Celula_Percentual = SUCESSO

    Exit Function

Erro_Saida_Celula_Percentual:

    Saida_Celula_Percentual = gErr

    Select Case gErr

        Case 195929, 195930
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195931)

    End Select

    Exit Function

End Function

Private Sub PercComiProNovo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercComiProNovo_Validate

    'Verifica se foi preenchido a Comissao de Venda
    If Len(Trim(PercComiProNovo.Text)) = 0 Then Exit Sub

    'Critica se é porcentagem
    lErro = Porcentagem_Critica(PercComiProNovo.Text)
    If lErro <> SUCESSO Then gError 200577

    'Formata
    PercComiProNovo.Text = Format(PercComiProNovo.Text, "Fixed")

    Exit Sub

Erro_PercComiProNovo_Validate:

    Cancel = True

    Select Case gErr

        Case 200577
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200578)
    
    End Select

    Exit Sub

End Sub
