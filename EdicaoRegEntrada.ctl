VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl EdicaoRegEntradaOcx 
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   ScaleHeight     =   5640
   ScaleWidth      =   9600
   Begin VB.Frame FrameAuxiliar 
      Caption         =   "Valores"
      Height          =   2940
      Left            =   315
      TabIndex        =   60
      Top             =   1695
      Visible         =   0   'False
      Width           =   9735
      Begin MSMask.MaskEdBox Total 
         Height          =   285
         Left            =   7470
         TabIndex        =   61
         Top             =   915
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ValorFrete 
         Height          =   285
         Left            =   1551
         TabIndex        =   62
         Top             =   915
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ValorDesconto 
         Height          =   285
         Left            =   165
         TabIndex        =   63
         Top             =   390
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ValorDespesas 
         Height          =   285
         Left            =   4335
         TabIndex        =   64
         Top             =   915
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ValorSeguro 
         Height          =   285
         Left            =   2937
         TabIndex        =   65
         Top             =   915
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox PercDescontoItens 
         Height          =   285
         Left            =   1410
         TabIndex        =   85
         ToolTipText     =   "Percentual de desconto dos itens"
         Top             =   2130
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ValorDescontoItens 
         Height          =   285
         Left            =   2790
         TabIndex        =   86
         ToolTipText     =   "Soma dos descontos dos itens"
         Top             =   2130
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label ValorProdutos2 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5535
         TabIndex        =   92
         Top             =   1620
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Base ISS"
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
         Index           =   15
         Left            =   2820
         TabIndex        =   91
         Top             =   1425
         Width           =   1065
      End
      Begin VB.Label ISSBase1 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2790
         TabIndex        =   90
         Top             =   1620
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
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
         Height          =   180
         Index           =   12
         Left            =   2820
         TabIndex        =   89
         Top             =   1920
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "% Desconto"
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
         Index           =   11
         Left            =   1470
         TabIndex        =   88
         Top             =   1920
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
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
         Height          =   180
         Index           =   7
         Left            =   5580
         TabIndex        =   87
         Top             =   1425
         Width           =   1065
      End
      Begin VB.Label ISSValor1 
         Caption         =   "ISSValor1"
         Height          =   240
         Left            =   120
         TabIndex        =   83
         Top             =   885
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label Label1 
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
         Index           =   13
         Left            =   3192
         TabIndex        =   82
         Top             =   720
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
         Left            =   300
         TabIndex        =   81
         Top             =   210
         Width           =   825
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   4
         Left            =   1888
         TabIndex        =   80
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label1 
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
         Index           =   32
         Left            =   4653
         TabIndex        =   79
         Top             =   720
         Width           =   840
      End
      Begin VB.Label LabelTotais 
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
         Left            =   7807
         TabIndex        =   78
         Top             =   705
         Width           =   450
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
         Left            =   6519
         TabIndex        =   77
         Top             =   720
         Width           =   255
      End
      Begin VB.Label SubTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   7470
         TabIndex        =   76
         Top             =   390
         Width           =   1125
      End
      Begin VB.Label Label18 
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
         Index           =   0
         Left            =   7650
         TabIndex        =   75
         Top             =   195
         Width           =   765
      End
      Begin VB.Label ICMSBase1 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1551
         TabIndex        =   74
         Top             =   390
         Width           =   1125
      End
      Begin VB.Label ICMSValor1 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2937
         TabIndex        =   73
         Top             =   390
         Width           =   1125
      End
      Begin VB.Label ICMSSubstBase1 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4323
         TabIndex        =   72
         Top             =   390
         Width           =   1500
      End
      Begin VB.Label ICMSSubstValor1 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6084
         TabIndex        =   71
         Top             =   390
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ICMS"
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
         Index           =   1
         Left            =   3267
         TabIndex        =   70
         Top             =   195
         Width           =   465
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Base ICMS"
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
         Left            =   1641
         TabIndex        =   69
         Top             =   195
         Width           =   945
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Base ICMS Subst"
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
         Left            =   4331
         TabIndex        =   68
         Top             =   195
         Width           =   1485
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "ICMS Subst"
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
         Left            =   6144
         TabIndex        =   67
         Top             =   195
         Width           =   1005
      End
      Begin VB.Label IPIValor1 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6084
         TabIndex        =   66
         Top             =   915
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4635
      Index           =   1
      Left            =   255
      TabIndex        =   35
      Top             =   750
      Width           =   9180
      Begin VB.Frame Frame2 
         Caption         =   "Identificação"
         Height          =   1485
         Index           =   0
         Left            =   180
         TabIndex        =   42
         Top             =   330
         Width           =   8895
         Begin VB.ComboBox TipoNFiscal 
            Height          =   315
            Left            =   1260
            TabIndex        =   1
            Top             =   315
            Width           =   3435
         End
         Begin VB.ComboBox Serie 
            Height          =   315
            Left            =   1260
            TabIndex        =   2
            Top             =   855
            Width           =   765
         End
         Begin MSMask.MaskEdBox NaturezaOp 
            Height          =   315
            Left            =   7530
            TabIndex        =   4
            Top             =   870
            Width           =   585
            _ExtentX        =   1032
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
         Begin MSMask.MaskEdBox NFiscal 
            Height          =   315
            Left            =   3945
            TabIndex        =   3
            Top             =   840
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin VB.Label Status 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7035
            TabIndex        =   48
            Top             =   330
            Width           =   1080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Status:"
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
            Index           =   5
            Left            =   6375
            TabIndex        =   47
            Top             =   390
            Width           =   615
         End
         Begin VB.Label LblNaturezaOp 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   5775
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   46
            Top             =   915
            Width           =   1725
         End
         Begin VB.Label TipoLabel 
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
            Left            =   750
            TabIndex        =   45
            Top             =   360
            Width           =   450
         End
         Begin VB.Label NFiscalLabel 
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
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   3150
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   44
            Top             =   900
            Width           =   720
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   690
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   43
            Top             =   900
            Width           =   510
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados do Cliente/Fornecedor"
         Height          =   810
         Index           =   2
         Left            =   180
         TabIndex        =   49
         Top             =   2145
         Width           =   8895
         Begin VB.ComboBox UFFornecedor 
            Height          =   315
            Left            =   7590
            TabIndex        =   8
            Top             =   360
            Width           =   675
         End
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   4935
            TabIndex        =   7
            Top             =   360
            Width           =   1635
         End
         Begin MSMask.MaskEdBox Cliente 
            Height          =   300
            Left            =   1605
            TabIndex        =   5
            Top             =   360
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Fornecedor 
            Height          =   300
            Left            =   1605
            TabIndex        =   6
            Top             =   360
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "UF:"
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
            Left            =   7230
            TabIndex        =   59
            Top             =   420
            Width           =   315
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
            Left            =   510
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   52
            Top             =   405
            Width           =   1005
         End
         Begin VB.Label LabelFilial 
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
            Left            =   4440
            TabIndex        =   51
            Top             =   435
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   510
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   50
            Top             =   405
            Width           =   1005
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datas"
         Height          =   735
         Index           =   4
         Left            =   150
         TabIndex        =   37
         Top             =   3270
         Width           =   8895
         Begin MSComCtl2.UpDown UpDownEmissao 
            Height          =   300
            Left            =   3075
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   285
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEmissao 
            Height          =   300
            Left            =   1995
            TabIndex        =   9
            Top             =   285
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownEntrada 
            Height          =   300
            Left            =   6840
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   300
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEntrada 
            Height          =   300
            Left            =   5775
            TabIndex        =   10
            Top             =   300
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label DataEntradaLabel 
            AutoSize        =   -1  'True
            Caption         =   "Entrada:"
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
            Left            =   4995
            TabIndex        =   41
            Top             =   360
            Width           =   735
         End
         Begin VB.Label DataEmissaoLabel 
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
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   1170
            TabIndex        =   40
            Top             =   315
            Width           =   735
         End
      End
      Begin VB.CheckBox CheckAlterada 
         Caption         =   "Alterada no módulo fiscal"
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
         Left            =   1290
         TabIndex        =   11
         Top             =   4230
         Value           =   1  'Checked
         Width           =   2490
      End
      Begin VB.CheckBox CheckIgnorar 
         Caption         =   "Ignorar esta nota fiscal nos livros fiscais"
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
         Left            =   4200
         TabIndex        =   12
         Top             =   4230
         Width           =   3810
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Tributacao"
      Height          =   4650
      Index           =   3
      Left            =   240
      TabIndex        =   36
      Top             =   750
      Visible         =   0   'False
      Width           =   9105
      Begin TelasFIS.TabTributacaoFat TabTrib 
         Height          =   4515
         Left            =   105
         TabIndex        =   84
         Top             =   150
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7964
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4695
      Index           =   2
      Left            =   240
      TabIndex        =   0
      Top             =   735
      Visible         =   0   'False
      Width           =   9045
      Begin VB.TextBox UFOrigem 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   930
         TabIndex        =   15
         Top             =   270
         Width           =   375
      End
      Begin VB.CommandButton BotaoPlanoConta 
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
         Height          =   420
         Left            =   7230
         TabIndex        =   30
         Top             =   4110
         Width           =   1665
      End
      Begin MSMask.MaskEdBox AliquotaIPILanc 
         Height          =   285
         Left            =   4770
         TabIndex        =   24
         Top             =   690
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "0%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox OutrasICMSLanc 
         Height          =   285
         Left            =   2010
         TabIndex        =   22
         Top             =   720
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox IsentaICMSLanc 
         Height          =   285
         Left            =   180
         TabIndex        =   21
         Top             =   690
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox BaseICMSLanc 
         Height          =   285
         Left            =   3600
         TabIndex        =   18
         Top             =   150
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox BaseIPILanc 
         Height          =   285
         Left            =   3420
         TabIndex        =   23
         Top             =   690
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.TextBox Observacoes 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3720
         MaxLength       =   255
         TabIndex        =   29
         Top             =   4410
         Width           =   3225
      End
      Begin MSMask.MaskEdBox ImpostoICMSLanc 
         Height          =   285
         Left            =   6480
         TabIndex        =   20
         Top             =   150
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox AliquotaICMSLanc 
         Height          =   285
         Left            =   5220
         TabIndex        =   19
         Top             =   150
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "0%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox FiscalLanc 
         Height          =   285
         Left            =   2790
         TabIndex        =   17
         Top             =   150
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ContabilLanc 
         Height          =   285
         Left            =   1950
         TabIndex        =   28
         Top             =   150
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ValorContabilLanc 
         Height          =   285
         Left            =   765
         TabIndex        =   16
         Top             =   150
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox IsentaIPILanc 
         Height          =   285
         Left            =   180
         TabIndex        =   26
         Top             =   4410
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox OutrasIPILanc 
         Height          =   285
         Left            =   2040
         TabIndex        =   27
         Top             =   4410
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridLancamentos 
         Height          =   3675
         Left            =   210
         TabIndex        =   14
         Top             =   330
         Width           =   8745
         _ExtentX        =   15425
         _ExtentY        =   6482
         _Version        =   393216
      End
      Begin MSMask.MaskEdBox ImpostoIPILanc 
         Height          =   285
         Left            =   6000
         TabIndex        =   25
         Top             =   690
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Aliquota IPI"
         Height          =   195
         Left            =   4950
         TabIndex        =   57
         Top             =   510
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "Outras ICMS"
         Height          =   195
         Left            =   2190
         TabIndex        =   56
         Top             =   510
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "Isenta ou Não Trib. ICMS"
         Height          =   195
         Left            =   180
         TabIndex        =   55
         Top             =   510
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Imposto  Creditado IPI"
         Height          =   195
         Left            =   6030
         TabIndex        =   54
         Top             =   510
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Base Cálculo IPI"
         Height          =   195
         Left            =   3510
         TabIndex        =   53
         Top             =   510
         Visible         =   0   'False
         Width           =   1170
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7740
      ScaleHeight     =   495
      ScaleWidth      =   1560
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   75
      Width           =   1620
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "EdicaoRegEntrada.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   585
         Picture         =   "EdicaoRegEntrada.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1095
         Picture         =   "EdicaoRegEntrada.ctx":068C
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoDocOriginal 
      Caption         =   "Documento Original"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5370
      TabIndex        =   31
      Top             =   120
      Width           =   2205
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5115
      Left            =   165
      TabIndex        =   13
      Top             =   375
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   9022
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Principais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Lançamentos no livro"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tributação"
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
Attribute VB_Name = "EdicaoRegEntradaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Sugestões --> Colocar um campo para informar a Descricao da Classificacao Contabil
'Pendencia --> Retestar a saida de Célula de Contabil

'Property Variables:
Dim m_Caption As String
Event Unload()

'Variaveis Globais
Public gobjNFiscal As ClassNFiscal
Public gobjLivRegES As ClassLivRegES
Public iAlterado As Integer
Private iFrameAtual As Integer
Private iClienteAlterado As Integer
Private iFornecedorAlterado As Integer
Private iSerieAlterada As Integer
Private iEmitenteNF As Integer

'---------------- Tributação ----------------
Public gobjTribTab As ClassTribTab

'Variaveis do Grid de Lancamentos
Public objGridLancamento As AdmGrid

Public iGrid_UF_Col As Integer
Public iGrid_ValorContabil_Col As Integer
Public iGrid_Contabil_Col As Integer
Public iGrid_Fiscal_Col As Integer
Public iGrid_BaseCalculoICMS_Col As Integer
Public iGrid_AliquotaICMS_Col As Integer
Public iGrid_ImpostoCreditadoICMS_Col As Integer
Public iGrid_IsentaICMS_Col As Integer
Public iGrid_OutrasICMS_Col As Integer
Public iGrid_BaseCalculoIPI_Col As Integer
Public iGrid_AliquotaIPI_Col As Integer
Public iGrid_ImpostoCreditadoIPI_Col As Integer
Public iGrid_IsentaIPI_Col As Integer
Public iGrid_OutrasIPI_Col As Integer
Public iGrid_Observacao_Col As Integer

'Eventos
Private WithEvents objEventoSerie As AdmEvento
Attribute objEventoSerie.VB_VarHelpID = -1
Private WithEvents objEventoNFiscal As AdmEvento
Attribute objEventoNFiscal.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoNatureza As AdmEvento
Attribute objEventoNatureza.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoNaturezaInterna As AdmEvento
Attribute objEventoNaturezaInterna.VB_VarHelpID = -1
Private WithEvents objEventoTiposDeTributacao As AdmEvento
Attribute objEventoTiposDeTributacao.VB_VarHelpID = -1
Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1

'Constantes da Tela
Const TAB_TRIBUTACAO = 3

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colEstado As New Collection
Dim objEstado As New ClassEstado

On Error GoTo Erro_Form_Load
        
    ''*** incluidos p/tratamento de tributacao *******************************
    Set gobjTribTab = New ClassTribTab
    lErro = gobjTribTab.Ativar(Me, NF_ENTRADA, 1)
    If lErro <> SUCESSO Then gError 69340
    '*** fim tributacao
    
    'Inicializa as variáveis relacionadas c\ as telas de Browse
    Set objEventoSerie = New AdmEvento
    Set objEventoNFiscal = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoNatureza = New AdmEvento
    Set objEventoFornecedor = New AdmEvento
    Set objGridLancamento = New AdmGrid
    Set objEventoTiposDeTributacao = New AdmEvento
    Set objEventoNaturezaInterna = New AdmEvento
    Set objEventoContaContabil = New AdmEvento
        
    iFrameAtual = 1
    CheckAlterada.Value = vbUnchecked
    CheckIgnorar.Value = vbUnchecked
    
    'Carrega as combos com os tipos de Notas Fiscais
    lErro = Carrega_TiposDocInfo()
    If lErro <> SUCESSO Then gError 69341
        
    'Carrega as combos de Séries
    lErro = Carrega_Serie()
    If lErro <> SUCESSO Then gError 69342
        
    'Preenche a ComboBox com as siglas dos Estados existentes no BD
    lErro = CF("Estados_Le_Todos", colEstado)
    If lErro <> SUCESSO Then gError 78019

    For Each objEstado In colEstado
       'Insere na ComboBox a sigla do Estado
       UFFornecedor.AddItem objEstado.sSigla
    Next
        
    'Inicializa mascara de ContabilLanc
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContabilLanc)
    If lErro <> SUCESSO Then gError 69343
        
    'Faz a Inicialização do Grid de Lancamentos
    lErro = Inicializa_Grid_Lancamentos(objGridLancamento)
    If lErro <> SUCESSO Then gError 69344
        
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case 69340, 69341, 69342, 69343, 69344, 78019
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159206)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoSerie = Nothing
    Set objEventoNFiscal = Nothing
    Set objEventoCliente = Nothing
    Set objEventoNatureza = Nothing
    Set objEventoFornecedor = Nothing
    Set objEventoTiposDeTributacao = Nothing
    Set objEventoNaturezaInterna = Nothing
    Set objEventoContaContabil = Nothing
    
    Set objGridLancamento = Nothing
    'Set gobjTribTab = Nothing

    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
   
    ''*** incluidos p/tratamento de tributacao *******************************
    If Not (gobjTribTab Is Nothing) Then Call gobjTribTab.Desativar
    Set gobjTribTab = Nothing
    '*** fim tributacao

    Set gobjLivRegES = Nothing
    Set gobjNFiscal = Nothing
    
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoPlanoConta_Click()

Dim lErro As Long
Dim iContaPreenchida As Integer
Dim sConta As String
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoPlanoConta_Click

    If GridLancamentos.Row = 0 Then gError 43770
    
    sConta = String(STRING_CONTA, 0)
    
    lErro = CF("Conta_Formata", GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Contabil_Col), sConta, iContaPreenchida)
    If lErro <> SUCESSO Then gError 43772
    
    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    'Chama PlanoContaFISLista
    Call Chama_Tela("PlanoContaFISLista", colSelecao, objPlanoConta, objEventoContaContabil)
    
    Exit Sub

Erro_BotaoPlanoConta_Click:

    Select Case gErr

        Case 43770
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 43772, 43773

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159207)

    End Select

    Exit Sub

End Sub

Private Sub GridLancamentos_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer
Dim iItemAtual As Integer, lErro As Long

On Error GoTo Erro_GridItens_KeyDown

    iLinhasExistentesAnterior = objGridLancamento.iLinhasExistentes
    iItemAtual = GridLancamentos.Row

    Call Grid_Trata_Tecla1(KeyCode, objGridLancamento)

    If objGridLancamento.iLinhasExistentes < iLinhasExistentesAnterior Then

'        lErro = gobjTribTab.Exclusao_Item_Grid(iItemAtual)
'        If lErro <> SUCESSO Then gError 22884
       
    End If
    
    Exit Sub
    
Erro_GridItens_KeyDown:

    Select Case gErr

        Case 22884, 55567

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159208)

    End Select

    Exit Sub

End Sub

Private Sub GridLancamentos_Scroll()

    Call Grid_Scroll(objGridLancamento)

End Sub

Private Sub objEventoContaContabil_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaContabil_evSelecao

    Set objPlanoConta = obj1

    If objPlanoConta.sConta <> "" Then
   
        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then gError 43774
        
        ContabilLanc.PromptInclude = False
        ContabilLanc.Text = sContaEnxuta
        ContabilLanc.PromptInclude = True
        
        GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Contabil_Col) = ContabilLanc.Text
    
    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case gErr

        Case 43774
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159209)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'se abriu o tab de tributacao
        If TabStrip1.SelectedItem.Index = TAB_TRIBUTACAO Then
            lErro = gobjTribTab.TabClick()
            If lErro <> SUCESSO Then gError 69345
            iAlterado = 0
        End If
        
        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual invisivel
        Frame1(iFrameAtual).Visible = False
        
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index
        
    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case 69345

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159210)

    End Select

    Exit Sub

End Sub

Private Function Carrega_TiposDocInfo() As Long
'Carrega na os Tipo de Documentos das Notas de Entrada

Dim lErro As Long
Dim colTipoDocInfo As New colTipoDocInfo
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim iIndice As Integer

On Error GoTo Erro_Carrega_TiposDocInfo

    Set colTipoDocInfo = gobjCRFAT.colTiposDocInfo

    'Carrega na combo só os Tipos ligados a Notas de Entrada
    For Each objTipoDocInfo In colTipoDocInfo
    
        If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFEXT Then
            TipoNFiscal.AddItem CStr(objTipoDocInfo.iCodigo) & SEPARADOR & objTipoDocInfo.sNomeReduzido
            TipoNFiscal.ItemData(TipoNFiscal.NewIndex) = objTipoDocInfo.iCodigo
            'seleciona DOCINFO_NFEEFN como default, como é feito na Versão light
            If objTipoDocInfo.iCodigo = DOCINFO_NFEEFN Then TipoNFiscal.ListIndex = TipoNFiscal.NewIndex
        End If
    Next

    Carrega_TiposDocInfo = SUCESSO

    Exit Function

Erro_Carrega_TiposDocInfo:

    Carrega_TiposDocInfo = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159211)

    End Select

    Exit Function

End Function

Private Function Carrega_Serie() As Long
'Carrega a combo de Séries com as séries lidas do BD

Dim lErro As Long
Dim colSerie As New colSerie
Dim objSerie As New ClassSerie

On Error GoTo Erro_Carrega_Serie

    'Lê as séries
    lErro = CF("Series_Le", colSerie)
    If lErro <> SUCESSO Then gError 69346

    'Carrega nas combos
    For Each objSerie In colSerie
        Serie.AddItem objSerie.sSerie
    Next

    Carrega_Serie = SUCESSO

    Exit Function

Erro_Carrega_Serie:

    Carrega_Serie = gErr

    Select Case gErr

        Case 69346

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159212)

    End Select

    Exit Function

End Function

Function Inicializa_Grid_Lancamentos(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Lancamentos

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("UF")
    objGridInt.colColuna.Add ("Valor Contabil")
    objGridInt.colColuna.Add ("Contabil")
    objGridInt.colColuna.Add ("Fiscal")
    objGridInt.colColuna.Add ("Base de Cálculo ICMS")
    objGridInt.colColuna.Add ("Alíquota ICMS")
    objGridInt.colColuna.Add ("Imposto Creditado ICMS")
    objGridInt.colColuna.Add ("Isenta ou não Trib. ICMS")
    objGridInt.colColuna.Add ("Outras ICMS")
    objGridInt.colColuna.Add ("Base de Calculo IPI")
    objGridInt.colColuna.Add ("Alíquota IPI")
    objGridInt.colColuna.Add ("Imposto Creditado IPI")
    objGridInt.colColuna.Add ("Isenta ou não Trib. IPI")
    objGridInt.colColuna.Add ("Outras IPI")
    objGridInt.colColuna.Add ("Observação")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (UFOrigem.Name)
    objGridInt.colCampo.Add (ValorContabilLanc.Name)
    objGridInt.colCampo.Add (ContabilLanc.Name)
    objGridInt.colCampo.Add (FiscalLanc.Name)
    objGridInt.colCampo.Add (BaseICMSLanc.Name)
    objGridInt.colCampo.Add (AliquotaICMSLanc.Name)
    objGridInt.colCampo.Add (ImpostoICMSLanc.Name)
    objGridInt.colCampo.Add (IsentaICMSLanc.Name)
    objGridInt.colCampo.Add (OutrasICMSLanc.Name)
    objGridInt.colCampo.Add (BaseIPILanc.Name)
    objGridInt.colCampo.Add (AliquotaIPILanc.Name)
    objGridInt.colCampo.Add (ImpostoIPILanc.Name)
    objGridInt.colCampo.Add (IsentaIPILanc.Name)
    objGridInt.colCampo.Add (OutrasIPILanc.Name)
    objGridInt.colCampo.Add (Observacoes.Name)

    
    iGrid_UF_Col = 1
    iGrid_ValorContabil_Col = 2
    iGrid_Contabil_Col = 3
    iGrid_Fiscal_Col = 4
    iGrid_BaseCalculoICMS_Col = 5
    iGrid_AliquotaICMS_Col = 6
    iGrid_ImpostoCreditadoICMS_Col = 7
    iGrid_IsentaICMS_Col = 8
    iGrid_OutrasICMS_Col = 9
    iGrid_BaseCalculoIPI_Col = 10
    iGrid_AliquotaIPI_Col = 11
    iGrid_ImpostoCreditadoIPI_Col = 12
    iGrid_IsentaIPI_Col = 13
    iGrid_OutrasIPI_Col = 14
    iGrid_Observacao_Col = 15
    
    'Grid do GridInterno
    objGridInt.objGrid = GridLancamentos

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_LANCAMENTOS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 10

    'Largura da primeira coluna
    GridLancamentos.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Não pode incluir linhas
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Lancamentos = SUCESSO

    Exit Function

End Function

Private Sub BotaoDocOriginal_Click()

Dim objTipoDocInfo As New ClassTipoDocInfo
Dim lErro As Long
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_BotaoDocOriginal_Click
    
    'Se o Tipo de Nota não está preenchido --> Erro
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 69347
    
    'Se não foi selecionado a Nota
    If (gobjNFiscal Is Nothing) Then gError 69506
    
    objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)

    'Lê o Tipo de Documento
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 69348

    'Se não achou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 69349
    
    objNFiscal.lNumIntDoc = gobjNFiscal.lNumIntDoc
    
    'Chama a Tela de Acordo com o Tipo selecionado
    Call Chama_Tela(objTipoDocInfo.sNomeTelaNFiscal, objNFiscal)
    
    Exit Sub
    
Erro_BotaoDocOriginal_Click:

    Select Case gErr
        
        Case 69347
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_PREENCHIDO", gErr)
        
        Case 69348
        
        Case 69349
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_CADASTRADO1", gErr, TipoNFiscal.Text)
        
        Case 69506
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_NAO_SELECIONADA", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159213)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_EdicaoRegEntrada()
'Limpa a Tela NFiscalEntrada, mas mantém a natureza e o tipo
    
Dim sNatureza As String
Dim sTipoNFiscal As String
Dim lErro As Long

    sNatureza = NaturezaOp.Text
    sTipoNFiscal = TipoNFiscal.Text
    
    'Chama o Limpa tela
    Call Limpa_Tela(Me)

    'Limpa os campos que não são limpos pelo Limpa_Tela
    Filial.Clear
    Status.Caption = ""
    Serie.Text = ""
    Serie.ListIndex = -1
    CheckAlterada.Value = vbUnchecked
    CheckIgnorar.Value = vbUnchecked
    UFFornecedor.Text = ""
    
    Call Grid_Limpa(objGridLancamento)
    
    Call gobjTribTab.Limpa_Tela
    
    NaturezaOp.Text = sNatureza
    TipoNFiscal.Text = sTipoNFiscal
        
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0
    
End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 69350
    
    'Habilita
    Call Habilita_DadosPrincipais
    
    'Limpa a Tela
    Call Limpa_Tela_EdicaoRegEntrada
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 69350

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159214)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 69351
    
    'Habilita os campos
    Call Habilita_DadosPrincipais
    
    'Limpa a Tela
    Call Limpa_Tela_EdicaoRegEntrada
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 69351

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 159215)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long
'Dispara o processo de Gravação dos Livros Fiscais do Registro de entrada

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim iIndice As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se foi selecionado uma Nota Fiscal
    If gobjNFiscal Is Nothing Then gError 69352
    
    'Verifica se o Tipo está Preenchido
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 69353
    
    'Verifica se a Série está prenchida
    If Len(Trim(Serie.Text)) = 0 Then gError 69354
    
    'Verifica se o Número da Nota está preenchida
    If Len(Trim(NFiscal.ClipText)) = 0 Then gError 69355
    
    'Verific a se a Natureza de Operacao está Preenchida
    If Len(Trim(NaturezaOp.ClipText)) = 0 Then gError 69356

    objTipoDocInfo.iCodigo = TipoNFiscal.ItemData(TipoNFiscal.ListIndex)
    
    'Lê o Tipo da NF
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 69357

    'Se não achou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 69358

    'De acordo com a sua Origem verifica se o Cliente ou Fornecedor estão preenchidos
    If objTipoDocInfo.iOrigem = DOCINFO_CLIENTE Then
        If Len(Trim(Cliente.ClipText)) = 0 Then gError 69359
    ElseIf objTipoDocInfo.iOrigem = DOCINFO_FORNECEDOR Then
        If Len(Trim(Fornecedor.ClipText)) = 0 Then gError 69360
    End If
    
    'Verifica se a Filial está preenchida
    If Len(Trim(Filial.Text)) = 0 Then gError 69361
        
    'Verifica se a Data de Entrada está preenchida
    If Len(Trim(DataEntrada.ClipText)) = 0 Then gError 69362
    
    'Verifica se a Date de Emissão está Preenchida
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 69363

    'Verifica se a Data de Entrada é menor que a Data de Emissao
    If CDate(DataEntrada.Text) < CDate(DataEmissao.Text) Then gError 69364

    'Verifica se tem pelo menos uma Linha no Grid
    If objGridLancamento.iLinhasExistentes = 0 Then gError 69365
    
    'Verifica os campos Obrigatórios no Grid
    For iIndice = 1 To objGridLancamento.iLinhasExistentes

        'Verifica se os campos do Grid estão todos preenchidos
        If objTipoDocInfo.iModeloArqICMS < 9000 And Len(Trim(GridLancamentos.TextMatrix(iIndice, iGrid_Fiscal_Col))) = 0 And gobjFIS.iRegravaNF = 0 Then gError 69366
        If Len(Trim(GridLancamentos.TextMatrix(iIndice, iGrid_UF_Col))) = 0 Then gError 69367
        'If Len(Trim(GridLancamentos.TextMatrix(iIndice, iGrid_ValorContabil_Col))) = 0 Then gError 69368

    Next
    
    'e for para Ignorar a Nota o Status passa a ser
    If CheckIgnorar.Value = vbChecked Then
        gobjLivRegES.iStatusLivro = STATUS_FIS_ALTERADA_EXCLUIDA
    Else
        gobjLivRegES.iStatusLivro = STATUS_FIS_ALTERADA
    End If
    
    gobjLivRegES.dPISValor = gobjTribTab.dPISValor
    gobjLivRegES.dCOFINSValor = gobjTribTab.dCOFINSValor
    
    gobjLivRegES.dtData = StrParaDate(DataEntrada.Text)
    
    'Recolhe os Ítens do Grid de Lancamentos
    lErro = Move_GridLancamento_Memoria(gobjLivRegES)
    If lErro <> SUCESSO Then gError 69369
    
    'Chama a Rotina de Gravação das Linhas do Livro e Os itens da Nota
    lErro = CF("EdicaoRegES_Grava", gobjLivRegES)
    If lErro <> SUCESSO Then gError 69370
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 69352
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_NAO_SELECIONADA", gErr)
        
        Case 69353
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)

        Case 69354
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)
        
        Case 69355
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", gErr)
        
        Case 69356
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_NAO_PREENCHIDA", gErr)
        
        Case 69357
        
        Case 69358
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)

        Case 69359
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 69360
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 69361
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 69362
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAENTRADA_NAO_PREENCHIDA", gErr)

        Case 69363
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
        
        Case 69365
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRIDLANCAMENTO_VAZIO", gErr)
        
        Case 69366
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FISCAL_GRID_NAO_PREENCHIDO", gErr, iIndice)
        
        Case 69367
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UF_GRID_NAO_PREENCHIDO", gErr, iIndice)
        
        Case 69368
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORCONTABIL_GRID_NAO_PREENCHIDO", gErr, iIndice)
        
        Case 69369, 69370
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159216)

    End Select

    Exit Function

End Function

Private Function Move_GridLancamento_Memoria(objLivRegES As ClassLivRegES) As Long
'Move o que está no Grid de Lancamentos para o objLivRegES

Dim objLivRegESLinha As New ClassLivRegESLinha
Dim sContaFormatada As String
Dim iContaPreenchida As Integer
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Move_GridLancamento_Memoria
    
    objLivRegES.sUF = UFFornecedor.Text
    
    'Para cada linha existente do Grid
    For iIndice = 1 To objGridLancamento.iLinhasExistentes
        
        Set objLivRegESLinha = objLivRegES.colLinhas(iIndice)
        
        'Formata conta para o BD
        If Len(Trim(GridLancamentos.TextMatrix(iIndice, iGrid_Contabil_Col))) > 0 Then
            
            'Guarda a conta
            lErro = CF("Conta_Formata", GridLancamentos.TextMatrix(iIndice, iGrid_Contabil_Col), sContaFormatada, iContaPreenchida)
            If lErro <> SUCESSO Then gError 69371
            
            objLivRegESLinha.sClassifContabil = sContaFormatada
            
        Else
            objLivRegESLinha.sClassifContabil = ""
        End If
        
        objLivRegESLinha.sObservacaoLivFisc = GridLancamentos.TextMatrix(iIndice, iGrid_Observacao_Col)
        objLivRegESLinha.lNumIntDocRegES = objLivRegES.lNumIntDoc
        objLivRegESLinha.dValorContabil = Arredonda_Moeda(StrParaDbl(GridLancamentos.TextMatrix(iIndice, iGrid_ValorContabil_Col)))
        objLivRegESLinha.dValorTotal = objLivRegESLinha.dValorContabil
        
    Next

    Move_GridLancamento_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_GridLancamento_Memoria:

    Move_GridLancamento_Memoria = gErr
    
    Select Case gErr

        Case 69371
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159217)

    End Select
        
    Exit Function

End Function

Private Sub TipoNFiscal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoNFiscal_Click()

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_TipoNFiscal_Click

    iAlterado = REGISTRO_ALTERADO

    If TipoNFiscal.ListIndex = -1 Then Exit Sub

    objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)

    'Lê o Tipo de Documento
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 69372

    'Se não achou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 69373

    iEmitenteNF = objTipoDocInfo.iEmitente
    
    'Coloca na Tela a Natureza Operação Padrão
    If objTipoDocInfo.iOrigem = DOCINFO_CLIENTE Then
    
        If iEmitenteNF <> DOCINFO_EMPRESA Then
        
            NaturezaOp.Text = NaturezaOp_Conv34(objTipoDocInfo.sNatOpExtPadrao, IIf(MaskedParaDate(DataEntrada) = DATA_NULA, MaskedParaDate(DataEmissao), MaskedParaDate(DataEntrada)))
            
        End If
            
        Call Habilita_Cliente
    
    End If
        
    If objTipoDocInfo.iOrigem = DOCINFO_FORNECEDOR Then
        
        If iEmitenteNF <> DOCINFO_EMPRESA Then
        
            NaturezaOp.Text = NaturezaOp_Conv34(objTipoDocInfo.sNatOpExtPadrao, IIf(MaskedParaDate(DataEntrada) = DATA_NULA, MaskedParaDate(DataEmissao), MaskedParaDate(DataEntrada)))
            
        End If
            
        Call Habilita_Fornecedor
    
    End If
    
    ''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.TipoNFiscal_Definir(objTipoDocInfo.iCodigo, objTipoDocInfo.sSigla)
    '*** fim tributacao

    Exit Sub

Erro_TipoNFiscal_Click:

    Select Case gErr

        Case 69372

        Case 69373
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159218)

    End Select

    Exit Sub

End Sub

Private Sub Habilita_Cliente()
'Desabilita o Fornecededor e Habilita o Cliente

Dim bCancel As Boolean
    
    If giTipoVersao = VERSAO_LIGHT Then
        LabelFilial.Visible = False
        Filial.Visible = False
    End If
    
    ClienteLabel.Visible = True
    Cliente.Visible = True
    FornecedorLabel.Visible = False
    Fornecedor.Visible = False
    iClienteAlterado = REGISTRO_ALTERADO
    Call Cliente_Validate(bCancel)

End Sub

Private Sub Habilita_Fornecedor()
'Desabilita o Cliente e habilita o Fornecedor

Dim bCancel As Boolean

    If giTipoVersao = VERSAO_LIGHT Then
        LabelFilial.Visible = True
        Filial.Visible = True
    End If

    FornecedorLabel.Visible = True
    Fornecedor.Visible = True
    Cliente.Visible = False
    ClienteLabel.Visible = False
    iFornecedorAlterado = REGISTRO_ALTERADO
    Call Fornecedor_Validate(bCancel)
    
End Sub

Private Sub Serie_Change()

    iAlterado = REGISTRO_ALTERADO
    iSerieAlterada = 1

End Sub

Private Sub Serie_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Serie_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Serie_Validate

    If iSerieAlterada = 0 Then Exit Sub
    
    'Verifica se foi preenchida
    If Len(Trim(Serie.Text)) = 0 Then Exit Sub

    'Verifica se foi selecionada
    If Serie.Text = Serie.List(Serie.ListIndex) Then Exit Sub

    'Tenta selecionar a serie
    lErro = Combo_Item_Igual(Serie)
    If lErro <> SUCESSO And lErro <> 12253 Then gError 69374

    'Se a serie não está cadastrada --> verifica se é uma nota interna, se for ==> erro, pois a serie não está cadastrada
    If lErro <> SUCESSO Then

        'Se o tipo está preenchido
        If Len(Trim(TipoNFiscal.Text)) > 0 Then
    
            objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)
            'Lê o Tipo de Documento
            lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
            If lErro <> SUCESSO Then gError 69375
    
            If objTipoDocInfo.iTipo = DOCINFO_NF_INT_ENTRADA Then gError 69376
    
        End If
        
        'Verifica o tamanho da série se maior que 3
        If Len(Trim(Serie.Text)) > 3 Then gError 69377

    End If
    
    iSerieAlterada = 0
    
    Exit Sub

Erro_Serie_Validate:

    Cancel = True

    Select Case gErr

        Case 69374, 69375

        Case 69376
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, Serie.Text)

        Case 69377
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_SERIE", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159219)

    End Select

    Exit Sub

End Sub

Private Sub NFiscal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NFiscal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NFiscal, iAlterado)

End Sub

Private Sub NFiscalLabel_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objLivRegES As New ClassLivRegES
Dim sSelecao As String

On Error GoTo Erro_NFiscalLabel_Click

    'Recolhe os dados da NFiscal que estão na tela
    lErro = Trata_Filtro_Browse(colSelecao, sSelecao)
    If lErro <> SUCESSO Then gError 78005
        
    'Chama a Tela passando o Filtro
    Call Chama_Tela("EdicaoRegEntrada_Lista", colSelecao, objLivRegES, objEventoNFiscal, sSelecao)
    
    Exit Sub

Erro_NFiscalLabel_Click:

    Select Case gErr

        Case 78005

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159220)

    End Select

    Exit Sub

End Sub

Function Trata_Filtro_Browse(colSelecao As Collection, sSelecao As String) As Long

Dim lErro As Long
Dim iCampoPreenchido As Integer
Dim objCliente As New ClassCliente
Dim objFornecedor As New ClassFornecedor
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Trata_Filtro_Browse

    'Se a Série foi preenchida
    If Len(Trim(Serie.Text)) > 0 Then

        If ISSerieEletronica(Serie.Text) Then
            sSelecao = sSelecao & "Serie = ? AND Modelo = ?"
        Else
            sSelecao = sSelecao & "Serie = ? AND Modelo <> ?"
        End If
        iCampoPreenchido = 1

        'Adiciona em no Filtro
        colSelecao.Add Desconverte_Serie_Eletronica(Serie.Text)
        colSelecao.Add 55
        
    End If

    'Se o o Número da Nota Fiscal foi preenchido
    If Len(Trim(NFiscal.ClipText)) > 0 Then

        If iCampoPreenchido = 1 Then
            sSelecao = sSelecao & " AND NumNotaFiscal = ?"
        Else
            sSelecao = sSelecao & " NumNotaFiscal = ?"
            iCampoPreenchido = 1
        End If

        'Adiciona em no Filtro, Produto < ProdutoFinal
        colSelecao.Add CLng(NFiscal.Text)

    End If

    If Len(Trim(TipoNFiscal.Text)) > 0 Then
    
        objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)
    
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO Then gError 78002
                
        If objTipoDocInfo.iOrigem = EMITENTE_FORNECEDOR Then
        
            'Se o Cliente ou Fornecedor foram preenchidos
            If Len(Trim(Fornecedor.ClipText)) > 0 Then
            
                objFornecedor.sNomeReduzido = Fornecedor.Text
                
                'Le a Razão Social do Fornecedor
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO Then gError 78003
                
                If iCampoPreenchido = 1 Then
                    sSelecao = sSelecao & " AND Nome = ?"
                Else
                    sSelecao = sSelecao & " Nome = ?"
                    iCampoPreenchido = 1
                End If
        
                'Adiciona em no Filtro, Produto < ProdutoFinal
                colSelecao.Add objFornecedor.sRazaoSocial
                    
            End If
            
        ElseIf objTipoDocInfo.iOrigem = EMITENTE_CLIENTE Then
        
            'Se o Cliente ou Fornecedor foram preenchidos
            If Len(Trim(Cliente.ClipText)) > 0 Then
            
                objCliente.sNomeReduzido = Cliente.Text
                
                'Le a Razão Social do Cliente
                lErro = CF("Cliente_Le_NomeReduzido", objCliente)
                If lErro <> SUCESSO Then gError 78004
                
                If iCampoPreenchido = 1 Then
                    sSelecao = sSelecao & " AND Nome = ?"
                Else
                    sSelecao = sSelecao & " Nome = ?"
                    iCampoPreenchido = 1
                End If
        
                'Adiciona em no Filtro, Produto < ProdutoFinal
                colSelecao.Add objCliente.sRazaoSocial
            
            End If
            
        End If
        
    End If

    'Se o o Número da Nota Fiscal foi preenchido
    If Len(Trim(DataEmissao.ClipText)) > 0 Then

        If iCampoPreenchido = 1 Then
            sSelecao = sSelecao & " AND DataEmissao = ?"
        Else
            sSelecao = sSelecao & " DataEmissao = ?"
            iCampoPreenchido = 1
        End If

        'Adiciona em no Filtro, Produto < ProdutoFinal
        colSelecao.Add CDate(DataEmissao.Text)

    End If

    Trata_Filtro_Browse = SUCESSO

    Exit Function

Erro_Trata_Filtro_Browse:

    Trata_Filtro_Browse = gErr

    Select Case gErr

        Case 78002, 78003, 78004

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159221)

    End Select

    Exit Function

End Function

Function Move_Identificacao_Memoria(objLivRegES As ClassLivRegES) As Long
'Move os dados que identificam a Nota Fiscal para o objNFiscal

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objCliente As New ClassCliente

On Error GoTo Erro_Move_Identificacao_Memoria
    
    If Len(Trim(DataEmissao.ClipText)) > 0 Then
        objLivRegES.dtDataEmissao = CDate(DataEmissao.Text)
    Else
        objLivRegES.dtDataEmissao = DATA_NULA
    End If

    If Len(Trim(DataEntrada.ClipText)) > 0 Then
        objLivRegES.dtData = CDate(DataEntrada.Text)
    Else
        objLivRegES.dtData = DATA_NULA
    End If

    objLivRegES.sSerie = Serie.Text
    
    If Len(Trim(NFiscal.ClipText)) > 0 Then
        objLivRegES.lNumNotaFiscal = NFiscal.Text
    End If
    
    objLivRegES.iFilialEmpresa = giFilialEmpresa
    
    Move_Identificacao_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Identificacao_Memoria:

    Move_Identificacao_Memoria = gErr
    
    Select Case gErr
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159222)

    End Select
            
    Exit Function

End Function

Private Sub LblNaturezaOp_Click()

Dim objNaturezaOp As New ClassNaturezaOp
Dim colSelecao As New Collection
Dim dtDataRef As Date

    'Se NaturezaOP estiver preenchida coloca no Obj
    If Len(Trim(NaturezaOp.ClipText)) > 0 Then objNaturezaOp.sCodigo = NaturezaOp.Text

    If Len(DataEmissao) > 0 Then
        dtDataRef = MaskedParaDate(DataEmissao)
    Else
        dtDataRef = DATA_NULA
    End If
        
    If dtDataRef < DATA_INICIO_CFOP4 Then
    
        If Len(Trim(DataEntrada)) > 0 Then
            dtDataRef = MaskedParaDate(DataEntrada)
        End If
    
    End If
    
    'Adiciona os limites de Natureza
    Select Case iEmitenteNF
    
        Case DOCINFO_EMPRESA
            colSelecao.Add NATUREZA_ENTRADA_COD_INICIAL
            colSelecao.Add NATUREZA_ENTRADA_COD_FINAL
            
        Case Else
            colSelecao.Add NATUREZA_SAIDA_COD_INICIAL
            colSelecao.Add NATUREZA_SAIDA_COD_FINAL
            
    End Select
    
    'Chama a Tela de browse de NaturezaOp
    Call Chama_Tela("NaturezaOpLista", colSelecao, objNaturezaOp, objEventoNatureza, "{fn LENGTH(Codigo) } = " & IIf(dtDataRef < DATA_INICIO_CFOP4, "3", "4"))

End Sub

Private Sub objEventoNatureza_evSelecao(obj1 As Object)

Dim objNaturezaOp As ClassNaturezaOp

    Set objNaturezaOp = obj1

    'Coloca a natureza Operação na tela
    NaturezaOp.Text = objNaturezaOp.sCodigo

    Me.Show

End Sub

Private Sub ClienteLabel_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    'Preenche NomeReduzido com o cliente da tela
    objCliente.sNomeReduzido = Cliente.Text

    'Chama Tela ClienteLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub Cliente_Change()

    iAlterado = REGISTRO_ALTERADO
    iClienteAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate

    If iClienteAlterado = 1 Then

        If Len(Trim(Cliente.Text)) > 0 Then

            lErro = TP_Cliente_Le3(Cliente, objCliente, iCodFilial)
            If lErro <> SUCESSO Then gError 69385

            lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
            If lErro <> SUCESSO Then gError 69386

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)

            'Seleciona filial na Combo Filial
            If colCodigoNome.Count = 1 Or iCodFilial <> 0 Then
            
                If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ
                
                Call CF("Filial_Seleciona", Filial, iCodFilial)
            
            ''*** incluidos p/tratamento de tributacao *******************************
                Call gobjTribTab.FilialCliente_Alterada(objCliente.lCodigo, iCodFilial)
            '*** fim tributacao
            
            End If
            
        ElseIf Len(Trim(Cliente.Text)) = 0 Then

            Filial.Clear

        End If

        iClienteAlterado = 0

    End If

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True
    
    Select Case gErr

        Case 69385
            
        Case 69386

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159223)

    End Select

    Exit Sub

End Sub

Private Sub Fornecedor_Change()

    iAlterado = REGISTRO_ALTERADO
    iFornecedorAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Validate

    If iFornecedorAlterado = 1 Then

        If Len(Trim(Fornecedor.Text)) > 0 Then

            'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
            lErro = TP_Fornecedor_Le3(Fornecedor, objFornecedor, iCodFilial)
            If lErro <> SUCESSO Then gError 69387

            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then gError 69388

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)

            If colCodigoNome.Count = 1 Or iCodFilial <> 0 Then
            
                If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ
                
                'Seleciona filial na Combo Filial
                Call CF("Filial_Seleciona", Filial, iCodFilial)
                
''*** incluidos p/tratamento de tributacao *******************************
                Call gobjTribTab.FilialFornecedor_Alterada(objFornecedor.lCodigo, iCodFilial)
'*** fim tributacao

            End If

        ElseIf Len(Trim(Fornecedor.Text)) = 0 Then

            Filial.Clear

        End If

        iFornecedorAlterado = 0

    End If

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True

    Select Case gErr

        Case 69387, 69388

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159224)

    End Select

    Exit Sub

End Sub

Private Sub FornecedorLabel_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'recolhe o Nome Reduzido da tela
    objFornecedor.sNomeReduzido = Fornecedor.Text

    'Chama a Tela de browse Fornecedores
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor
Dim bCancel As Boolean

    Set objFornecedor = obj1

    'Coloca o Fornecedor na tela
    Fornecedor.Text = objFornecedor.lCodigo

    Call Fornecedor_Validate(bCancel)

    Me.Show

End Sub


Private Sub SerieLabel_Click()

Dim objSerie As New ClassSerie
Dim colSelecao As Collection

    'recolhe a serie da tela
    objSerie.sSerie = Serie.Text

    'Chama a Tela de Browse SerieLista
    Call Chama_Tela("SerieLista", colSelecao, objSerie, objEventoSerie)

End Sub

Private Sub objEventoSerie_evSelecao(obj1 As Object)

Dim objSerie As ClassSerie

    Set objSerie = obj1

    'Coloca a Série na Tela
    Serie.Text = objSerie.sSerie

    Me.Show

End Sub

Private Sub objEventoNFiscal_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objLivRegES As New ClassLivRegES

On Error GoTo Erro_objEventoNFiscal_evSelecao

    Set objLivRegES = obj1
    
    'Traz a NotaFiscal de entrada para a a tela
    lErro = Traz_EdicaoRegEntrada_Tela(objLivRegES)
    If lErro <> SUCESSO Then gError 69390
    
    'Fecha o Comando de Setas
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoNFiscal_evSelecao:

    Select Case gErr

        Case 69390

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159225)

    End Select

    Exit Sub

End Sub

Private Sub TipoNFiscal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_TipoNFiscal_Validate

    'Verifica se o tipo está preenchido
    If Len(Trim(TipoNFiscal.Text)) = 0 Then Exit Sub
    
    'Verifica se foi selecionado
    If TipoNFiscal.List(TipoNFiscal.ListIndex) = TipoNFiscal.Text Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(TipoNFiscal, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 69391

    'Se não conseguir --> Erro
    If lErro <> SUCESSO Then gError 69392

    'tenta ler a natureza de operacao
    objTipoDocInfo.iCodigo = TipoNFiscal.ItemData(TipoNFiscal.ListIndex)
        
    'Lê o Tipo da NF
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 69393
    
    'Se não achou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 69394

    'Se a Natureza de Operação ainda não foi preenchida
    If Len(Trim(NaturezaOp.ClipText)) = 0 Then

        NaturezaOp.Text = objTipoDocInfo.sNaturezaOperacaoPadrao

    End If
    
    Exit Sub

Erro_TipoNFiscal_Validate:

    Cancel = True

    Select Case gErr

        Case 69391, 69393

        Case 69392
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, TipoNFiscal.Text)

        Case 69394
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159226)

    End Select

    Exit Sub

End Sub

Sub Desabilita_DadosPrincipais()

    TipoNFiscal.Enabled = False
    Serie.Enabled = False
    NFiscal.Enabled = False
    SerieLabel.Enabled = False
    Cliente.Enabled = False
    Fornecedor.Enabled = False
    Filial.Enabled = False
    DataEmissao.Enabled = False
'    DataEntrada.Enabled = False
    FornecedorLabel.Enabled = False
    ClienteLabel.Enabled = False
    NFiscalLabel.Enabled = False
    TipoLabel.Enabled = False
    LabelFilial.Enabled = False
'    DataEntradaLabel.Enabled = False
    DataEmissaoLabel.Enabled = False
    UpDownEmissao.Enabled = False
'    UpDownEntrada.Enabled = False
    
End Sub

Sub Habilita_DadosPrincipais()

    TipoNFiscal.Enabled = True
    SerieLabel.Enabled = True
    Serie.Enabled = True
    NFiscal.Enabled = True
    NaturezaOp.Enabled = True
    Cliente.Enabled = True
    Fornecedor.Enabled = True
    Filial.Enabled = True
    DataEmissao.Enabled = True
    DataEntrada.Enabled = True
    FornecedorLabel.Enabled = True
    ClienteLabel.Enabled = True
    NFiscalLabel.Enabled = True
    TipoLabel.Enabled = True
    LabelFilial.Enabled = True
    DataEntradaLabel.Enabled = True
    DataEmissaoLabel.Enabled = True
    UpDownEmissao.Enabled = True
    UpDownEntrada.Enabled = True
    
End Sub

Function Traz_EdicaoRegEntrada_Tela(objLivRegES As ClassLivRegES) As Long
'Traz os dados do Livro de Reg entrada para a Tela

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim bCancel As Boolean
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_Traz_EdicaoRegEntrada_Tela

    'Limpa a Tela (nao pode fechar as setas)
    Call Limpa_Tela_EdicaoRegEntrada
    Call Desabilita_DadosPrincipais
    
    objNFiscal.lNumIntDoc = objLivRegES.lNumIntDocOrigem
        
    'Tenta ler a nota Fiscal passada por parametro
    lErro = CF("NFiscal_Le", objNFiscal)
    If lErro <> SUCESSO And lErro <> 31442 Then gError 69402
        
    'Se não encontrou a NFiscal --> erro
    If lErro <> SUCESSO Then gError 69403
    
    objLivRegES.lNumIntNF = objNFiscal.lNumIntDoc
    
    'Le o Livro Fiscal de Acordo com a Nota Fiscal
    lErro = CF("LivRegES_Le_NotaFiscal", objLivRegES)
    If lErro <> SUCESSO And lErro <> 69464 Then gError 69406
    
    'Se não encontrar --> Erro
    If lErro = 69464 Then gError 69505
    
    'Lê os Ítens da Nota Fiscal
    lErro = CF("NFiscalItens_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 69395
            
    'Lê a parte de Tributação
    lErro = CF("NFiscal_Le_Tributacao", objNFiscal)
    If lErro <> SUCESSO Then gError 69396

    'Lê a parte de tributacao dos Ítens da Nota Fiscal
    lErro = CF("NFiscalItens_Le_Trib", objNFiscal)
    If lErro <> SUCESSO Then gError 69397
    
    If objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMP Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPSE Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPPC Then
    
        'carregar os dados da importacao
        lErro = CF("NFiscal_Le_NFImportacao", objNFiscal)
        If lErro <> SUCESSO And lErro <> 184739 Then gError 69397
        
    End If
    
    'Pega a tributacao do Livro e não da Nota
    Call Preenche_Tributacao_ItemRegES(objLivRegES, objNFiscal)
    
'    'Preenche o Grid de Lancamentos
'    lErro = Preenche_GridLancamentos(objLivRegES)
'    If lErro <> SUCESSO Then gError 69398
    
    If objLivRegES.iStatusLivro = STATUS_FIS_ORIGINAL Then
        CheckAlterada.Value = vbUnchecked
        CheckIgnorar.Value = vbUnchecked
    ElseIf objLivRegES.iStatusLivro = STATUS_FIS_ORIGINAL_EXCLUIDA Then
        CheckAlterada.Value = vbUnchecked
        CheckIgnorar.Value = vbChecked
    ElseIf objLivRegES.iStatusLivro = STATUS_FIS_ALTERADA Then
        CheckAlterada.Value = vbChecked
        CheckIgnorar.Value = vbUnchecked
    ElseIf objLivRegES.iStatusLivro = STATUS_FIS_ALTERADA_EXCLUIDA Then
        CheckAlterada.Value = vbChecked
        CheckIgnorar.Value = vbChecked
    End If
    
    'Preenche o Status da Nota Fiscal
    If objNFiscal.iStatus = STATUS_LANCADO Then
        Status.Caption = STRING_STATUS_LANCADO
    ElseIf objNFiscal.iStatus = STATUS_BAIXADO Then
        Status.Caption = STRING_STATUS_BAIXADO
    ElseIf objNFiscal.iStatus = STATUS_CANCELADO Then
        Status.Caption = STRING_STATUS_CANCELADO
    End If
        
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'Lê o Tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 69399

    'Se não encontrou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 69400

    'Coloca o Tipo na tela
    TipoNFiscal.Text = objTipoDocInfo.iCodigo
    Call TipoNFiscal_Validate(bCancel)
    
    'De acordo com a Origem do tipo Coloca o Cliente ou o fornecedor na tela
    If objTipoDocInfo.iOrigem = DOCINFO_CLIENTE Then
    
        Call Habilita_Cliente
        Cliente.Text = objNFiscal.lCliente
        Call Cliente_Validate(bCancel)
        Filial.Text = objNFiscal.iFilialCli
        Call Filial_Validate(bCancel)
    ElseIf objTipoDocInfo.iOrigem = DOCINFO_FORNECEDOR Then
    
        Call Habilita_Fornecedor
        Fornecedor.Text = objNFiscal.lFornecedor
        Call Fornecedor_Validate(bCancel)
        Filial.Text = objNFiscal.iFilialForn
        Call Filial_Validate(bCancel)
        
    End If
    
    'Preenhce  a tela com os dados vindos em objNFiscal
    NaturezaOp.Text = objNFiscal.sNaturezaOp
    Serie.Text = objNFiscal.sSerie
    NFiscal.Text = objNFiscal.lNumNotaFiscal

    'Preenche o Grid de Lancamentos
    lErro = Preenche_GridLancamentos(objLivRegES)
    If lErro <> SUCESSO Then gError 69398
    
    'Apuração
    Call DateParaMasked(DataEmissao, objNFiscal.dtDataEmissao)
    Call DateParaMasked(DataEntrada, objLivRegES.dtData)
    
    'Valores Auxiliar para a Tributacao
    SubTotal.Caption = Format(objNFiscal.dValorProdutos, "Standard")
    ValorFrete.Text = Format(objNFiscal.dValorFrete, "Standard")
    ValorSeguro.Text = Format(objNFiscal.dValorSeguro, "Standard")
    ValorDespesas.Text = Format(objNFiscal.dValorOutrasDespesas, "Standard")
    ValorDesconto.Text = Format(objNFiscal.dValorDesconto, "Standard")
    Total.Text = Format(objNFiscal.dValorTotal, "Standard")
    
    UFFornecedor.Text = objLivRegES.sUF
    
    ''*** incluidos p/tratamento de tributacao *******************************
'???    gobjTribTab.mvariFrameAtual = 2
    
    lErro = gobjTribTab.Traz_NFiscal_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 69401
    
    '*** fim tributacao
        
    Set gobjNFiscal = objNFiscal
    Set gobjLivRegES = objLivRegES

    Traz_EdicaoRegEntrada_Tela = SUCESSO

    iAlterado = 0

    Exit Function

Erro_Traz_EdicaoRegEntrada_Tela:

    Traz_EdicaoRegEntrada_Tela = gErr

    Select Case gErr
        
        Case 69395, 69396, 69397, 69398, 69399, 69406
        
        Case 69400
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)
        
        Case 69401
        
        Case 69505
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIVROENTRADA_NAO_CADASTRADO_NFISCAL", gErr, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159227)

    End Select

    Exit Function

End Function

Sub Preenche_Tributacao_ItemRegES(objLivRegES As ClassLivRegES, objNFiscal As ClassNFiscal)

Dim objLivRegESItemNF As New ClassLivRegESItemNF
Dim objItemNF As New ClassItemNF
    
    For Each objLivRegESItemNF In objLivRegES.ColItensNF
    
        For Each objItemNF In objNFiscal.ColItensNF
            
            'Caso seja um item
            If objLivRegESItemNF.iNumItem = objItemNF.iItem Then
        
                objItemNF.objTributacaoItemNF.dICMSAliquota = objLivRegESItemNF.dAliquotaICMS
                objItemNF.objTributacaoItemNF.dICMSBase = objLivRegESItemNF.dValorICMSBase
                objItemNF.objTributacaoItemNF.dICMSSubstBase = objLivRegESItemNF.dValorICMSSubstBase
                objItemNF.objTributacaoItemNF.dIPIValor = objLivRegESItemNF.dValorIPI
                objItemNF.objTributacaoItemNF.iICMSTipo = objLivRegESItemNF.iTipoTribICMS
                objItemNF.objTributacaoItemNF.iIPITipo = objLivRegESItemNF.iTipoTribIPI
                objItemNF.objTributacaoItemNF.dIPIAliquota = objLivRegESItemNF.dAliquotaIPI
                objItemNF.objTributacaoItemNF.dICMSPercRedBase = objLivRegESItemNF.dRedBaseICMS
                objItemNF.objTributacaoItemNF.dIPIPercRedBase = objLivRegESItemNF.dRedBaseIPI
                objItemNF.objTributacaoItemNF.iTipoTributacao = objLivRegESItemNF.iTipoTributacao
                objItemNF.objTributacaoItemNF.dICMSSubstAliquota = objLivRegESItemNF.dAliquotaSubst
                objItemNF.objTributacaoItemNF.dICMSSubstValor = objLivRegESItemNF.dValorSubst
                objItemNF.objTributacaoItemNF.dICMSValor = objLivRegESItemNF.dValorICMS
                objItemNF.objTributacaoItemNF.dIPIBaseCalculo = objLivRegESItemNF.dValorIPIBase
                objItemNF.objTributacaoItemNF.sNaturezaOp = objLivRegESItemNF.sCFOP
                objItemNF.objTributacaoItemNF.dICMSCredito = objLivRegESItemNF.dValorICMSCredito
                
                Exit For
                        
            End If
        
        Next
        
        'XXXX
'
'        'Caso seja frete
'        If objLivRegESItemNF.iNumItem = 991 Then
'
''            Set objNFiscal.objTributacaoNF.objTributacaoFrete = New ClassTribComplNF
'
'            objNFiscal.objTributacaoNF.objTributacaoFrete.dICMSAliquota = objLivRegESItemNF.dAliquotaICMS
'            objNFiscal.objTributacaoNF.objTributacaoFrete.dICMSBase = objLivRegESItemNF.dValorICMSBase
'            objNFiscal.objTributacaoNF.objTributacaoFrete.dICMSSubstBase = objLivRegESItemNF.dValorICMSSubstBase
'            objNFiscal.objTributacaoNF.objTributacaoFrete.dIPIValor = objLivRegESItemNF.dValorIPI
'            objNFiscal.objTributacaoNF.objTributacaoFrete.iICMSTipo = objLivRegESItemNF.iTipoTribICMS
'            objNFiscal.objTributacaoNF.objTributacaoFrete.iIPITipo = objLivRegESItemNF.iTipoTribIPI
'            objNFiscal.objTributacaoNF.objTributacaoFrete.dIPIAliquota = objLivRegESItemNF.dAliquotaIPI
'            objNFiscal.objTributacaoNF.objTributacaoFrete.dICMSPercRedBase = objLivRegESItemNF.dRedBaseICMS
'            objNFiscal.objTributacaoNF.objTributacaoFrete.dIPIPercRedBase = objLivRegESItemNF.dRedBaseIPI
'            objNFiscal.objTributacaoNF.objTributacaoFrete.iTipoTributacao = objLivRegESItemNF.iTipoTributacao
'            objNFiscal.objTributacaoNF.objTributacaoFrete.dICMSSubstAliquota = objLivRegESItemNF.dAliquotaSubst
'            objNFiscal.objTributacaoNF.objTributacaoFrete.dICMSSubstValor = objLivRegESItemNF.dValorSubst
'            objNFiscal.objTributacaoNF.objTributacaoFrete.dICMSValor = objLivRegESItemNF.dValorICMS
'            objNFiscal.objTributacaoNF.objTributacaoFrete.dIPIBaseCalculo = objLivRegESItemNF.dValorIPIBase
'            objNFiscal.objTributacaoNF.objTributacaoFrete.sNaturezaOp = objLivRegESItemNF.sCFOP
'
'        'Caso seja Outras
'        ElseIf objLivRegESItemNF.iNumItem = 999 Then
'
''            Set objNFiscal.objTributacaoNF.objTributacaoOutras = New ClassTribComplNF
'
'            objNFiscal.objTributacaoNF.objTributacaoOutras.dICMSAliquota = objLivRegESItemNF.dAliquotaICMS
'            objNFiscal.objTributacaoNF.objTributacaoOutras.dICMSBase = objLivRegESItemNF.dValorICMSBase
'            objNFiscal.objTributacaoNF.objTributacaoOutras.dICMSSubstBase = objLivRegESItemNF.dValorICMSSubstBase
'            objNFiscal.objTributacaoNF.objTributacaoOutras.dIPIValor = objLivRegESItemNF.dValorIPI
'            objNFiscal.objTributacaoNF.objTributacaoOutras.iICMSTipo = objLivRegESItemNF.iTipoTribICMS
'            objNFiscal.objTributacaoNF.objTributacaoOutras.iIPITipo = objLivRegESItemNF.iTipoTribIPI
'            objNFiscal.objTributacaoNF.objTributacaoOutras.dIPIAliquota = objLivRegESItemNF.dAliquotaIPI
'            objNFiscal.objTributacaoNF.objTributacaoOutras.dICMSPercRedBase = objLivRegESItemNF.dRedBaseICMS
'            objNFiscal.objTributacaoNF.objTributacaoOutras.dIPIPercRedBase = objLivRegESItemNF.dRedBaseIPI
'            objNFiscal.objTributacaoNF.objTributacaoOutras.iTipoTributacao = objLivRegESItemNF.iTipoTributacao
'            objNFiscal.objTributacaoNF.objTributacaoOutras.dICMSSubstAliquota = objLivRegESItemNF.dAliquotaSubst
'            objNFiscal.objTributacaoNF.objTributacaoOutras.dICMSSubstValor = objLivRegESItemNF.dValorSubst
'            objNFiscal.objTributacaoNF.objTributacaoOutras.dICMSValor = objLivRegESItemNF.dValorICMS
'            objNFiscal.objTributacaoNF.objTributacaoOutras.dIPIBaseCalculo = objLivRegESItemNF.dValorIPIBase
'            objNFiscal.objTributacaoNF.objTributacaoOutras.sNaturezaOp = objLivRegESItemNF.sCFOP
'
'        'Caso seja Seguro
'        ElseIf objLivRegESItemNF.iNumItem = 992 Then
'
''            Set objNFiscal.objTributacaoNF.objTributacaoSeguro = New ClassTribComplNF
'
'            objNFiscal.objTributacaoNF.objTributacaoSeguro.dICMSAliquota = objLivRegESItemNF.dAliquotaICMS
'            objNFiscal.objTributacaoNF.objTributacaoSeguro.dICMSBase = objLivRegESItemNF.dValorICMSBase
'            objNFiscal.objTributacaoNF.objTributacaoSeguro.dICMSSubstBase = objLivRegESItemNF.dValorICMSSubstBase
'            objNFiscal.objTributacaoNF.objTributacaoSeguro.dIPIValor = objLivRegESItemNF.dValorIPI
'            objNFiscal.objTributacaoNF.objTributacaoSeguro.iICMSTipo = objLivRegESItemNF.iTipoTribICMS
'            objNFiscal.objTributacaoNF.objTributacaoSeguro.iIPITipo = objLivRegESItemNF.iTipoTribIPI
'            objNFiscal.objTributacaoNF.objTributacaoSeguro.dIPIAliquota = objLivRegESItemNF.dAliquotaIPI
'            objNFiscal.objTributacaoNF.objTributacaoSeguro.dICMSPercRedBase = objLivRegESItemNF.dRedBaseICMS
'            objNFiscal.objTributacaoNF.objTributacaoSeguro.dIPIPercRedBase = objLivRegESItemNF.dRedBaseIPI
'            objNFiscal.objTributacaoNF.objTributacaoSeguro.iTipoTributacao = objLivRegESItemNF.iTipoTributacao
'            objNFiscal.objTributacaoNF.objTributacaoSeguro.dICMSSubstAliquota = objLivRegESItemNF.dAliquotaSubst
'            objNFiscal.objTributacaoNF.objTributacaoSeguro.dICMSSubstValor = objLivRegESItemNF.dValorSubst
'            objNFiscal.objTributacaoNF.objTributacaoSeguro.dICMSValor = objLivRegESItemNF.dValorICMS
'            objNFiscal.objTributacaoNF.objTributacaoSeguro.dIPIBaseCalculo = objLivRegESItemNF.dValorIPIBase
'            objNFiscal.objTributacaoNF.objTributacaoSeguro.sNaturezaOp = objLivRegESItemNF.sCFOP
'
'        End If
'
    Next
        
End Sub

Public Function Trata_Parametros(Optional objLivRegES As ClassLivRegES) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se alguma nota foi passada por parametro
    If Not (objLivRegES Is Nothing) Then
        
        'Traz a NotaFiscal de entrada para a a tela
        lErro = Traz_EdicaoRegEntrada_Tela(objLivRegES)
        If lErro <> SUCESSO Then gError 69407

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 69402, 69406, 69407

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159228)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Preenche_GridLancamentos(objLivRegES As ClassLivRegES) As Long

Dim objLivRegESLinha As New ClassLivRegESLinha
Dim iIndice As Integer
Dim sContaEnxuta As String
Dim lErro As Long

On Error GoTo Erro_Preenche_GridLancamentos

    iIndice = 0
    Call Grid_Limpa(objGridLancamento)
        
    'Para cada Item da Coleção
    For Each objLivRegESLinha In objLivRegES.colLinhas

        iIndice = iIndice + 1
        
        'Preenche o Grid
        GridLancamentos.TextMatrix(iIndice, iGrid_AliquotaICMS_Col) = Format(objLivRegESLinha.dAliquotaICMS, "Percent")
        GridLancamentos.TextMatrix(iIndice, iGrid_AliquotaIPI_Col) = Format(objLivRegESLinha.dAliquotaIPI, "Percent")
        GridLancamentos.TextMatrix(iIndice, iGrid_BaseCalculoICMS_Col) = Format(objLivRegESLinha.dValorICMSBase, "Standard")
        GridLancamentos.TextMatrix(iIndice, iGrid_BaseCalculoIPI_Col) = Format(objLivRegESLinha.dValorIPIBase, "Standard")
        
        If objLivRegESLinha.sClassifContabil <> "" Then

            lErro = Mascara_RetornaContaEnxuta(objLivRegESLinha.sClassifContabil, sContaEnxuta)
            If lErro <> SUCESSO Then gError 69408

            ContabilLanc.PromptInclude = False
            ContabilLanc.Text = sContaEnxuta
            ContabilLanc.PromptInclude = True

            GridLancamentos.TextMatrix(iIndice, iGrid_Contabil_Col) = ContabilLanc.Text

        End If
            
        GridLancamentos.TextMatrix(iIndice, iGrid_Fiscal_Col) = objLivRegESLinha.sNaturezaOp
        GridLancamentos.TextMatrix(iIndice, iGrid_ImpostoCreditadoICMS_Col) = Format(objLivRegESLinha.dValorICMS, "Standard")
        GridLancamentos.TextMatrix(iIndice, iGrid_ImpostoCreditadoIPI_Col) = Format(objLivRegESLinha.dValorIPI, "Standard")
        GridLancamentos.TextMatrix(iIndice, iGrid_IsentaICMS_Col) = Format(objLivRegESLinha.dValorICMSIsentoNaoTrib, "Standard")
        GridLancamentos.TextMatrix(iIndice, iGrid_IsentaIPI_Col) = Format(objLivRegESLinha.dValorIPIIsentoNaoTrib, "Standard")
        GridLancamentos.TextMatrix(iIndice, iGrid_Observacao_Col) = objLivRegESLinha.sObservacaoLivFisc
        GridLancamentos.TextMatrix(iIndice, iGrid_OutrasICMS_Col) = Format(objLivRegESLinha.dValorICMSOutras, "Standard")
        GridLancamentos.TextMatrix(iIndice, iGrid_OutrasIPI_Col) = Format(objLivRegESLinha.dValorIPIOutras, "Standard")
        GridLancamentos.TextMatrix(iIndice, iGrid_UF_Col) = objLivRegES.sUF
        GridLancamentos.TextMatrix(iIndice, iGrid_ValorContabil_Col) = Format(objLivRegESLinha.dValorContabil, "Standard")

    Next

    objGridLancamento.iLinhasExistentes = objLivRegES.colLinhas.Count
    
    Preenche_GridLancamentos = SUCESSO
    
    Exit Function
    
Erro_Preenche_GridLancamentos:

    Preenche_GridLancamentos = gErr

    Select Case gErr

        Case 69408
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objLivRegESLinha.sClassifContabil)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159229)

    End Select

    Exit Function

End Function

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    'Preenche campo Cliente
    Cliente.Text = objCliente.sNomeReduzido

    'Executa o Validate
    Call Cliente_Validate(bCancel)

    Me.Show

End Sub

Private Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Filial_Click()

    iAlterado = REGISTRO_ALTERADO

    'Verifica se alguma filial foi selecionada
    If Filial.ListIndex = -1 Then Exit Sub
     
    If Cliente.Visible = True Then
        Call Trata_FilialCliente
    Else
        Call Trata_FilialForn
    End If
    
End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objFilialCliente As New ClassFilialCliente
Dim sNomeRed As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 69409

    'Se não encontra o ítem com o código informado
    If lErro = 6730 Then

        If Fornecedor.Visible = True Then

            'Verifica de o fornecedor foi digitado
            If Len(Trim(Fornecedor.Text)) = 0 Then gError 69410

            sNomeRed = Fornecedor.Text

            objFilialFornecedor.iCodFilial = iCodigo
            
            'Pesquisa se existe Filial com o código extraído
            lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sNomeRed, objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 18272 Then gError 69411

            'Se não encontrou a Filial Fornecedor --> erro
            If lErro = 18272 Then gError 69412

            'Mostra a Filial Fornecedor na tela
            Filial.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

            Call gobjTribTab.FilialFornecedor_Alterada(objFilialFornecedor.lCodFornecedor, objFilialFornecedor.iCodFilial)

        Else

            'Verifica se o Cliente foi preenchido
            If Len(Trim(Cliente.ClipText)) = 0 Then gError 69413

            sNomeRed = Cliente.Text
            
            'Lê a Filial Cliente
            lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sNomeRed, objFilialCliente)
            If lErro <> SUCESSO And lErro <> 17660 Then gError 69414

            'Se não encontrou a Filial Cliente --> erro
            If lErro = 17660 Then gError 69415

            'Mostra a Filial Cliente na tela
            Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

            Call gobjTribTab.FilialCliente_Alterada(objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
        
        End If

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 69416

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 69411, 69409

        Case 69412
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 69410
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 69413
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 69415
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If

        Case 69416
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159230)

    End Select

    Exit Sub

End Sub

Private Function Trata_FilialCliente() As Long

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente

On Error GoTo Erro_Trata_FilialCliente
    
    If Filial.ListIndex <> -1 Then
    
        objFilialCliente.iCodFilial = Filial.ItemData(Filial.ListIndex)
        'Lê a FilialCliente
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Cliente.Text, objFilialCliente)
        If lErro <> SUCESSO Then gError 69417
        
        Call gobjTribTab.FilialCliente_Alterada(objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
    
    End If
    
    Trata_FilialCliente = SUCESSO
    
    Exit Function
    
Erro_Trata_FilialCliente:

    Trata_FilialCliente = gErr
    
    Select Case gErr
    
        Case 69417
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159231)
            
    End Select
    
    Exit Function
    
End Function

Private Function Trata_FilialForn() As Long

Dim lErro As Long
Dim objFilialForn As New ClassFilialFornecedor

On Error GoTo Erro_Trata_FilialForn
    
    objFilialForn.iCodFilial = Codigo_Extrai(Filial.Text)
    
    If objFilialForn.iCodFilial <> 0 Then
    
        'Lê a Filial
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilialForn)
        If lErro <> SUCESSO Then gError 69418
        
        Call gobjTribTab.FilialFornecedor_Alterada(objFilialForn.lCodFornecedor, objFilialForn.iCodFilial)
    
    End If
    
    Trata_FilialForn = SUCESSO
     
    Exit Function
    
Erro_Trata_FilialForn:

    Trata_FilialForn = gErr
     
    Select Case gErr
          
        Case 69418
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159232)
     
    End Select
     
    Exit Function

End Function

Private Sub DataEmissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    'Verifica se a data de emissao foi digitada
    If Len(Trim(DataEmissao.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then gError 69419

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True

    Select Case gErr

        Case 69419

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159233)

    End Select

    Exit Sub

End Sub

Private Sub UFFornecedor_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub UFFornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_UFFornecedor_Validate

    'Verifica se foi preenchida a ComboBox Sigla
    If Len(Trim(UFFornecedor.Text)) > 0 Then

        'Verifica se existe o ítem na List da Combo. Se existir seleciona.
        lErro = Combo_Item_Igual_CI(UFFornecedor)
        If lErro <> SUCESSO And lErro <> 58583 Then gError 78011

        'Não existe o item na ComboBox Sigla
        If lErro = 58583 Then gError 78012
    
    End If
        
    For iIndice = 1 To objGridLancamento.iLinhasExistentes
        GridLancamentos.TextMatrix(iIndice, iGrid_UF_Col) = UFFornecedor.Text
    Next
    
    If Not (gobjLivRegES Is Nothing) Then
        gobjLivRegES.sUF = UFFornecedor.Text
    End If
    
    Exit Sub

Erro_UFFornecedor_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 78011, 78013

        Case 78012
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SIGLA_ESTADO_NAO_CADASTRADA", gErr, UFFornecedor.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159234)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 69420
    
    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case gErr

        Case 69420

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159235)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 69421

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case gErr

        Case 69421

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159236)

    End Select

    Exit Sub

End Sub

Private Sub DataEntrada_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEntrada_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEntrada_Validate

    'Verifica se a data de Entrada foi digitada
    If Len(Trim(DataEntrada.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataEntrada.Text)
    If lErro <> SUCESSO Then gError 69422

    Exit Sub

Erro_DataEntrada_Validate:

    Cancel = True

    Select Case gErr

        Case 69422

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159237)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEntrada_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEntrada_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(DataEntrada, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 69423
    
    Exit Sub

Erro_UpDownEntrada_DownClick:

    Select Case gErr

        Case 69423

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159238)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEntrada_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEntrada_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEntrada, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 69424

    Exit Sub

Erro_UpDownEntrada_UpClick:

    Select Case gErr

        Case 69424

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159239)

    End Select

    Exit Sub

End Sub

Public Sub GridLancamentos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridLancamento, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridLancamento, iAlterado)
    End If

End Sub

Public Sub GridLancamentos_EnterCell()

    Call Grid_Entrada_Celula(objGridLancamento, iAlterado)

End Sub

Public Sub GridLancamentos_GotFocus()

    Call Grid_Recebe_Foco(objGridLancamento)

End Sub

Public Sub GridLancamentos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridLancamento, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridLancamento, iAlterado)
    End If

End Sub

Public Sub GridLancamentos_LeaveCell()

    Call Saida_Celula(objGridLancamento)

End Sub

Public Sub GridLancamentos_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridLancamento)
    
End Sub

Public Sub GridLancamentos_RowColChange()

    Call Grid_RowColChange(objGridLancamento)

End Sub

Public Sub Observacoes_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Observacoes_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridLancamento)

End Sub

Public Sub Observacoes_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridLancamento)

End Sub

Public Sub Observacoes_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridLancamento.objControle = Observacoes
    lErro = Grid_Campo_Libera_Foco(objGridLancamento)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'Verifica qual a coluna do Grid em questão
        Select Case objGridInt.objGrid.Col

            Case iGrid_ValorContabil_Col
                lErro = Saida_Celula_ValorContabil(objGridInt)
                If lErro <> SUCESSO Then gError 69425
            
            'Observação
            Case iGrid_Observacao_Col
                lErro = Saida_Celula_Observacao(objGridInt)
                If lErro <> SUCESSO Then gError 69425

            Case iGrid_Contabil_Col
                lErro = Saida_Celula_ContabilLanc(objGridInt)
                If lErro <> SUCESSO Then gError 69566
            
        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 69425

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 69425, 69566, 69425
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159240)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Observacao(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Observacao

    Set objGridInt.objControle = Observacoes
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 69427
        
    Saida_Celula_Observacao = SUCESSO

    Exit Function

Erro_Saida_Celula_Observacao:

    Saida_Celula_Observacao = gErr

    Select Case gErr

        Case 69427
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159241)
            
    End Select

    Exit Function

End Function

Private Function Saida_Celula_ContabilLanc(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaMascarada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim iContaPreenchida As Integer

On Error GoTo Erro_Saida_Celula_ContabilLanc
    
    Set objGridLancamento.objControle = ContabilLanc
    
    If Len(Trim(ContabilLanc.ClipText)) > 0 Then
    
        'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
        lErro = CF("ContaSimples_Critica_Modulo", sContaFormatada, ContabilLanc.ClipText, objPlanoConta, MODULO_CONTABILIDADE)
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 69561
        
        If lErro = SUCESSO Then
        
            sContaFormatada = objPlanoConta.sConta
            
            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
            If lErro <> SUCESSO Then gError 69562
            
            ContabilLanc.PromptInclude = False
            ContabilLanc.Text = sContaMascarada
            ContabilLanc.PromptInclude = True
        
        'se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then
    
            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContabilLanc.Text, sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
            If lErro <> SUCESSO And lErro <> 5700 Then gError 69563
    
            'conta não cadastrada
            If lErro = 5700 Then gError 69564
             
        End If
    
    Else
        
        ContabilLanc.PromptInclude = False
        ContabilLanc.Text = ""
        ContabilLanc.PromptInclude = True
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 69565
    
    Saida_Celula_ContabilLanc = SUCESSO

    Exit Function

Erro_Saida_Celula_ContabilLanc:

    Saida_Celula_ContabilLanc = gErr

    Select Case gErr

        Case 69561, 69563, 69565
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 69562
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 69564
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContabilLanc.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada
                
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
                            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159242)

    End Select

    Exit Function

End Function

Private Sub ContabilLanc_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ContabilLanc_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridLancamento)

End Sub

Private Sub ContabilLanc_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridLancamento)

End Sub

Private Sub ContabilLanc_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridLancamento.objControle = ContabilLanc
    lErro = Grid_Campo_Libera_Foco(objGridLancamento)
    If lErro <> SUCESSO Then Cancel = True

End Sub

'- - -  - - - - - - - Sistema de Setas - - - - - - - - - - - - - - - -

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objLivRegES As New ClassLivRegES
Dim iIndice As Integer

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "LivRegES"

    'Lê os dados da Tela Notas Fiscais a Pagar
    lErro = Move_Identificacao_Memoria(objLivRegES)
    If lErro <> SUCESSO Then gError 31423

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "FilialEmpresa", objLivRegES.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "NumIntDoc", objLivRegES.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "NumIntDocLivFechado", objLivRegES.lNumIntDocLivFechado, 0, "NumIntDocLivFechado"
    colCampoValor.Add "TipoNumIntDocOrigem", objLivRegES.iTipoNumIntDocOrigem, 0, "TipoNumIntDocOrigem"
    colCampoValor.Add "NumIntDocOrigem", objLivRegES.lNumIntDocOrigem, 0, "NumIntDocOrigem"
    colCampoValor.Add "DestinoOrigem", objLivRegES.iDestinoOrigem, 0, "DestinoOrigem"
    colCampoValor.Add "StatusLivro", objLivRegES.iStatusLivro, 0, "StatusLivro"
    colCampoValor.Add "Tipo", objLivRegES.iTipo, 0, "Tipo"
    colCampoValor.Add "CGC", objLivRegES.sCgc, STRING_CGC, "CGC"
    colCampoValor.Add "InscricaoEstadual", objLivRegES.sInscricaoEstadual, STRING_INSCR_EST, "InscricaoEstadual"
    colCampoValor.Add "Nome", objLivRegES.sNome, STRING_NOME_REGES, "Nome"
    colCampoValor.Add "DataEmissao", objLivRegES.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "Data", objLivRegES.dtData, 0, "Data"
    colCampoValor.Add "UF", objLivRegES.sNome, STRING_ESTADO, "UF"
    colCampoValor.Add "Serie", objLivRegES.sSerie, STRING_SERIE, "Serie"
    colCampoValor.Add "NumNotaFiscal", objLivRegES.lNumNotaFiscal, 0, "NumNotaFiscal"
    colCampoValor.Add "Situacao", objLivRegES.iSituacao, 0, "Situacao"
    colCampoValor.Add "Modelo", objLivRegES.iModelo, 0, "Modelo"
    colCampoValor.Add "Emitente", objLivRegES.iEmitente, 0, "Emitente"
    colCampoValor.Add "Destinatario", objLivRegES.iDestinatario, 0, "Destinatario"
    colCampoValor.Add "Origem", objLivRegES.iOrigem, 0, "Origem"
    colCampoValor.Add "NumIntEmitente", objLivRegES.lNumIntEmitente, 0, "NumIntEmitente"
    colCampoValor.Add "NumIntRegApuracaoICMS", objLivRegES.lNumIntRegApuracaoICMS, 0, "NumIntRegApuracaoICMS"
    colCampoValor.Add "NumIntRegApuracaoIPI", objLivRegES.lNumIntRegApuracaoIPI, 0, "NumIntRegApuracaoIPI"
    
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "Tipo", OP_IGUAL, TIPO_REGES_ENTRADA
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 31423

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159243)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objLivRegES As New ClassLivRegES

On Error GoTo Erro_Tela_Preenche

    objLivRegES.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor

    If objLivRegES.lNumIntDoc <> 0 Then

        objLivRegES.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
        objLivRegES.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor
        objLivRegES.lNumIntDocLivFechado = colCampoValor.Item("NumIntDocLivFechado").vValor
        objLivRegES.iTipoNumIntDocOrigem = colCampoValor.Item("TipoNumIntDocOrigem").vValor
        objLivRegES.lNumIntDocOrigem = colCampoValor.Item("NumIntDocOrigem").vValor
        objLivRegES.iDestinoOrigem = colCampoValor.Item("DestinoOrigem").vValor
        objLivRegES.iStatusLivro = colCampoValor.Item("StatusLivro").vValor
        objLivRegES.iTipo = colCampoValor.Item("Tipo").vValor
        objLivRegES.sCgc = colCampoValor.Item("CGC").vValor
        objLivRegES.sInscricaoEstadual = colCampoValor.Item("InscricaoEstadual").vValor
        objLivRegES.sNome = colCampoValor.Item("Nome").vValor
        objLivRegES.dtDataEmissao = colCampoValor.Item("DataEmissao").vValor
        objLivRegES.dtData = colCampoValor.Item("Data").vValor
        objLivRegES.sUF = colCampoValor.Item("UF").vValor
        objLivRegES.sSerie = colCampoValor.Item("Serie").vValor
        objLivRegES.lNumNotaFiscal = colCampoValor.Item("NumNotaFiscal").vValor
        objLivRegES.iSituacao = colCampoValor.Item("Situacao").vValor
        objLivRegES.iModelo = colCampoValor.Item("Modelo").vValor
        objLivRegES.iEmitente = colCampoValor.Item("Emitente").vValor
        objLivRegES.iDestinatario = colCampoValor.Item("Destinatario").vValor
        objLivRegES.iOrigem = colCampoValor.Item("Origem").vValor
        objLivRegES.lNumIntEmitente = colCampoValor.Item("NumIntEmitente").vValor
        objLivRegES.lNumIntRegApuracaoICMS = colCampoValor.Item("NumIntRegApuracaoICMS").vValor
        objLivRegES.lNumIntRegApuracaoIPI = colCampoValor.Item("NumIntRegApuracaoIPI").vValor

        'Coloca os dados da NFiscal na tela
        lErro = Traz_EdicaoRegEntrada_Tela(objLivRegES)
        If lErro <> SUCESSO Then gError 31424

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 31424

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159244)

    End Select

    Exit Sub

End Sub

Public Function ValorTotal_Calcula() As Long
'Função Auxliar para a Tributação

Dim lErro As Long

On Error GoTo Erro_ValorTotal_Calcula

    lErro = gobjTribTab.AtualizarTributacao(gobjNFiscal)
    If lErro <> SUCESSO Then gError 69460

    ValorTotal_Calcula = SUCESSO

    Exit Function

Erro_ValorTotal_Calcula:

    ValorTotal_Calcula = gErr

    Select Case gErr

        Case 69460

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159261)

    End Select

    Exit Function

End Function

'*** fim tributacao

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Edição de Registro de Entrada"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "EdicaoRegEntrada"
    
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

Public Property Let MousePointer(ByVal iTipo As Integer)
    Parent.MousePointer = iTipo
End Property

Public Property Get MousePointer() As Integer
    MousePointer = Parent.MousePointer
End Property

'**** fim do trecho a ser copiado *****

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub

Private Sub Label17_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label17(Index), Source, X, Y)
End Sub

Private Sub Label17_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label17(Index), Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label13(Index), Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13(Index), Button, Shift, X, Y)
End Sub

Private Sub Label18_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label18(Index), Source, X, Y)
End Sub

Private Sub Label18_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label18(Index), Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label15(Index), Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15(Index), Button, Shift, X, Y)
End Sub

Private Sub Label16_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label16(Index), Source, X, Y)
End Sub

Private Sub Label16_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label16(Index), Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label4(Index), Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4(Index), Button, Shift, X, Y)
End Sub


Private Sub ClienteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ClienteLabel, Source, X, Y)
End Sub

Private Sub ClienteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ClienteLabel, Button, Shift, X, Y)
End Sub

Private Sub LabelFilial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFilial, Source, X, Y)
End Sub

Private Sub LabelFilial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFilial, Button, Shift, X, Y)
End Sub

Private Sub FornecedorLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecedorLabel, Source, X, Y)
End Sub

Private Sub FornecedorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel, Button, Shift, X, Y)
End Sub

Private Sub Status_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Status, Source, X, Y)
End Sub

Private Sub Status_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Status, Button, Shift, X, Y)
End Sub

Private Sub LblNaturezaOp_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblNaturezaOp, Source, X, Y)
End Sub

Private Sub LblNaturezaOp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblNaturezaOp, Button, Shift, X, Y)
End Sub

Private Sub TipoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoLabel, Source, X, Y)
End Sub

Private Sub TipoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoLabel, Button, Shift, X, Y)
End Sub

Private Sub NFiscalLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NFiscalLabel, Source, X, Y)
End Sub

Private Sub NFiscalLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NFiscalLabel, Button, Shift, X, Y)
End Sub

Private Sub SerieLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SerieLabel, Source, X, Y)
End Sub

Private Sub SerieLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SerieLabel, Button, Shift, X, Y)
End Sub

Private Sub DataEntradaLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataEntradaLabel, Source, X, Y)
End Sub

Private Sub DataEntradaLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataEntradaLabel, Button, Shift, X, Y)
End Sub

Private Sub DataEmissaoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataEmissaoLabel, Source, X, Y)
End Sub

Private Sub DataEmissaoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataEmissaoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub

Private Sub Label43_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label43, Source, X, Y)
End Sub

Private Sub Label43_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label43, Button, Shift, X, Y)
End Sub

Private Sub Label39_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label39, Source, X, Y)
End Sub

Private Sub Label39_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label39, Button, Shift, X, Y)
End Sub

Private Sub Label35_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label35, Source, X, Y)
End Sub

Private Sub Label35_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label35, Button, Shift, X, Y)
End Sub

Private Sub Label22_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label22, Source, X, Y)
End Sub

Private Sub Label22_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label22, Button, Shift, X, Y)
End Sub

Private Sub ICMSSubstValor1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSSubstValor1, Source, X, Y)
End Sub

Private Sub ICMSSubstValor1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSSubstValor1, Button, Shift, X, Y)
End Sub

Private Sub ICMSSubstBase1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSSubstBase1, Source, X, Y)
End Sub

Private Sub ICMSSubstBase1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSSubstBase1, Button, Shift, X, Y)
End Sub

Private Sub ICMSValor1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSValor1, Source, X, Y)
End Sub

Private Sub ICMSValor1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSValor1, Button, Shift, X, Y)
End Sub

Private Sub ICMSBase1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSBase1, Source, X, Y)
End Sub

Private Sub ICMSBase1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSBase1, Button, Shift, X, Y)
End Sub

Private Sub SubTotal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SubTotal, Source, X, Y)
End Sub

Private Sub SubTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SubTotal, Button, Shift, X, Y)
End Sub

Private Sub IPIValor1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(IPIValor1, Source, X, Y)
End Sub

Private Sub IPIValor1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(IPIValor1, Button, Shift, X, Y)
End Sub

Private Sub LabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTotais, Source, X, Y)
End Sub

Private Sub LabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTotais, Button, Shift, X, Y)
End Sub

Public Sub ValorContabilLanc_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorContabilLanc_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridLancamento)

End Sub

Public Sub ValorContabilLanc_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridLancamento)

End Sub

Public Sub ValorContabilLanc_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridLancamento.objControle = ValorContabilLanc
    lErro = Grid_Campo_Libera_Foco(objGridLancamento)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_ValorContabil(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Valor Parcela que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ValorContabil

    Set objGridInt.objControle = ValorContabilLanc

    'Verifica se valor está preenchido
    If Len(ValorContabilLanc.ClipText) > 0 Then

        'Critica se valor é positivo
        lErro = Valor_NaoNegativo_Critica(ValorContabilLanc.Text)
        If lErro <> SUCESSO Then gError 26724

        ValorContabilLanc.Text = Format(ValorContabilLanc.Text, "Standard")
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 26725

    Saida_Celula_ValorContabil = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorContabil:

    Saida_Celula_ValorContabil = gErr

    Select Case gErr

        Case 26724, 26725, 51061
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158006)

    End Select

    Exit Function

End Function

