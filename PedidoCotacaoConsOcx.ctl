VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.UserControl PedidoCotacaoConsOcx 
   ClientHeight    =   9195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   KeyPreview      =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   16995
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8295
      Index           =   2
      Left            =   135
      TabIndex        =   21
      Top             =   675
      Visible         =   0   'False
      Width           =   16680
      Begin VB.ComboBox Moeda 
         Height          =   288
         Left            =   1956
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   120
         Width           =   2475
      End
      Begin VB.Frame Frame9 
         Caption         =   "Valores"
         Height          =   1245
         Index           =   1
         Left            =   135
         TabIndex        =   41
         Top             =   7005
         Width           =   8685
         Begin VB.Label DescontoPrazo 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5655
            TabIndex        =   78
            Top             =   855
            Width           =   1200
         End
         Begin VB.Label DescontoVista 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5625
            TabIndex        =   77
            Top             =   315
            Width           =   1200
         End
         Begin VB.Label ValorDespesas 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2895
            TabIndex        =   76
            Top             =   615
            Width           =   1200
         End
         Begin VB.Label ValorSeguro 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1590
            TabIndex        =   75
            Top             =   600
            Width           =   1200
         End
         Begin VB.Label ValorFrete 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   150
            TabIndex        =   74
            Top             =   600
            Width           =   1200
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Desconto a Prazo"
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
            Left            =   5565
            TabIndex        =   54
            Top             =   645
            Width           =   1515
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "IPI a Prazo"
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
            Index           =   2
            Left            =   4275
            TabIndex        =   53
            Top             =   645
            Width           =   960
         End
         Begin VB.Label IPIValorPrazo 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4185
            TabIndex        =   52
            Top             =   855
            Width           =   1200
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Total a Vista"
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
            Left            =   7320
            TabIndex        =   51
            Top             =   135
            Width           =   1245
         End
         Begin VB.Label TotalVista 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7290
            TabIndex        =   50
            Top             =   330
            Width           =   1200
         End
         Begin VB.Label IPIValorVista 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4170
            TabIndex        =   49
            Top             =   315
            Width           =   1200
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "IPI a Vista"
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
            Left            =   4305
            TabIndex        =   48
            Top             =   120
            Width           =   900
         End
         Begin VB.Label Label28 
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
            Index           =   0
            Left            =   1860
            TabIndex        =   47
            Top             =   375
            Width           =   615
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Desconto a Vista"
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
            Left            =   5565
            TabIndex        =   46
            Top             =   120
            Width           =   1470
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
            Left            =   450
            TabIndex        =   45
            Top             =   375
            Width           =   450
         End
         Begin VB.Label TotalPrazo 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7290
            TabIndex        =   44
            Top             =   840
            Width           =   1200
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
            Left            =   3105
            TabIndex        =   43
            Top             =   375
            Width           =   840
         End
         Begin VB.Label LabelTotais 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Total a Prazo"
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
            Left            =   7290
            TabIndex        =   42
            Top             =   645
            Width           =   1245
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Itens"
         Height          =   6435
         Left            =   120
         TabIndex        =   37
         Top             =   432
         Width           =   16260
         Begin MSMask.MaskEdBox TotalPrazoRS 
            Height          =   228
            Left            =   6480
            TabIndex        =   85
            Top             =   2340
            Width           =   1584
            _ExtentX        =   2805
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
         Begin MSMask.MaskEdBox TotalVIstaRS 
            Height          =   228
            Left            =   4680
            TabIndex        =   86
            Top             =   2340
            Width           =   1584
            _ExtentX        =   2805
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.TextBox DescProduto 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   390
            MaxLength       =   50
            TabIndex        =   55
            Top             =   1050
            Width           =   4000
         End
         Begin MSMask.MaskEdBox Produto 
            Height          =   225
            Left            =   495
            TabIndex        =   56
            Top             =   825
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TaxaFinanceira 
            Height          =   225
            Left            =   4185
            TabIndex        =   57
            Top             =   1395
            Width           =   1380
            _ExtentX        =   2434
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
         Begin MSMask.MaskEdBox QuantEntrega 
            Height          =   210
            Left            =   2280
            TabIndex        =   58
            Top             =   1530
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   370
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
         Begin MSMask.MaskEdBox ValorIPIVista 
            Height          =   225
            Left            =   1935
            TabIndex        =   59
            Top             =   825
            Width           =   1545
            _ExtentX        =   2725
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
            Left            =   2280
            TabIndex        =   60
            Top             =   1995
            Width           =   1410
            _ExtentX        =   2487
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
         Begin MSMask.MaskEdBox PrazoEntrega 
            Height          =   225
            Left            =   3675
            TabIndex        =   61
            Top             =   825
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
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
         Begin MSMask.MaskEdBox TotVista 
            Height          =   210
            Left            =   2280
            TabIndex        =   62
            Top             =   1770
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   370
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
         Begin MSMask.MaskEdBox TotPrazo 
            Height          =   225
            Left            =   375
            TabIndex        =   63
            Top             =   1995
            Width           =   1380
            _ExtentX        =   2434
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
         Begin MSMask.MaskEdBox PrecoVista 
            Height          =   210
            Left            =   390
            TabIndex        =   64
            Top             =   1770
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   370
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
         Begin MSMask.MaskEdBox UnidadeMed 
            Height          =   225
            Left            =   615
            TabIndex        =   65
            Top             =   1290
            Width           =   630
            _ExtentX        =   1111
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   225
            Left            =   390
            TabIndex        =   66
            Top             =   1530
            Width           =   1140
            _ExtentX        =   2011
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
         Begin MSMask.MaskEdBox ValorIPIPrazo 
            Height          =   225
            Left            =   4365
            TabIndex        =   67
            Top             =   1275
            Width           =   1590
            _ExtentX        =   2805
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
         Begin MSMask.MaskEdBox AliquotaICM 
            Height          =   225
            Left            =   7185
            TabIndex        =   68
            Top             =   1215
            Width           =   1320
            _ExtentX        =   2328
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
         Begin MSMask.MaskEdBox PrecoPrazo 
            Height          =   225
            Left            =   7335
            TabIndex        =   69
            Top             =   885
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
            Format          =   "#,##0.00###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Observacao 
            Height          =   225
            Left            =   5385
            TabIndex        =   70
            Top             =   1935
            Width           =   2070
            _ExtentX        =   3651
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
         Begin MSFlexGridLib.MSFlexGrid GridProdutos 
            Height          =   2520
            Left            =   165
            TabIndex        =   38
            Top             =   240
            Width           =   15945
            _ExtentX        =   28125
            _ExtentY        =   4445
            _Version        =   393216
            Rows            =   4
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Total A Vista"
            Height          =   15
            Left            =   4170
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   40
            Top             =   4350
            Width           =   900
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Produto         "
            Height          =   195
            Left            =   435
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   39
            Top             =   375
            Width           =   600
         End
      End
      Begin MSMask.MaskEdBox TaxaConversao 
         Height          =   312
         Left            =   6216
         TabIndex        =   82
         Top             =   108
         Width           =   924
         _ExtentX        =   1640
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00##"
         PromptChar      =   " "
      End
      Begin VB.Label Label28 
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
         Height          =   192
         Index           =   1
         Left            =   1296
         TabIndex        =   84
         Top             =   168
         Width           =   648
      End
      Begin VB.Label LabelTaxa 
         AutoSize        =   -1  'True
         Caption         =   "Taxa Conversão:"
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
         ForeColor       =   &H80000008&
         Height          =   192
         Left            =   4752
         TabIndex        =   83
         Top             =   168
         Width           =   1452
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8250
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   675
      Width           =   16620
      Begin VB.Frame Frame5 
         Caption         =   "Cabeçalho"
         Height          =   3285
         Left            =   135
         TabIndex        =   11
         Top             =   255
         Width           =   8580
         Begin VB.ComboBox TipoFrete 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "PedidoCotacaoConsOcx.ctx":0000
            Left            =   1425
            List            =   "PedidoCotacaoConsOcx.ctx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   1725
            Width           =   1035
         End
         Begin VB.Frame Frame2 
            Caption         =   "Local de Entrega"
            Height          =   930
            Left            =   165
            TabIndex        =   22
            Top             =   2085
            Width           =   8340
            Begin VB.Frame FrameTipo 
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   675
               Index           =   0
               Left            =   4635
               TabIndex        =   26
               Top             =   180
               Width           =   3495
               Begin VB.Label Label37 
                  AutoSize        =   -1  'True
                  Caption         =   "Filial:"
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
                  Left            =   540
                  TabIndex        =   28
                  Top             =   195
                  Width           =   465
               End
               Begin VB.Label FilialEmpresa 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   300
                  Left            =   1275
                  TabIndex        =   27
                  Top             =   165
                  Width           =   2145
               End
            End
            Begin VB.Frame Frame3 
               Caption         =   "Tipo"
               Enabled         =   0   'False
               Height          =   585
               Left            =   420
               TabIndex        =   23
               Top             =   225
               Width           =   3645
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
                  Left            =   1905
                  TabIndex        =   25
                  Top             =   225
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
                  Left            =   300
                  TabIndex        =   24
                  Top             =   225
                  Width           =   1515
               End
            End
            Begin VB.Frame FrameTipo 
               BorderStyle     =   0  'None
               Height          =   645
               Index           =   1
               Left            =   4785
               TabIndex        =   29
               Top             =   195
               Visible         =   0   'False
               Width           =   3495
               Begin VB.Label FornecedorLabel 
                  AutoSize        =   -1  'True
                  Caption         =   "Fornecedor:"
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
                  Index           =   1
                  Left            =   90
                  TabIndex        =   33
                  Top             =   60
                  Width           =   1035
               End
               Begin VB.Label FornecDestino 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   300
                  Left            =   1170
                  TabIndex        =   32
                  Top             =   15
                  Width           =   2145
               End
               Begin VB.Label Label32 
                  AutoSize        =   -1  'True
                  Caption         =   "Filial:"
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
                  Left            =   600
                  TabIndex        =   31
                  Top             =   330
                  Width           =   465
               End
               Begin VB.Label FilialFornec 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   300
                  Left            =   1170
                  TabIndex        =   30
                  Top             =   330
                  Width           =   2145
               End
            End
         End
         Begin VB.Label CondicaoPagamento 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5655
            TabIndex        =   72
            Top             =   1230
            Width           =   2175
         End
         Begin VB.Label Contato 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1425
            TabIndex        =   71
            Top             =   1245
            Width           =   1365
         End
         Begin VB.Label Fornecedor 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1425
            TabIndex        =   36
            Top             =   780
            Width           =   2205
         End
         Begin VB.Label Codigo 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1425
            TabIndex        =   35
            Top             =   315
            Width           =   810
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
            Height          =   195
            Left            =   450
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   20
            Top             =   375
            Width           =   930
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
            Left            =   345
            TabIndex        =   19
            Top             =   840
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
            Left            =   5130
            TabIndex        =   18
            Top             =   840
            Width           =   465
         End
         Begin VB.Label Filial 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5655
            TabIndex        =   17
            Top             =   780
            Width           =   2175
         End
         Begin VB.Label Label18 
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
            Left            =   435
            TabIndex        =   16
            Top             =   1785
            Width           =   945
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
            Left            =   660
            TabIndex        =   15
            Top             =   1290
            Width           =   735
         End
         Begin VB.Label CondicaoPagtoLabel 
            AutoSize        =   -1  'True
            Caption         =   "Cond Pagto a Prazo:"
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
            Left            =   3825
            TabIndex        =   14
            Top             =   1260
            Width           =   1770
         End
         Begin VB.Label Label24 
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
            Left            =   4620
            TabIndex        =   13
            Top             =   375
            Width           =   975
         End
         Begin VB.Label Comprador 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5655
            TabIndex        =   12
            Top             =   315
            Width           =   2145
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Datas"
         Height          =   1185
         Left            =   135
         TabIndex        =   5
         Top             =   3600
         Width           =   8580
         Begin VB.Label DataBaixa 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5580
            TabIndex        =   80
            Top             =   750
            Width           =   1110
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Data Baixa:"
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
            Left            =   4485
            TabIndex        =   79
            Top             =   795
            Width           =   1005
         End
         Begin VB.Label DataValidade 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   2145
            TabIndex        =   73
            Top             =   750
            Width           =   1110
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Data Validade:"
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
            Left            =   795
            TabIndex        =   10
            Top             =   795
            Width           =   1275
         End
         Begin VB.Label Label3 
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
            Left            =   1590
            TabIndex        =   9
            Top             =   360
            Width           =   480
         End
         Begin VB.Label Data 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   2145
            TabIndex        =   8
            Top             =   315
            Width           =   1110
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Data Emissão:"
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
            Left            =   4260
            TabIndex        =   7
            Top             =   360
            Width           =   1230
         End
         Begin VB.Label DataEmissao 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5580
            TabIndex        =   6
            Top             =   315
            Width           =   1110
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   525
      Left            =   15720
      ScaleHeight     =   465
      ScaleWidth      =   1110
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   1170
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   90
         Picture         =   "PedidoCotacaoConsOcx.ctx":0018
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Imprimir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   615
         Picture         =   "PedidoCotacaoConsOcx.ctx":011A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   8730
      Left            =   105
      TabIndex        =   3
      Top             =   360
      Width           =   16800
      _ExtentX        =   29633
      _ExtentY        =   15399
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Principais"
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
Attribute VB_Name = "PedidoCotacaoConsOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Variáveis globais
Dim iAlterado As Integer
Dim iChamaTela As Integer
Dim gobjPedidoCotacao As ClassPedidoCotacao

'GridItens
Dim bExibirColReal As Boolean
Dim objGridItens As AdmGrid
Dim iFrameTipoDestinoAtual As Integer
Dim iFrameAtual As Integer
Dim iGrid_Produto_Col As Integer
Dim iGrid_DescProduto_Col As Integer
Dim iGrid_UnidadeMed_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_PrecoVista_Col As Integer
Dim iGrid_TotalVista_Col As Integer
Dim iGrid_TotalVista_RS_Col As Integer
Dim iGrid_PrecoPrazo_Col As Integer
Dim iGrid_TotalPrazo_Col As Integer
Dim iGrid_TotalPrazo_RS_Col As Integer
Dim iGrid_TaxaFinanceira_Col As Integer
Dim iGrid_PrazoEntrega_Col As Integer
Dim iGrid_QuantEntrega_Col As Integer
Dim iGrid_AliquotaIPI_Col As Integer
Dim iGrid_ValorIPIVista_Col As Integer
Dim iGrid_ValorIPIPrazo_Col As Integer
Dim iGrid_AliquotaICMS_Col As Integer
Dim iGrid_Observacao_Col As Integer

'Eventos dos Browses
Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1

Private Sub BotaoFechar_Click()
    
    Unload Me
    
End Sub

Private Sub BotaoImprimir_Click()

Dim lErro As Long
Dim objPedidoCotacao As New ClassPedidoCotacao
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_BotaoImprimir_Click

    Set gobjPedidoCotacao = New ClassPedidoCotacao

    'Verifica se os campos obrigatórios foram preenchidos
    If Len(Trim(Codigo.Caption)) = 0 Then gError 76043
    
    gobjPedidoCotacao.lCodigo = StrParaLong(Codigo.Caption)
    gobjPedidoCotacao.iFilialEmpresa = giFilialEmpresa
    
    'Verifica se o Pedido de Cotacao informado existe
    lErro = CF("PedidoCotacaoTodos_Le", gobjPedidoCotacao)
    If lErro <> SUCESSO And lErro <> 67783 Then gError 76044
    
    'Se o Pedido de Cotacao não existe ==> erro
    If lErro = 67783 Then gError 76045
    
    'Executa o relatório
    lErro = objRelatorio.ExecutarDireto("Pedido de Cotação Consulta", "PEDCOTTO.NumIntDoc = @NPEDCOT", 0, "PEDCOT", "NPEDCOT", gobjPedidoCotacao.lNumIntDoc)
    If lErro <> SUCESSO Then gError 76046
    
    'Preenche a Data de Entrada com a Data Atual
    DataEmissao.Caption = Format(gdtDataHoje, "dd/mm/yy")

    'Atualiza data de emissao no BD para a data atual
    lErro = CF("PedidoCotacao_Atualiza_DataEmissao", gobjPedidoCotacao)
    If lErro <> SUCESSO And lErro <> 56348 Then gError 89859
    
    Exit Sub
    
Erro_BotaoImprimir_Click:

    Select Case gErr
    
        Case 76043
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDCOTACAO_IMPRESSAO", gErr)
            
        Case 76044, 76046, 89859
        
        Case 76045
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOTACAO_NAO_ENCONTRADO", gErr, gobjPedidoCotacao.lCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 164692)

    End Select
    
    Exit Sub

End Sub

Private Sub CodigoLabel_Click()

Dim objPedidoCotacao As New ClassPedidoCotacao
Dim colSelecao As New Collection

    'Coloca no objPedidoCotacao o código do pedido da tela
    gobjPedidoCotacao.lCodigo = StrParaLong(Codigo.Caption)

    'Chama a tela de PedidoCotacaoTodosLista
    Call Chama_Tela("PedidoCotacaoTodosLista", colSelecao, gobjPedidoCotacao, objEventoCodigo)

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
Dim objComprador As New ClassComprador
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Form_Load

    iFrameTipoDestinoAtual = 0
    iFrameAtual = 1
    
    bExibirColReal = True
     
    Set objEventoCodigo = New AdmEvento
    Set gobjPedidoCotacao = New ClassPedidoCotacao

    '#############################
    'Inserido por Wagner
    Call Formata_Controles
    '#############################

    'Inicializa a máscara do produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 67736

    Quantidade.Format = FORMATO_ESTOQUE
    QuantEntrega.Format = FORMATO_ESTOQUE

    Set objGridItens = New AdmGrid

    'Inicializa o grid de itens
    lErro = Inicializa_Grid_Itens(objGridItens)
    If lErro <> SUCESSO Then gError 67737

    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 67736, 67737 'Erros tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164693)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub


Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Libera as variáveis globais
    Set objGridItens = Nothing
    Set objEventoCodigo = Nothing
    Set gobjPedidoCotacao = Nothing

    'Libera o comando de setas
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Function Inicializa_Grid_Itens(objGrid As AdmGrid) As Long
'Inicializa o Grid de Itens

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Itens

    Set objGrid.objForm = Me

    'Títulos das colunas
    objGrid.colColuna.Add (" ")
    objGrid.colColuna.Add ("Produto")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("U.M.")
    objGrid.colColuna.Add ("Quantidade")
    objGrid.colColuna.Add ("Preço A Vista")
    objGrid.colColuna.Add ("Total A Vista")
    'Se a moeda for Diferente de Real => Exibe as Colunas de Comparacao
    If bExibirColReal = True Then
        objGrid.colColuna.Add ("Total A Vista (R$)")
    End If
    objGrid.colColuna.Add ("Preço A Prazo")
    objGrid.colColuna.Add ("Total A Prazo")
    'Se a moeda for Diferente de Real => Exibe as Colunas de Comparacao
    If bExibirColReal = True Then
        objGrid.colColuna.Add ("Total A Prazo (R$)")
    End If
    objGrid.colColuna.Add ("Taxa Financeira")
    objGrid.colColuna.Add ("Prazo Entrega")
    objGrid.colColuna.Add ("Quantidade Entrega")
    objGrid.colColuna.Add ("Alíquota IPI")
    objGrid.colColuna.Add ("Valor IPI a Vista")
    objGrid.colColuna.Add ("Valor IPI a Prazo")
    objGrid.colColuna.Add ("Alíquota ICMS")
    objGrid.colColuna.Add ("Observação")

    'Controles que participam do Grid
    objGrid.colCampo.Add (Produto.Name)
    objGrid.colCampo.Add (DescProduto.Name)
    objGrid.colCampo.Add (UnidadeMed.Name)
    objGrid.colCampo.Add (Quantidade.Name)
    objGrid.colCampo.Add (PrecoVista.Name)
    objGrid.colCampo.Add (TotVista.Name)
    If bExibirColReal = True Then
        objGrid.colCampo.Add (TotalVIstaRS.Name)
    End If
    objGrid.colCampo.Add (PrecoPrazo.Name)
    objGrid.colCampo.Add (TotPrazo.Name)
    If bExibirColReal = True Then
        objGrid.colCampo.Add (TotalPrazoRS.Name)
    End If
    objGrid.colCampo.Add (TaxaFinanceira.Name)
    objGrid.colCampo.Add (PrazoEntrega.Name)
    objGrid.colCampo.Add (QuantEntrega.Name)
    objGrid.colCampo.Add (AliquotaIPI.Name)
    objGrid.colCampo.Add (ValorIPIVista.Name)
    objGrid.colCampo.Add (ValorIPIPrazo.Name)
    objGrid.colCampo.Add (AliquotaICM.Name)
    objGrid.colCampo.Add (Observacao.Name)

    'Grid do GridInterno
    objGrid.objGrid = GridProdutos

    'Colunas da Grid
    iGrid_Produto_Col = 1
    iGrid_DescProduto_Col = 2
    iGrid_UnidadeMed_Col = 3
    iGrid_Quantidade_Col = 4
    iGrid_PrecoVista_Col = 5
    iGrid_TotalVista_Col = 6
    
    If bExibirColReal = False Then
        
        iGrid_PrecoPrazo_Col = 7
        iGrid_TotalPrazo_Col = 8
        iGrid_TaxaFinanceira_Col = 9
        iGrid_PrazoEntrega_Col = 10
        iGrid_QuantEntrega_Col = 11
        iGrid_AliquotaIPI_Col = 12
        iGrid_ValorIPIVista_Col = 13
        iGrid_ValorIPIPrazo_Col = 14
        iGrid_AliquotaICMS_Col = 15
        iGrid_Observacao_Col = 16
        
    Else
          
        iGrid_TotalVista_RS_Col = 7
        iGrid_PrecoPrazo_Col = 8
        iGrid_TotalPrazo_Col = 9
        iGrid_TotalPrazo_RS_Col = 10
        iGrid_TaxaFinanceira_Col = 11
        iGrid_PrazoEntrega_Col = 12
        iGrid_QuantEntrega_Col = 13
        iGrid_AliquotaIPI_Col = 14
        iGrid_ValorIPIVista_Col = 15
        iGrid_ValorIPIPrazo_Col = 16
        iGrid_AliquotaICMS_Col = 17
        iGrid_Observacao_Col = 18
        
    End If

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAX_ITENS_PEDIDO_COTACAO + 1

    'Linhas visíveis do grid
    objGrid.iLinhasVisiveis = 22
    
    'largura total do grid
    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    'É proibído excluir e incluir linhas do grid
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    objGrid.iExecutaRotinaEnable = GRID_NAO_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGrid)

    Inicializa_Grid_Itens = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Itens:

    Inicializa_Grid_Itens = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164694)

    End Select

    Exit Function

End Function

Public Function Trata_Parametros(Optional objPedidoCotacao As ClassPedidoCotacao) As Long

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Trata_Parametros

    'Verifica se alguma nota foi passada por parametro
    If Not (objPedidoCotacao Is Nothing) Then
                
        Set gobjPedidoCotacao = objPedidoCotacao
        
        'Lê o Pedido de Cotação
        lErro = CF("PedidoCotacaoTodos_Le", gobjPedidoCotacao)
        If lErro <> SUCESSO And lErro <> 67783 Then gError 67738
        
        'Se o pedido existe --> exibir seus dados
        If lErro = SUCESSO Then

            lErro = Traz_PedidoCotacao_Tela(gobjPedidoCotacao)
            If lErro <> SUCESSO Then gError 67739

        End If

    'Se não foi passado nenhum Pedido de Cotação como parâmetro
    Else

        iChamaTela = 1

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 67738, 67739 'Erros tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164695)

    End Select

    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim iIndice As Integer
Dim objPedidoCotacao As New ClassPedidoCotacao
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "PedidoCotacaoTodos"

    'Se o código estiver preenchido, carrega o obj com o código da tela
    gobjPedidoCotacao.lCodigo = StrParaLong(Codigo.Caption)
    gobjPedidoCotacao.iFilialEmpresa = giFilialEmpresa
        
    'Se o Fornecedor foi preenchido
    If Len(Trim(Fornecedor.Caption)) > 0 Then
        
        objFornecedor.sNomeReduzido = Fornecedor.Caption
        
        'Lê o Fornecedor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 67784
        
        'Se o Fornecedor não está cadastrado, Erro
        If lErro = 6681 Then gError 67785
        
        gobjPedidoCotacao.lFornecedor = objFornecedor.lCodigo
        gobjPedidoCotacao.iFilial = Codigo_Extrai(Filial.Caption)
        
    End If
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", gobjPedidoCotacao.lCodigo, 0, "Codigo"
    colCampoValor.Add "Fornecedor", gobjPedidoCotacao.lFornecedor, 0, "Fornecedor"
    colCampoValor.Add "Filial", gobjPedidoCotacao.iFilial, 0, "Filial"
    
    'Adiciona filtro
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 67784
        
        Case 67785
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164696)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPedidoCotacao As New ClassPedidoCotacao

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set gobjPedidoCotacao = obj1

    'Chama Traz_PedidoCotacao_Tela
    lErro = Traz_PedidoCotacao_Tela(gobjPedidoCotacao)
    If lErro <> SUCESSO Then gError 67740

    'Fecha o sistema de setas
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 67740 'Erro tratado na rotina chamada.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164697)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip1_Click()
    
    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index

    End If

End Sub

Private Sub TipoDestino_Click(Index As Integer)

    'Se o Tipo de Destino não mudou, sai da rotina
    If Index = iFrameTipoDestinoAtual Then Exit Sub

    'Torna Frame correspondente a Index visivel
    FrameTipo(Index).Visible = True

    'Torna Frame atual invisivel
    FrameTipo(iFrameTipoDestinoAtual).Visible = False

    'Armazena novo valor de iFrameTipoDestinoAtual
    iFrameTipoDestinoAtual = Index

   Exit Sub

End Sub

Function Limpa_Tela_PedidoCotacao()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_PedidoCotacao

    Call Limpa_Tela(Me)

    'Limpa o restante da tela
    Codigo.Caption = ""
    Fornecedor.Caption = ""
    Data.Caption = ""
    DataEmissao.Caption = ""
    TipoFrete.ListIndex = -1
    Filial.Caption = ""
    DataValidade.Caption = ""
    CondicaoPagamento.Caption = ""
    FilialEmpresa.Caption = ""
    TotalVista.Caption = ""
    TotalPrazo.Caption = ""
    IPIValorVista.Caption = ""
    IPIValorPrazo.Caption = ""
    FornecDestino.Caption = ""
    FilialFornec.Caption = ""
    TipoDestino(0).Value = False
    TipoDestino(1).Value = False

    'Limpa o grid
    Call Grid_Limpa(objGridItens)

    Limpa_Tela_PedidoCotacao = SUCESSO

    Exit Function

Erro_Limpa_Tela_PedidoCotacao:

    Limpa_Tela_PedidoCotacao = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164698)

    End Select

    Exit Function

End Function

Function Traz_PedidoCotacao_Tela(objPedidoCotacao As ClassPedidoCotacao) As Long
'Traz o Pedido de Cotação para a tela

Dim lErro As Long
Dim objCotacao As New ClassCotacao
Dim objItemPedCotacao As New ClassItemPedCotacao
Dim iIndice As Integer
Dim objItensCotacao As ClassItemCotacao
Dim dValorFrete As Double
Dim dValorSeguro As Double
Dim dValorDespesa As Double
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objFornecedor As New ClassFornecedor
Dim objFilialEmpresa As New AdmFiliais
Dim objCondicaoPagamento As New ClassCondicaoPagto
Dim objComprador As New ClassComprador

On Error GoTo Erro_Traz_PedidoCotacao_Tela
    
    'Guarda na variável global o pedido que está sendo trazido para a tela ...
    Set gobjPedidoCotacao = objPedidoCotacao
    
    'Limpa a Tela
    Call Limpa_Tela_PedidoCotacao
    
    'Lê o Pedido de Cotação
    lErro = CF("PedidoCotacaoTodos_Le", gobjPedidoCotacao)
    If lErro <> SUCESSO Then gError 67778
    
    'Lê os itens do pedido.
    lErro = CF("ItensPedCotacaoTodos_Le", gobjPedidoCotacao)
    If lErro <> SUCESSO Then gError 67749

    objCotacao.lNumIntDoc = gobjPedidoCotacao.lCotacao

    'Lê a tabela de Cotação baixadas ou não
    lErro = CF("CotacaoTodas_Le", objCotacao)
    If lErro <> SUCESSO And lErro <> 67761 Then gError 67779
    
    'Se a Cotação não está cadastrada, Erro
    If lErro = 67761 Then gError 67762
    
    objComprador.iCodigo = objCotacao.iComprador
    
    'Verifica se o usuário do sistema é um comprador
    lErro = CF("Comprador_Le", objComprador)
    If lErro <> SUCESSO And lErro <> 50064 Then gError 67785
    
    'Se não achou o comprador --> erro.
    If lErro = 50064 Then gError 67786
    
    Comprador.Caption = objComprador.sCodUsuario
    
    'Preenche a tela com os dados de PedidoCotacao e Cotacao
    Codigo.Caption = gobjPedidoCotacao.lCodigo
    
    'Condição de Pagamento
    If gobjPedidoCotacao.iCondPagtoPrazo <> 0 Then

        'Preenche o objCondicaoPagamento com o código.
        objCondicaoPagamento.iCodigo = gobjPedidoCotacao.iCondPagtoPrazo

        'Tenta ler CondicaoPagto com esse código no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagamento)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 67763
        
        'Se não encontrou CondicaoPagto, Erro
        If lErro <> SUCESSO Then gError 67764

        CondicaoPagamento.Caption = objCondicaoPagamento.iCodigo & SEPARADOR & objCondicaoPagamento.sDescReduzida
    
    End If

    
    'Preenche o TipoDestino
    If objCotacao.iTipoDestino <> -1 Then
    
        TipoDestino(objCotacao.iTipoDestino).Value = True
    
        If iFrameTipoDestinoAtual = TIPO_DESTINO_EMPRESA Then
    
            objFilialEmpresa.iCodFilial = objCotacao.iFilialEmpresa
    
            'Lê a FilialEmpresa
            lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
            If lErro <> SUCESSO And lErro <> 27378 Then gError 67765
            If lErro = 27378 Then gError 67766
    
            'Coloca a FilialEmpresa na tela
            FilialEmpresa.Caption = objFilialEmpresa.sNome
    
        ElseIf iFrameTipoDestinoAtual = TIPO_DESTINO_FORNECEDOR Then
    
            objFornecedor.lCodigo = objCotacao.lFornCliDestino
    
            'Lê o Fornecedor
            lErro = CF("Fornecedor_Le", objFornecedor)
            If lErro <> SUCESSO And lErro <> 18272 Then gError 67767
            If lErro = 18272 Then gError 67768
    
            'Coloca o Fornecedor na tela.
            FornecDestino.Caption = objFornecedor.sNomeReduzido
    
            objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
            objFilialFornecedor.iCodFilial = objCotacao.iFilialDestino
    
            'Lê a FilialFornecedor
            lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", objFornecedor.sNomeReduzido, objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 18272 Then gError 67769
            
            'Se nao encontrou
            If lErro = 18272 Then gError 67770
    
            'Coloca a Filial na tela
            FilialFornec.Caption = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome

        End If
        
    End If
    
    'Passa o código do fornecedor de objpedidocotacao para o objfornecedor
    objFornecedor.lCodigo = gobjPedidoCotacao.lFornecedor

    'Lê o Fornecedor
    lErro = CF("Fornecedor_Le", objFornecedor)
    If lErro <> SUCESSO And lErro <> 18272 Then gError 67771
    
    'Se não encontrou, Erro
    If lErro = 18272 Then gError 67772

    'Coloca o NomeReduzido do Fornecedor na tela
    Fornecedor.Caption = objFornecedor.sNomeReduzido

    'Passa o CodFornecedor e o CodFilial para o objfilialfornecedor
    objFilialFornecedor.lCodFornecedor = gobjPedidoCotacao.lFornecedor
    objFilialFornecedor.iCodFilial = gobjPedidoCotacao.iFilial

    'Lê o filialforncedor
    lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", objFornecedor.sNomeReduzido, objFilialFornecedor)
    If lErro <> SUCESSO And lErro <> 18272 Then gError 67773
    
    'Se nao encontrou
    If lErro = 18272 Then gError 67774

    'Coloca a filial na tela
    Filial.Caption = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome

    Contato.Caption = gobjPedidoCotacao.sContato
    
    If gobjPedidoCotacao.dtData <> DATA_NULA Then
        Data.Caption = Format(gobjPedidoCotacao.dtData, "dd/mm/yyyy")
    End If
    
    If gobjPedidoCotacao.dtDataEmissao <> DATA_NULA Then
        DataEmissao.Caption = Format(gobjPedidoCotacao.dtDataEmissao, "dd/mm/yyyy")
    End If
    
    If gobjPedidoCotacao.dtDataValidade <> DATA_NULA Then
        DataValidade.Caption = Format(gobjPedidoCotacao.dtDataValidade, "dd/mm/yyyy")
    End If
    
    If gobjPedidoCotacao.dtDataValidade <> DATA_NULA Then
        DataValidade.Caption = Format(gobjPedidoCotacao.dtDataValidade, "dd/mm/yyyy")
    End If
    
    If gobjPedidoCotacao.dtDataBaixa <> DATA_NULA Then
        DataBaixa.Caption = Format(gobjPedidoCotacao.dtDataBaixa, "dd/mm/yyyy")
    End If
    
    'Preenche a combo TipoFrete
    For iIndice = 0 To TipoFrete.ListCount - 1
        If gobjPedidoCotacao.iTipoFrete = TipoFrete.ItemData(iIndice) Then
            TipoFrete.ListIndex = iIndice
        End If
    Next

    lErro = Carrega_Moeda(gobjPedidoCotacao)
    If lErro <> SUCESSO Then gError 108975
    
    lErro = Preenche_GridItens
    If lErro <> SUCESSO Then gError 108975
    
    'Chama a função Totais_Calcula que calcula os totais do pedido
    lErro = Totais_Calcula()
    If lErro <> SUCESSO Then gError 67776

    Traz_PedidoCotacao_Tela = SUCESSO

    Exit Function

Erro_Traz_PedidoCotacao_Tela:

    Traz_PedidoCotacao_Tela = gErr

    Select Case gErr

        Case 67749, 67763, 67775, 67765, 67767, 67769, 67771, 67773, 67776, 67778, 67779, 67785, 108975

        Case 67762
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COTACAO_NAO_CADASTRADA", gErr, gobjPedidoCotacao.lCotacao)

        Case 67764
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagamento.iCodigo)
        
        Case 67766
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case 67768
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)

        Case 67770, 67774
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objFilialFornecedor.iCodFilial, objFilialFornecedor.lCodFornecedor)

        Case 67772
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)
        
        Case 67786
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPRADOR_NAO_CADASTRADO", gErr, objComprador.iCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164699)

    End Select

    Exit Function

End Function

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objPedidoCotacao As New ClassPedidoCotacao

On Error GoTo Erro_Tela_Preenche

    'Passa o código da coleção de campos valores para o objPedidoCotacao
    gobjPedidoCotacao.lCodigo = colCampoValor.Item("Codigo").vValor
    gobjPedidoCotacao.lFornecedor = colCampoValor.Item("Fornecedor").vValor
    gobjPedidoCotacao.iFilial = colCampoValor.Item("Filial").vValor
    
    'Coloca giFilialEmpresa em objPedidoCotacao.giFilialEmpresa
    gobjPedidoCotacao.iFilialEmpresa = giFilialEmpresa

    'Chama Traz_PedidoCotacao_Tela
    lErro = Traz_PedidoCotacao_Tela(gobjPedidoCotacao)
    If lErro <> SUCESSO Then gError 67743

    iAlterado = 0

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 67741, 67743 'Erros tratados nas rotinas chamadas.

        Case 67742
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOTACAO_NAO_ENCONTRADO", gErr, gobjPedidoCotacao.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164700)

    End Select

    Exit Sub

End Sub

Function Totais_Calcula() As Long
'Função que vai atualizar o Valor IPI total do pedido,
'vai calcular o total a vista e o total a prazo do pedido

Dim lErro As Long
Dim dValorIPIVista As Double
Dim dValorIPIPrazo As Double
Dim dValorIPITotalVista As Double
Dim dValorIPITotalPrazo As Double
Dim dValorTotalVista As Double
Dim dTotalVista As Double
Dim dValorTotalPrazo As Double
Dim dTotalPrazo As Double
Dim dValorFrete As Double
Dim dSeguro As Double
Dim dDescontoVista As Double
Dim dDescontoPrazo As Double
Dim iLinha As Integer

On Error GoTo Erro_Totais_Calcula

    'Soma do IPI, TotalVista e TotalPrazo de todos os itens
    For iLinha = 1 To objGridItens.iLinhasExistentes

        'IPIVista
        dValorIPIVista = StrParaDbl(GridProdutos.TextMatrix(iLinha, iGrid_ValorIPIVista_Col))
        dValorIPITotalVista = dValorIPITotalVista + dValorIPIVista

        'IPIPrazo
        dValorIPIPrazo = StrParaDbl(GridProdutos.TextMatrix(iLinha, iGrid_ValorIPIPrazo_Col))
        dValorIPITotalPrazo = dValorIPITotalPrazo + dValorIPIPrazo

        'TotalVista
        dTotalVista = StrParaDbl(GridProdutos.TextMatrix(iLinha, iGrid_TotalVista_Col))
        dValorTotalVista = dValorTotalVista + dTotalVista

        'TotalPrazo
        dTotalPrazo = StrParaDbl(GridProdutos.TextMatrix(iLinha, iGrid_TotalPrazo_Col))
        dValorTotalPrazo = dValorTotalPrazo + dTotalPrazo

    Next

    'Converte para Double os valores de ValorFrete, Seguro e Desconto
    dValorFrete = StrParaDbl(ValorFrete.Caption)
    dSeguro = StrParaDbl(ValorSeguro.Caption)
    dDescontoVista = StrParaDbl(DescontoVista.Caption)
    dDescontoPrazo = StrParaDbl(DescontoPrazo.Caption)

    'Coloca o IPI na tela
    IPIValorVista.Caption = Format(dValorIPITotalVista, "Standard")
    IPIValorPrazo.Caption = Format(dValorIPITotalPrazo, "Standard")

    'Depois de somar o TotalVista dos itens, somar o resultado
    'com o ValorFrete Seguro, Desconto e IPI.
    dValorTotalVista = dValorTotalVista + dValorFrete + dSeguro + dValorIPITotalVista - dDescontoVista + StrParaDbl(ValorDespesas.Caption)

    'Coloca o TotalVista na tela
    TotalVista.Caption = Format(dValorTotalVista, TotalVIstaRS.Format) 'Alterado por Wagner

    'Depois de somar o TotalPrazo dos itens, somar o resultado
    'com o ValorFrete Seguro, Desconto e IPI.
    dValorTotalPrazo = dValorTotalPrazo + dValorFrete + dSeguro + dValorIPITotalPrazo - dDescontoPrazo + StrParaDbl(ValorDespesas.Caption)

    'Coloca o TotalPrazo na tela
    TotalPrazo.Caption = Format(dValorTotalPrazo, TotalPrazoRS.Format) 'Alterado por Wagner

    Totais_Calcula = SUCESSO

    Exit Function

Erro_Totais_Calcula:

    Totais_Calcula = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164701)

    End Select

    Exit Function

End Function

Function Preenche_GridItens() As Long
'Preenche o Grid com os itens de PedidoCotacao

Dim lErro As Long, dTaxa As Double
Dim iIndice As Integer
Dim dTaxaConversao As Double
Dim objItemPedCotacao As New ClassItemPedCotacao
Dim sProdutoEnxuto As String
Dim objItensCotacao As New ClassItemCotacao
Dim dValorDespesa As Double, dValorFrete As Double, dValorSeguro As Double
Dim dValorDescontoVista As Double, dValorDescontoPrazo As Double
Dim objProduto As New ClassProduto, objObservacao As New ClassObservacao
Dim objCotacaoProduto As New ClassCotacaoProduto
Dim dTaxaFinanceira As Double
Dim dPercentual As Double
Dim dValorTotalPrazo As Double, dValorTotalVista As Double, dPrecoUnitarioVista As Double, dPrecoUnitarioPrazo As Double
Dim iCondPagto As Integer, dValorIPIVista As Double, dValorIPIPrazo As Double
Dim iContPrazo As Integer, iContVista As Integer

On Error GoTo Erro_Preenche_GridItens

    iIndice = 0: iContPrazo = 0: iContVista = 0
    
    'Para cada item da Coleção
    For Each objItemPedCotacao In gobjPedidoCotacao.colItens
        
        iIndice = iIndice + 1

        'Passa o produto para o objProduto
        objProduto.sCodigo = objItemPedCotacao.sProduto

        'Lê o produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 53692
        'Não encontrou
        If lErro = 28030 Then gError 53693

        'Formata o Produto
        lErro = Mascara_RetornaProdutoEnxuto(objItemPedCotacao.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 53653

        'Coloca o produto na tela.
        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True

        'Preenche o Grid
        GridProdutos.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
        GridProdutos.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItemPedCotacao.sUM
        GridProdutos.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemPedCotacao.dQuantidade)
        GridProdutos.TextMatrix(iIndice, iGrid_DescProduto_Col) = objProduto.sDescricao
        
        'Janaina
        lErro = CF("ItemPedCotacao_Le_CotacaoProduto", objItemPedCotacao, objCotacaoProduto)
        'Janaina
        If lErro <> SUCESSO And lErro <> 76250 Then gError 76251
        If lErro = 76250 Then gError 76252
               
        dValorDescontoVista = 0
        dValorDescontoPrazo = 0
        dTaxaFinanceira = 0
        dPrecoUnitarioVista = 0
        dPrecoUnitarioPrazo = 0
        
        If (objItemPedCotacao.colItensCotacao.Count = 0 And objProduto.dIPIAliquota > 0) Then GridProdutos.TextMatrix(iIndice, iGrid_AliquotaIPI_Col) = Format(objProduto.dIPIAliquota, "Percent")
        
        For Each objItensCotacao In objItemPedCotacao.colItensCotacao
            
            'Se a moeda é igual a da tela então preenche
            If objItensCotacao.iMoeda = Codigo_Extrai(Moeda.List(Moeda.ListIndex)) Then
            
                'Preenche a taxa
                If StrParaDbl(TaxaConversao.Text) <> objItensCotacao.dTaxa And objItensCotacao.iMoeda <> MOEDA_REAL Then TaxaConversao.Text = Format(objItensCotacao.dTaxa, FORMATO_TAXA_CONVERSAO_MOEDA)
                
                'Verifica a condição de pagamento para preencher os valores
                'do grid com os valores a vista ou a prazo.
                If objItensCotacao.iCondPagto = COD_A_VISTA Then
                    
                    'Preenche o Grid com os valores a Vista
                    If objItensCotacao.dPrecoUnitario > 0 Then GridProdutos.TextMatrix(iIndice, iGrid_PrecoVista_Col) = Format(objItensCotacao.dPrecoUnitario, gobjCOM.sFormatoPrecoUnitario) ' "STANDARD") 'Alterado por Wagner
                    If objItensCotacao.dValorTotal > 0 Then GridProdutos.TextMatrix(iIndice, iGrid_TotalVista_Col) = Format(objItensCotacao.dValorTotal, TotalVIstaRS.Format) 'Alterado por Wagner
                    If objItensCotacao.dAliquotaIPI > 0 Then
                        dValorIPIVista = (objItensCotacao.dAliquotaIPI * objItensCotacao.dValorTotal)
                        GridProdutos.TextMatrix(iIndice, iGrid_ValorIPIVista_Col) = Format(dValorIPIVista, "Fixed")
                        
                    End If
                    
                    dValorTotalVista = objItensCotacao.dValorTotal + dValorTotalVista
                    dPrecoUnitarioVista = objItensCotacao.dPrecoUnitario
                    
                Else
                    
                    'Preenche o Grid com os valores a Prazo
                    If objItensCotacao.dPrecoUnitario > 0 Then GridProdutos.TextMatrix(iIndice, iGrid_PrecoPrazo_Col) = Format(objItensCotacao.dPrecoUnitario, gobjCOM.sFormatoPrecoUnitario) ' "STANDARD") 'Alterado por Wagner
                    If objItensCotacao.dValorTotal > 0 Then GridProdutos.TextMatrix(iIndice, iGrid_TotalPrazo_Col) = Format(objItensCotacao.dValorTotal, TotalPrazoRS.Format) 'Alterado por Wagner
                    If objItensCotacao.dAliquotaIPI > 0 Then
            
                        dValorIPIPrazo = (objItensCotacao.dAliquotaIPI * objItensCotacao.dValorTotal)
                        GridProdutos.TextMatrix(iIndice, iGrid_ValorIPIPrazo_Col) = Format(dValorIPIPrazo, "Fixed")
                    End If
                        
                    dValorTotalPrazo = objItensCotacao.dValorTotal + dValorTotalPrazo
                    dPrecoUnitarioPrazo = objItensCotacao.dPrecoUnitario
                
                End If
                
                dTaxaConversao = objItensCotacao.dTaxa
                
                'Recolhe os dados necessarios para o cálculo da Taxa financeira
                iCondPagto = Codigo_Extrai(CondicaoPagamento.Caption)
                
                'Preenche o Grid com os itens da objItemPedidoCotacao.colItensCotacao
                If objItensCotacao.iPrazoEntrega > 0 Then GridProdutos.TextMatrix(iIndice, iGrid_PrazoEntrega_Col) = objItensCotacao.iPrazoEntrega
                If objItensCotacao.dQuantEntrega > 0 Then GridProdutos.TextMatrix(iIndice, iGrid_QuantEntrega_Col) = Format(objItensCotacao.dQuantEntrega, "Standard")
                If objItensCotacao.dAliquotaIPI > 0 Then GridProdutos.TextMatrix(iIndice, iGrid_AliquotaIPI_Col) = Format(objItensCotacao.dAliquotaIPI, "Percent")
                If objItensCotacao.dAliquotaICMS > 0 Then GridProdutos.TextMatrix(iIndice, iGrid_AliquotaICMS_Col) = Format(objItensCotacao.dAliquotaICMS, "Percent")
    
                If objItensCotacao.lObservacao <> 0 Then
    
                    objObservacao.lNumInt = objItensCotacao.lObservacao
                    'Lê a tabela de observações.
                    lErro = CF("Observacao_Le", objObservacao)
                    If lErro <> SUCESSO And lErro <> 53827 Then gError 53829
                    If lErro = 53827 Then gError 53830
                Else
                    objObservacao.sObservacao = objItensCotacao.sObservacao
    
                End If
    
                GridProdutos.TextMatrix(iIndice, iGrid_Observacao_Col) = objObservacao.sObservacao
        
            End If
            
        Next
        
        If dPrecoUnitarioVista <> dPrecoUnitarioPrazo Then
            
            If dPrecoUnitarioVista > 0 And dPrecoUnitarioPrazo > 0 And (dPrecoUnitarioPrazo > dPrecoUnitarioVista) Then
                'Calcula a Taxa Financeira para o ItemCotacao
                lErro = Calcula_TaxaFinanceira(dPrecoUnitarioPrazo, dPrecoUnitarioVista, dTaxaFinanceira, iCondPagto)
                If lErro <> SUCESSO Then gError 76431
            End If
        End If
        
        If dTaxaConversao > 0 Then Call ComparativoMoedaReal_Calcula(dTaxaConversao, iIndice)
        GridProdutos.TextMatrix(iIndice, iGrid_TaxaFinanceira_Col) = Format(dTaxaFinanceira, "Percent")
        
   Next
    
    For Each objItemPedCotacao In gobjPedidoCotacao.colItens
        For Each objItensCotacao In objItemPedCotacao.colItensCotacao
            If objItensCotacao.iMoeda = Codigo_Extrai(Moeda.Text) Then 'Exit For
                If iContPrazo > 0 And iContVista > 0 Then Exit For
                If objItensCotacao.dPrecoUnitario > 0 Then
                    If objItensCotacao.iCondPagto = COD_A_VISTA Then
                        iContVista = 1
                        dPercentual = objItensCotacao.dValorTotal / dValorTotalVista
                        dValorDescontoVista = objItensCotacao.dValorDesconto / dPercentual
                        dValorFrete = objItensCotacao.dValorFrete / dPercentual
                        dValorSeguro = objItensCotacao.dValorSeguro / dPercentual
                        dValorDespesa = objItensCotacao.dOutrasDespesas / dPercentual
                    Else
                        iContPrazo = 1
                        dPercentual = objItensCotacao.dValorTotal / dValorTotalPrazo
                        dValorDescontoPrazo = objItensCotacao.dValorDesconto / dPercentual
                        dValorFrete = objItensCotacao.dValorFrete / dPercentual
                        dValorSeguro = objItensCotacao.dValorSeguro / dPercentual
                        dValorDespesa = objItensCotacao.dOutrasDespesas / dPercentual
                    End If
                End If
            End If
        Next
        If iContPrazo > 0 And iContVista > 0 Then Exit For
    Next
    
    'Coloca os valores de frete seguro e despesas na tela
    ValorFrete.Caption = Format(dValorFrete, "Standard")
    ValorSeguro.Caption = Format(dValorSeguro, "Standard")
    ValorDespesas.Caption = Format(dValorDespesa, "Standard")
    DescontoVista.Caption = Format(dValorDescontoVista, "Standard")
    DescontoPrazo.Caption = Format(dValorDescontoPrazo, "Standard")

    'Atualiza o número de linhas existentes
    objGridItens.iLinhasExistentes = iIndice

    Call Grid_Refresh_Checkbox(objGridItens)
    
    Preenche_GridItens = SUCESSO

    Exit Function

Erro_Preenche_GridItens:

    Preenche_GridItens = gErr

    Select Case gErr

        Case 53653, 53692, 53829, 76251, 76431 'Erros tratados nas rotinas chamadas.

        Case 53830
            Call Rotina_Erro(vbOKOnly, "ERRO_OBSERVACAO_NAO_CADASTRADA", gErr, objItensCotacao.lObservacao)

        Case 53693
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 76252
            Call Rotina_Erro(vbOKOnly, "ERRO_COTACAOPRODUTO_NAO_ENCONTRADA", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164702)

    End Select

    Exit Function


End Function


'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Pedido de Cotação Consulta"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "PedidoCotacaoCons"
    
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
Private Sub Label6_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label6(Index), Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6(Index), Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label2(Index), Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2(Index), Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label4(Index), Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4(Index), Button, Shift, X, Y)
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

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Data_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Data, Source, X, Y)
End Sub

Private Sub Data_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Data, Button, Shift, X, Y)
End Sub

Private Sub DataEmissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataEmissao, Source, X, Y)
End Sub

Private Sub DataEmissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataEmissao, Button, Shift, X, Y)
End Sub

Private Sub Label18_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label18, Source, X, Y)
End Sub

Private Sub Label18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label18, Button, Shift, X, Y)
End Sub

Private Sub Label21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label21, Source, X, Y)
End Sub

Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label21, Button, Shift, X, Y)
End Sub

Private Sub CondicaoPagtoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CondicaoPagtoLabel, Source, X, Y)
End Sub

Private Sub CondicaoPagtoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CondicaoPagtoLabel, Button, Shift, X, Y)
End Sub

Private Sub Codigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Codigo, Source, X, Y)
End Sub

Private Sub Codigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Codigo, Button, Shift, X, Y)
End Sub

Private Sub LabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTotais, Source, X, Y)
End Sub

Private Sub LabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTotais, Button, Shift, X, Y)
End Sub

Private Sub Label19_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label19, Source, X, Y)
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label19, Button, Shift, X, Y)
End Sub

Private Sub TotalPrazo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalPrazo, Source, X, Y)
End Sub

Private Sub TotalPrazo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalPrazo, Button, Shift, X, Y)
End Sub

Private Sub Label20_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label20, Source, X, Y)
End Sub

Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label20, Button, Shift, X, Y)
End Sub

Private Sub IPIValorVista_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(IPIValorVista, Source, X, Y)
End Sub

Private Sub IPIValorVista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(IPIValorVista, Button, Shift, X, Y)
End Sub

Private Sub TotalVista_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalVista, Source, X, Y)
End Sub

Private Sub TotalVista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalVista, Button, Shift, X, Y)
End Sub

Private Sub Label25_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label25, Source, X, Y)
End Sub

Private Sub Label25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label25, Button, Shift, X, Y)
End Sub

Private Sub IPIValorPrazo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(IPIValorPrazo, Source, X, Y)
End Sub

Private Sub IPIValorPrazo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(IPIValorPrazo, Button, Shift, X, Y)
End Sub

Private Sub FilialFornec_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialFornec, Source, X, Y)
End Sub

Private Sub FilialFornec_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialFornec, Button, Shift, X, Y)
End Sub

Private Sub Label32_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label32, Source, X, Y)
End Sub

Private Sub Label32_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label32, Button, Shift, X, Y)
End Sub

Private Sub FornecDestino_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecDestino, Source, X, Y)
End Sub

Private Sub FornecDestino_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecDestino, Button, Shift, X, Y)
End Sub

Private Sub FilialEmpresa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialEmpresa, Source, X, Y)
End Sub

Private Sub FilialEmpresa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialEmpresa, Button, Shift, X, Y)
End Sub

Private Sub Label37_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label37, Source, X, Y)
End Sub

Private Sub Label37_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label37, Button, Shift, X, Y)
End Sub

Private Sub Label24_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label24, Source, X, Y)
End Sub

Private Sub Label24_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label24, Button, Shift, X, Y)
End Sub

Private Sub Comprador_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Comprador, Source, X, Y)
End Sub

Private Sub Comprador_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Comprador, Button, Shift, X, Y)
End Sub


'Tratamento do GridItens
Private Sub GridProdutos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridProdutos_EnterCell()

    Call Grid_Entrada_Celula(objGridItens, iAlterado)

End Sub

Private Sub GridProdutos_GotFocus()

    Call Grid_Recebe_Foco(objGridItens)

End Sub

Private Sub GridProdutos_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridItens)

End Sub

Private Sub GridProdutos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridProdutos_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridItens)

End Sub

Private Sub GridProdutos_RowColChange()

    Call Grid_RowColChange(objGridItens)

End Sub

Private Sub GridProdutos_Scroll()

    Call Grid_Scroll(objGridItens)

End Sub


Private Sub FornecedorLabel_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(FornecedorLabel(Index), Source, X, Y)
End Sub

Private Sub FornecedorLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel(Index), Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label15(Index), Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15(Index), Button, Shift, X, Y)
End Sub



Private Sub CondicaoPagamento_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CondicaoPagamento, Source, X, Y)
End Sub

Private Sub CondicaoPagamento_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CondicaoPagamento, Button, Shift, X, Y)
End Sub

Private Sub Contato_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Contato, Source, X, Y)
End Sub

Private Sub Contato_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Contato, Button, Shift, X, Y)
End Sub

Private Sub DataValidade_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataValidade, Source, X, Y)
End Sub

Private Sub DataValidade_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataValidade, Button, Shift, X, Y)
End Sub

Private Sub DescontoPrazo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescontoPrazo, Source, X, Y)
End Sub

Private Sub DescontoPrazo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescontoPrazo, Button, Shift, X, Y)
End Sub

Private Sub DescontoVista_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescontoVista, Source, X, Y)
End Sub

Private Sub DescontoVista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescontoVista, Button, Shift, X, Y)
End Sub

Private Sub ValorDespesas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorDespesas, Source, X, Y)
End Sub

Private Sub ValorDespesas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorDespesas, Button, Shift, X, Y)
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

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Function Calcula_PrecoPrazo_TaxaFinanceira(objCondPagto As ClassCondicaoPagto, dPrecoVista As Double, dTaxaFinanceira As Double, dPrecoPrazo As Double) As Long
'Calcula o Preco a Prazo a partir da taxa financeira informada

Dim lErro As Long
Dim iIntervalo As Integer
Dim dTotalDias As Double
Dim iIndice As Integer

On Error GoTo Erro_Calcula_PrecoPrazo_TaxaFinanceira

    dPrecoPrazo = 0

    'Verifica se o intervalo entre parcelas não é  mensal
    If objCondPagto.iMensal = 1 Then
        iIntervalo = 30
        
    Else
        iIntervalo = objCondPagto.iIntervaloParcelas
    End If
    
    If iIntervalo = 0 And objCondPagto.iDiasParaPrimeiraParcela = 0 Then
        dPrecoPrazo = dPrecoVista
    Else

        'Faz o cálculo do Total a Prazo a partir da Taxa Financeira
        For iIndice = 0 To objCondPagto.iNumeroParcelas - 1
    
            dPrecoPrazo = dPrecoPrazo + (dPrecoVista / objCondPagto.iNumeroParcelas) * ((1 + dTaxaFinanceira) ^ ((objCondPagto.iDiasParaPrimeiraParcela + iIndice * iIntervalo) / 30))
    
        Next

    End If

    Calcula_PrecoPrazo_TaxaFinanceira = SUCESSO

    Exit Function

Erro_Calcula_PrecoPrazo_TaxaFinanceira:

    Calcula_PrecoPrazo_TaxaFinanceira = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164703)

    End Select

    Exit Function

End Function

Function Carrega_Moeda(ByVal objPedidoCotacao As ClassPedidoCotacao) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim objItemPedCotacao As New ClassItemPedCotacao
Dim objItensCotacao As New ClassItemCotacao
Dim bAchou As Boolean
Dim ColMoedasUsadas As New Collection
Dim colMoedas As New Collection

On Error GoTo Erro_Carrega_Moeda
    
    Moeda.Clear
    
    lErro = CF("Moedas_Le_Todas", colMoedas)
    If lErro <> SUCESSO Then gError 108976
    
    'se não existem moedas cadastradas
    If colMoedas.Count = 0 Then gError 108977
    
    'Para cada produto ...
    For Each objItemPedCotacao In objPedidoCotacao.colItens
        
        iIndice = iIndice + 1
            
            'Verifica as n cotacoes atribuidas ...
            For Each objItensCotacao In objItemPedCotacao.colItensCotacao
            
                bAchou = False
                
                'E caso ainda nao esteja na colecao de moedas usadas => Adiciona
                If ColMoedasUsadas.Count > 0 Then
                
                    'Para cada item já existente na colecao, verifica a existencia
                    For iIndice2 = 1 To ColMoedasUsadas.Count
                        If Codigo_Extrai(ColMoedasUsadas.Item(iIndice2)) = objItensCotacao.iMoeda Then
                            bAchou = True
                            Exit For
                        End If
                    Next
                    
                End If
                    
                If bAchou = False Then
                    
                    iIndice2 = 1
                
                    Do While colMoedas.Item(iIndice2).iCodigo <> objItensCotacao.iMoeda
                        iIndice2 = iIndice2 + 1
                    Loop
                    
                    If colMoedas.Item(iIndice2).iCodigo = objItensCotacao.iMoeda Then
                        ColMoedasUsadas.Add colMoedas.Item(iIndice2).iCodigo & SEPARADOR & colMoedas.Item(iIndice2).sNome
                    End If
                    
                End If
                
            Next
            
    Next
    
    If ColMoedasUsadas.Count > 0 Then
        
        For iIndice = 1 To ColMoedasUsadas.Count
            Moeda.AddItem ColMoedasUsadas.Item(iIndice)
        Next
    
        If ColMoedasUsadas.Count = 1 Then Moeda.ListIndex = 0
        
    End If
    
    Carrega_Moeda = SUCESSO
    
    Exit Function
    
Erro_Carrega_Moeda:

    Carrega_Moeda = gErr
    
    Select Case gErr
    
        Case 108976
        
        Case 108977
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDAS_NAO_CADASTRADAS", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164704)
    
    End Select

End Function

Private Sub Moeda_Click()

Dim lErro As Long

On Error GoTo Erro_Moeda_Click

    'Limpa a cotacao
    TaxaConversao.Text = ""
    
    'Se a moeda selecionada for = REAL
    If Codigo_Extrai(Moeda.List(Moeda.ListIndex)) = MOEDA_REAL Then
        bExibirColReal = False
    Else
        bExibirColReal = True
    End If
    
    Call Grid_Limpa(objGridItens)
    
    Set objGridItens = New AdmGrid
    
    'Inicializa o grid de itens
    lErro = Inicializa_Grid_Itens(objGridItens)
    If lErro <> SUCESSO Then gError 53622
    
    If Moeda.ListIndex >= 0 Then
    
        'Preenche o grid de Itens
        lErro = Preenche_GridItens()
        If lErro <> SUCESSO Then gError 108951
        
    End If
    
    Call Totais_Calcula
    
    Exit Sub
    
Erro_Moeda_Click:

    Select Case gErr
    
        Case 108951, 108960
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164705)
            
    End Select
    
End Sub

Private Sub ComparativoMoedaReal_Calcula(ByVal dTaxa As Double, ByVal iLinha As Integer)
'Preenche as colunas INFORMATIVAS de proporção da moeda R$.

Dim iIndice As Integer

On Error GoTo Erro_ComparativoMoedaReal_Calcula

    'Preço A vista em R$
    GridProdutos.TextMatrix(iLinha, iGrid_TotalVista_RS_Col) = Format(StrParaDbl(GridProdutos.TextMatrix(iLinha, iGrid_TotalVista_Col)) * dTaxa, TotalVIstaRS.Format) 'Alterado por Wagner
    
    'Preço A prazo em R$
    GridProdutos.TextMatrix(iLinha, iGrid_TotalPrazo_RS_Col) = Format(StrParaDbl(GridProdutos.TextMatrix(iLinha, iGrid_TotalPrazo_Col)) * dTaxa, TotalPrazoRS.Format) 'Alterado por Wagner

    Exit Sub
    
Erro_ComparativoMoedaReal_Calcula:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164706)

    End Select

End Sub

Function Calcula_TaxaFinanceira(dPrecoPrazo As Double, dPrecoVista As Double, dTaxaFinanceira As Double, iCondPagto As Integer) As Long
'Faz o cálculo da taxa financeira, a partir do PrecoPrazo, PrecoVista e CondicaoPagto

Dim lErro As Long
Dim dTaxaIni As Double
Dim dTaxaFim As Double
Dim dTaxaMeio As Double
Dim dValorPresente1 As Double
Dim dValorPresente2 As Double
Dim dValorPresente3 As Double
Dim dValor As Double
Dim bAchou As Boolean
Dim dValorPrazoIni As Double
Dim dValorPrazoFim As Double
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_Calcula_TaxaFinanceira

    objCondicaoPagto.iCodigo = iCondPagto

    'Lê a condicao de Pagamento
    lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
    If lErro <> SUCESSO Then gError 83303

    'se for uma condição a vista
    If objCondicaoPagto.iMensal = 0 And objCondicaoPagto.iIntervaloParcelas = 0 And objCondicaoPagto.iDiasParaPrimeiraParcela = 0 Then
        dTaxaFinanceira = 0
    Else

        'Define os intervalos onde preco vista será tolerado (para menos e para mais)
        dValorPrazoIni = dPrecoPrazo - 0.001
        dValorPrazoFim = dPrecoPrazo + 0.001
    
        'Define o intervalo da taxa financeira
        dTaxaIni = 0
        dTaxaFim = 1
    
        bAchou = False
    
        'Enquanto o Preco Vista nao estiver no intervalo definido
        Do While Not bAchou
    
            lErro = Calcula_PrecoPrazo_TaxaFinanceira(objCondicaoPagto, dPrecoVista, dTaxaFim, dValorPresente1)
            If lErro <> SUCESSO Then gError 76211
    
            lErro = Calcula_PrecoPrazo_TaxaFinanceira(objCondicaoPagto, dPrecoVista, dTaxaIni, dValorPresente2)
            If lErro <> SUCESSO Then gError 76210
    
            If dPrecoPrazo = dValorPresente1 Then
                dTaxaFinanceira = dTaxaFim
                Exit Do
            ElseIf dPrecoPrazo = dValorPresente2 Then
                dTaxaFinanceira = dTaxaIni
                Exit Do
            End If
    
            'Define a taxa intermediaria
            dTaxaMeio = ((dTaxaIni + dTaxaFim) / 2)
    
            If (dValorPresente1 > dPrecoPrazo) And (dPrecoPrazo > dValorPresente2) Then
    
                lErro = Calcula_PrecoPrazo_TaxaFinanceira(objCondicaoPagto, dPrecoVista, dTaxaMeio, dValorPresente3)
                If lErro <> SUCESSO Then gError 76212
    
                If (dValorPrazoFim >= dValorPresente3) And (dValorPrazoIni <= dValorPresente3) Then
                    bAchou = True
                    dTaxaFinanceira = dTaxaMeio
    
                ElseIf (dValorPrazoFim < dValorPresente3) Then
    
                    'altera a taxa inicial
                    dTaxaFim = dTaxaMeio
    
                Else
    
                    dTaxaIni = dTaxaMeio
    
                End If
    
            Else
    
                dTaxaIni = dTaxaFim
                dTaxaFim = dTaxaFim + 1
    
            End If

        Loop

    End If

    Calcula_TaxaFinanceira = SUCESSO

    Exit Function

Erro_Calcula_TaxaFinanceira:

    Calcula_TaxaFinanceira = gErr

    Select Case gErr

        Case 76210 To 76214, 83303
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164707)

    End Select

    Exit Function

End Function

'##############################################
'Inserido por Wagner
Private Sub Formata_Controles()

    PrecoPrazo.Format = gobjCOM.sFormatoPrecoUnitario
    PrecoVista.Format = gobjCOM.sFormatoPrecoUnitario

End Sub
'##############################################


