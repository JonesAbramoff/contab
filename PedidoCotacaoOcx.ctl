VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl PedidoCotacaoOcx 
   ClientHeight    =   9195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   KeyPreview      =   -1  'True
   ScaleHeight     =   9195
   ScaleMode       =   0  'User
   ScaleWidth      =   16995
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame8"
      Height          =   7920
      Index           =   3
      Left            =   90
      TabIndex        =   97
      Top             =   1050
      Visible         =   0   'False
      Width           =   16695
      Begin MSMask.MaskEdBox FilialReq 
         Height          =   225
         Left            =   1785
         TabIndex        =   98
         Top             =   720
         Visible         =   0   'False
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CodigoReq 
         Height          =   225
         Left            =   3465
         TabIndex        =   99
         Top             =   810
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CodPV 
         Height          =   225
         Left            =   4785
         TabIndex        =   101
         Top             =   765
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridRequisicoes 
         Height          =   7635
         Left            =   285
         TabIndex        =   100
         Top             =   180
         Width           =   7110
         _ExtentX        =   12541
         _ExtentY        =   13467
         _Version        =   393216
         Rows            =   16
         Cols            =   8
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   7920
      Index           =   2
      Left            =   150
      TabIndex        =   55
      Top             =   1065
      Visible         =   0   'False
      Width           =   16530
      Begin VB.Frame Frame7 
         Caption         =   "Valores"
         Height          =   7890
         Left            =   0
         TabIndex        =   56
         Top             =   -15
         Width           =   16485
         Begin VB.CommandButton BotaoSalvarGrid 
            Caption         =   "Salvar"
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
            Left            =   6696
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Salva a cotação para a moeda selecionada."
            Top             =   231
            Width           =   888
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
            Left            =   6120
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Traz a última cotação para a moeda selecionada."
            Top             =   231
            Width           =   345
         End
         Begin VB.CommandButton BotaoLimparGrid 
            Caption         =   "Limpar"
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
            Left            =   7644
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Retira a cotação para a moeda selecionada."
            Top             =   231
            Width           =   888
         End
         Begin VB.ComboBox Moeda 
            Height          =   315
            Left            =   912
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   252
            Width           =   2475
         End
         Begin VB.Frame Frame9 
            Caption         =   "Valores"
            Height          =   1245
            Index           =   1
            Left            =   120
            TabIndex        =   65
            Top             =   6570
            Width           =   8445
            Begin MSMask.MaskEdBox ValorFrete 
               Height          =   285
               Left            =   105
               TabIndex        =   20
               Top             =   600
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox DescontoVista 
               Height          =   288
               Left            =   5448
               TabIndex        =   23
               Top             =   336
               Width           =   1512
               _ExtentX        =   2672
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ValorDespesas 
               Height          =   285
               Left            =   2730
               TabIndex        =   22
               Top             =   600
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ValorSeguro 
               Height          =   285
               Left            =   1410
               TabIndex        =   21
               Top             =   600
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox DescontoPrazo 
               Height          =   288
               Left            =   5448
               TabIndex        =   24
               Top             =   840
               Width           =   1512
               _ExtentX        =   2672
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
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
               Height          =   192
               Index           =   1
               Left            =   5424
               TabIndex        =   94
               Top             =   648
               Width           =   1512
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
               Height          =   192
               Index           =   2
               Left            =   4260
               TabIndex        =   93
               Top             =   648
               Width           =   960
            End
            Begin VB.Label IPIValorPrazo 
               BorderStyle     =   1  'Fixed Single
               Height          =   288
               Left            =   4128
               TabIndex        =   92
               Top             =   840
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
               Height          =   192
               Left            =   7092
               TabIndex        =   91
               Top             =   132
               Width           =   1248
            End
            Begin VB.Label TotalVista 
               BorderStyle     =   1  'Fixed Single
               Height          =   288
               Left            =   7068
               TabIndex        =   90
               Top             =   336
               Width           =   1200
            End
            Begin VB.Label IPIValorVista 
               BorderStyle     =   1  'Fixed Single
               Height          =   288
               Left            =   4128
               TabIndex        =   89
               Top             =   336
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
               Height          =   192
               Index           =   1
               Left            =   4260
               TabIndex        =   88
               Top             =   132
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
               Left            =   1725
               TabIndex        =   87
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
               Height          =   192
               Index           =   0
               Left            =   5424
               TabIndex        =   86
               Top             =   132
               Width           =   1476
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
               TabIndex        =   85
               Top             =   375
               Width           =   450
            End
            Begin VB.Label TotalPrazo 
               BorderStyle     =   1  'Fixed Single
               Height          =   288
               Left            =   7068
               TabIndex        =   84
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
               Left            =   2940
               TabIndex        =   83
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
               Height          =   192
               Left            =   7092
               TabIndex        =   82
               Top             =   648
               Width           =   1248
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Itens"
            Height          =   5955
            Left            =   120
            TabIndex        =   60
            Top             =   615
            Width           =   16230
            Begin MSMask.MaskEdBox TotalPrazoRS 
               Height          =   228
               Left            =   6336
               TabIndex        =   95
               Top             =   2160
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
               Left            =   4536
               TabIndex        =   96
               Top             =   2160
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
               TabIndex        =   63
               Top             =   1050
               Width           =   4000
            End
            Begin VB.CheckBox Exclusivo 
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
               Left            =   360
               TabIndex        =   69
               Top             =   510
               Width           =   1530
            End
            Begin VB.TextBox Observacao 
               BorderStyle     =   0  'None
               Height          =   225
               Left            =   5670
               MaxLength       =   255
               TabIndex        =   67
               Top             =   1695
               Width           =   2625
            End
            Begin MSMask.MaskEdBox Produto 
               Height          =   225
               Left            =   495
               TabIndex        =   62
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
               TabIndex        =   70
               Top             =   1410
               Width           =   1380
               _ExtentX        =   2434
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
            Begin MSMask.MaskEdBox QuantEntrega 
               Height          =   210
               Left            =   2310
               TabIndex        =   71
               Top             =   1530
               Width           =   1560
               _ExtentX        =   2752
               _ExtentY        =   370
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
            Begin MSMask.MaskEdBox ValorIPIVista 
               Height          =   225
               Left            =   1935
               TabIndex        =   72
               Top             =   840
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
               Left            =   2310
               TabIndex        =   73
               Top             =   2010
               Width           =   1410
               _ExtentX        =   2487
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
            Begin MSMask.MaskEdBox PrazoEntrega 
               Height          =   225
               Left            =   3675
               TabIndex        =   74
               Top             =   810
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   397
               _Version        =   393216
               BorderStyle     =   0
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
            Begin MSMask.MaskEdBox TotVista 
               Height          =   210
               Left            =   2280
               TabIndex        =   75
               Top             =   1770
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   370
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
            Begin MSMask.MaskEdBox TotPrazo 
               Height          =   225
               Left            =   375
               TabIndex        =   76
               Top             =   2010
               Width           =   1380
               _ExtentX        =   2434
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
            Begin MSMask.MaskEdBox PrecoVista 
               Height          =   210
               Left            =   390
               TabIndex        =   77
               Top             =   1770
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   370
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
               Format          =   "#,##0.00###"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox UnidadeMed 
               Height          =   225
               Left            =   630
               TabIndex        =   68
               Top             =   1305
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
               TabIndex        =   64
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
               Left            =   4380
               TabIndex        =   78
               Top             =   1170
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
               Left            =   6420
               TabIndex        =   79
               Top             =   1215
               Width           =   1320
               _ExtentX        =   2328
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
            Begin MSMask.MaskEdBox PrecoPrazo 
               Height          =   225
               Left            =   7335
               TabIndex        =   80
               Top             =   885
               Width           =   1260
               _ExtentX        =   2223
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
               Format          =   "#,##0.00###"
               PromptChar      =   " "
            End
            Begin MSFlexGridLib.MSFlexGrid GridProdutos 
               Height          =   5430
               Left            =   75
               TabIndex        =   61
               Top             =   360
               Width           =   16050
               _ExtentX        =   28310
               _ExtentY        =   9578
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
               TabIndex        =   81
               Top             =   4350
               Width           =   900
            End
         End
         Begin MSMask.MaskEdBox TaxaConversao 
            Height          =   312
            Left            =   5172
            TabIndex        =   17
            Top             =   240
            Width           =   924
            _ExtentX        =   1614
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "###,##0.00##"
            PromptChar      =   " "
         End
         Begin VB.Label LabelTaxa 
            AutoSize        =   -1  'True
            Caption         =   "Taxa Conversão:"
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
            Left            =   3708
            TabIndex        =   58
            Top             =   300
            Width           =   1452
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
            Left            =   252
            TabIndex        =   57
            Top             =   300
            Width           =   648
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   7950
      Index           =   1
      Left            =   120
      TabIndex        =   26
      Top             =   1065
      Width           =   16650
      Begin VB.CommandButton BotaoPedidosAtualizar 
         Caption         =   "Pedidos de Cotação a Atualizar"
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
         Left            =   120
         TabIndex        =   7
         Top             =   6390
         Width           =   3030
      End
      Begin VB.CommandButton BotaoPedidosAtualizados 
         Caption         =   "Pedidos de Cotação Atualizados"
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
         Left            =   3360
         TabIndex        =   8
         Top             =   6390
         Width           =   3105
      End
      Begin VB.Frame Frame6 
         Caption         =   "Datas"
         Height          =   960
         Left            =   135
         TabIndex        =   48
         Top             =   4890
         Width           =   10020
         Begin MSComCtl2.UpDown UpDownDataValidade 
            Height          =   300
            Left            =   7890
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   345
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataValidade 
            Height          =   300
            Left            =   6855
            TabIndex        =   6
            Top             =   330
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label DataEmissao 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   3855
            TabIndex        =   52
            Top             =   330
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
            Left            =   2580
            TabIndex        =   51
            Top             =   375
            Width           =   1230
         End
         Begin VB.Label Data 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1140
            TabIndex        =   50
            Top             =   360
            Width           =   1110
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
            Left            =   615
            TabIndex        =   49
            Top             =   375
            Width           =   480
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
            Left            =   5535
            TabIndex        =   53
            Top             =   375
            Width           =   1275
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Cabeçalho"
         Height          =   4110
         Left            =   135
         TabIndex        =   27
         Top             =   345
         Width           =   10035
         Begin VB.ComboBox CondicaoPagamento 
            Height          =   315
            ItemData        =   "PedidoCotacaoOcx.ctx":0000
            Left            =   6405
            List            =   "PedidoCotacaoOcx.ctx":0002
            TabIndex        =   2
            Top             =   1290
            Width           =   2955
         End
         Begin VB.Frame Frame2 
            Caption         =   "Local de Entrega"
            Height          =   1185
            Left            =   105
            TabIndex        =   38
            Top             =   2490
            Width           =   9540
            Begin VB.Frame FrameTipo 
               BorderStyle     =   0  'None
               Height          =   645
               Index           =   1
               Left            =   4785
               TabIndex        =   43
               Top             =   300
               Visible         =   0   'False
               Width           =   4650
               Begin VB.Label FilialFornec 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   300
                  Left            =   1455
                  TabIndex        =   47
                  Top             =   330
                  Width           =   3045
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
                  Left            =   885
                  TabIndex        =   46
                  Top             =   375
                  Width           =   465
               End
               Begin VB.Label FornecDestino 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   300
                  Left            =   1470
                  TabIndex        =   45
                  Top             =   15
                  Width           =   3045
               End
               Begin VB.Label FornecDestinoLabel 
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
                  Left            =   315
                  TabIndex        =   44
                  Top             =   60
                  Width           =   1035
               End
            End
            Begin VB.Frame Frame3 
               Caption         =   "Tipo"
               Height          =   585
               Left            =   420
               TabIndex        =   39
               Top             =   345
               Width           =   3645
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
                  TabIndex        =   4
                  Top             =   225
                  Width           =   1515
               End
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
                  TabIndex        =   5
                  Top             =   240
                  Width           =   1335
               End
            End
            Begin VB.Frame FrameTipo 
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   675
               Index           =   0
               Left            =   4650
               TabIndex        =   40
               Top             =   180
               Width           =   4605
               Begin VB.Label FilialEmpresa 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   300
                  Left            =   1575
                  TabIndex        =   42
                  Top             =   345
                  Width           =   3045
               End
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
                  Left            =   1005
                  TabIndex        =   41
                  Top             =   390
                  Width           =   465
               End
            End
         End
         Begin VB.ComboBox TipoFrete 
            Height          =   315
            ItemData        =   "PedidoCotacaoOcx.ctx":0004
            Left            =   1425
            List            =   "PedidoCotacaoOcx.ctx":000E
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1800
            Width           =   2715
         End
         Begin MSMask.MaskEdBox Contato 
            Height          =   300
            Left            =   1425
            TabIndex        =   1
            Top             =   1290
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   300
            Left            =   1425
            TabIndex        =   0
            Top             =   270
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin VB.Label Comprador 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6405
            TabIndex        =   30
            Top             =   255
            Width           =   2910
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
            Left            =   5385
            TabIndex        =   29
            Top             =   315
            Width           =   975
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
            Left            =   4590
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   36
            Top             =   1350
            Width           =   1770
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
            Left            =   645
            TabIndex        =   35
            Top             =   1350
            Width           =   735
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
            Left            =   450
            TabIndex        =   37
            Top             =   1845
            Width           =   945
         End
         Begin VB.Label Filial 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6405
            TabIndex        =   34
            Top             =   780
            Width           =   2940
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
            Left            =   5895
            TabIndex        =   33
            Top             =   840
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
            Left            =   345
            TabIndex        =   31
            Top             =   840
            Width           =   1035
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
            TabIndex        =   28
            Top             =   315
            Width           =   930
         End
         Begin VB.Label Fornecedor 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1425
            TabIndex        =   32
            Top             =   780
            Width           =   2730
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   11145
      ScaleHeight     =   495
      ScaleWidth      =   5610
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   30
      Width           =   5670
      Begin VB.CommandButton BotaoEmail 
         Height          =   345
         Left            =   2580
         Picture         =   "PedidoCotacaoOcx.ctx":001C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Enviar email"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoPedCompra 
         Caption         =   "Gera Pedido de Compra"
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
         Left            =   108
         TabIndex        =   9
         ToolTipText     =   "Atualiza o Pedido de Cotação e gera o Pedido de Compras"
         Top             =   75
         Width           =   2385
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   4095
         Picture         =   "PedidoCotacaoOcx.ctx":09BE
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   3090
         Picture         =   "PedidoCotacaoOcx.ctx":0B48
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Imprimir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   5100
         Picture         =   "PedidoCotacaoOcx.ctx":0C4A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   4590
         Picture         =   "PedidoCotacaoOcx.ctx":0DC8
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   3600
         Picture         =   "PedidoCotacaoOcx.ctx":12FA
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   8820
      Left            =   75
      TabIndex        =   25
      Top             =   345
      Width           =   16815
      _ExtentX        =   29660
      _ExtentY        =   15558
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Principais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Itens"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Requisições"
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
Attribute VB_Name = "PedidoCotacaoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim giMoedaAnterior As Integer
Dim gbCarregandoTela As Boolean
Dim gColMoedasUsadas As Collection
Dim bExibirColReal As Boolean
Dim gobjPedidoCotacao As ClassPedidoCotacao
Dim gbPrecoGridAlterado As Boolean
Dim giPodeAumentarQuant As Integer
 
Dim iAlterado As Integer
Dim iChamaTela As Integer

Dim objGridItens As AdmGrid
Dim iFrameTipoDestinoAtual As Integer
Dim giFrameAtual As Integer
Dim iGrid_Exclusivo_Col As Integer
Dim iGrid_Produto_Col As Integer
Dim iGrid_DescProduto_Col As Integer
Dim iGrid_UnidadeMed_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_PrecoVista_Col As Integer
Dim iGrid_TotalVista_Col As Integer
Dim iGrid_PrecoPrazo_Col As Integer
Dim iGrid_TotalPrazo_Col As Integer
Dim iGrid_TaxaFinanceira_Col As Integer
Dim iGrid_PrazoEntrega_Col As Integer
Dim iGrid_QuantEntrega_Col As Integer
Dim iGrid_AliquotaIPI_Col As Integer
Dim iGrid_ValorIPIVista_Col As Integer
Dim iGrid_ValorIPIPrazo_Col As Integer
Dim iGrid_AliquotaICMS_Col As Integer
Dim iGrid_Observacao_Col As Integer
Dim iGrid_TotalVista_RS_Col As Integer
Dim iGrid_TotalPrazo_RS_Col As Integer

'GridPV
Dim objGridRequisicoes As AdmGrid
Dim iGrid_FilialReq_Col As Integer
Dim iGrid_CodigoReq_Col As Integer
Dim iGrid_CodPV_Col As Integer


Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoCondicaoPagto As AdmEvento
Attribute objEventoCondicaoPagto.VB_VarHelpID = -1
Private WithEvents objEventoBotaoPedAtualizar As AdmEvento
Attribute objEventoBotaoPedAtualizar.VB_VarHelpID = -1
Private WithEvents objEventoBotaoPedAtualizados As AdmEvento
Attribute objEventoBotaoPedAtualizados.VB_VarHelpID = -1

Private Sub AliquotaICM_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub AliquotaICM_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub AliquotaICM_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub AliquotaICM_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = AliquotaICM
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub AliquotaIPI_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub AliquotaIPI_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub AliquotaIPI_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub AliquotaIPI_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = AliquotaIPI
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub BotaoEmail_Click()

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim sMailTo As String
Dim objPedidoCotacao As New ClassPedidoCotacao
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objEndereco As New ClassEndereco, sInfoEmail As String

On Error GoTo Erro_BotaoEmail_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se os campos obrigatórios foram preenchidos
    If Len(Trim(Codigo.Text)) = 0 Then gError 129311
    
    objPedidoCotacao.lCodigo = StrParaLong(Codigo.Text)
    objPedidoCotacao.iFilialEmpresa = giFilialEmpresa
    
    'Verifica se o Pedido de Cotacao informado existe
    lErro = CF("PedidoCotacaoTodos_Le", objPedidoCotacao)
    If lErro <> SUCESSO And lErro <> 67783 Then gError 129312
    
    'Se o Pedido de Cotacao não existe ==> erro
    If lErro = 67783 Then gError 129313
    
    If objPedidoCotacao.lFornecedor <> 0 And objPedidoCotacao.iFilial <> 0 Then

        objFilialFornecedor.lCodFornecedor = objPedidoCotacao.lFornecedor
        objFilialFornecedor.iCodFilial = objPedidoCotacao.iFilial

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
    
    'Dispara a impressão do relatório
    lErro = objRelatorio.ExecutarDiretoEmail("Pedido de Cotação", "PEDCOTTO.NumIntDoc = @NPEDCOT", 0, "PEDCOT", "NPEDCOT", objPedidoCotacao.lNumIntDoc, "TTO_EMAIL", sMailTo, "TSUBJECT", "Pedido de Cotação " & CStr(objPedidoCotacao.lCodigo), "TALIASATTACH", "PedCot" & CStr(objPedidoCotacao.lCodigo), "TINFO_EMAIL", sInfoEmail)
    If lErro <> SUCESSO Then gError 129317
    
    'Preenche a Data de Entrada com a Data Atual
    DataEmissao.Caption = Format(gdtDataHoje, "dd/mm/yyyy")

    'Atualiza data de emissao no BD para a data atual
    lErro = CF("PedidoCotacao_Atualiza_DataEmissao", objPedidoCotacao)
    If lErro <> SUCESSO And lErro <> 56348 Then gError 129318

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub
    
Erro_BotaoEmail_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 129317
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDCOTACAO_IMPRESSAO", gErr)
            
        Case 129311, 129312, 129314, 129315, 129318
        
        Case 129313
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOTACAO_NAO_ENCONTRADO", gErr, objPedidoCotacao.lCodigo)
            
        Case 129316
            Call Rotina_Erro(vbOKOnly, "ERRO_EMAIL_NAO_ENCONTRADO", gErr, objPedidoCotacao.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164717)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objPedidoCotacao As New ClassPedidoCotacao
Dim vbMsgRes As VbMsgBoxResult
Dim lCodigo As Long

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o código foi informado
    If Len(Trim(Codigo.Text)) = 0 Then gError 53679

    'Recolhe os dados do pedido da tela
    lErro = Move_PedidoCotacao_Memoria(objPedidoCotacao)
    If lErro <> SUCESSO Then gError 53683

    'Passa o código e a filial empresa para o obj.
    objPedidoCotacao.lCodigo = StrParaLong(Codigo.Text)
    objPedidoCotacao.iFilialEmpresa = giFilialEmpresa

    'Lê o Pedido de Cotação no BD.
    lErro = CF("PedidoCotacao_Le", objPedidoCotacao)
    If lErro <> SUCESSO And lErro <> 62867 Then gError 53680

    'Se não encontrou o Pedido de Cotação --> erro
    If lErro = 62867 Then gError 53682

    'Pede a confirmação da exclusão do Pedido de Venda
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_PEDIDOCOTACAO", objPedidoCotacao.lCodigo)
    If vbMsgRes = vbYes Then

        'Faz a exclusão do Pedido de Cotação
        lErro = CF("PedidoCotacao_Exclui", objPedidoCotacao)
        If lErro <> SUCESSO Then gError 53681
    
        'Limpa a Tela de Pedido de Cotação (menos o comprador)
        Call Limpa_Tela_PedidoCotacao
    
        'Fecha o comando se setas
        lErro = ComandoSeta_Fechar(Me.Name)

        iAlterado = 0
        
        gbPrecoGridAlterado = False
        
    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 53679
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOTACAO_NAO_SELECIONADO", gErr)

        Case 53680, 53681, 53683 'Erros tratados nas rotinas chamadas.

        Case 53682
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOTACAO_NAO_ENCONTRADO", gErr, objPedidoCotacao.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164718)

    End Select
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar

    'Chama a função Gravar_Registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 53619

    'Limpa a tela
    Call Limpa_Tela_PedidoCotacao
    
    Set gobjPedidoCotacao = New ClassPedidoCotacao
    Set gColMoedasUsadas = New Collection

    iAlterado = 0
    
    gbPrecoGridAlterado = False

    Exit Sub

Erro_BotaoGravar:

    Select Case gErr

        Case 53619 'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164719)

    End Select

    Exit Sub

End Sub

Private Sub BotaoImprimir_Click()

Dim lErro As Long
Dim objPedidoCotacao As New ClassPedidoCotacao
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_BotaoImprimir_Click

    'Verifica se os campos obrigatórios foram preenchidos
    If Len(Trim(Codigo.Text)) = 0 Then gError 76039
    
    objPedidoCotacao.lCodigo = StrParaLong(Codigo.Text)
    objPedidoCotacao.iFilialEmpresa = giFilialEmpresa
    
    'Verifica se o Pedido de Cotacao informado existe
    lErro = CF("PedidoCotacaoTodos_Le", objPedidoCotacao)
    If lErro <> SUCESSO And lErro <> 67783 Then gError 76040
    
    'Se o Pedido de Cotacao não existe ==> erro
    If lErro = 67783 Then gError 76041
    
    'Executa o relatório
    lErro = objRelatorio.ExecutarDireto("Pedido de Cotação", "PEDCOTTO.NumIntDoc = @NPEDCOT", 0, "PEDCOT", "NPEDCOT", objPedidoCotacao.lNumIntDoc)
    If lErro <> SUCESSO Then gError 76042
    
    'Preenche a Data de Entrada com a Data Atual
    DataEmissao.Caption = Format(gdtDataHoje, "dd/mm/yyyy")

    'Atualiza data de emissao no BD para a data atual
    lErro = CF("PedidoCotacao_Atualiza_DataEmissao", objPedidoCotacao)
    If lErro <> SUCESSO And lErro <> 56348 Then gError 56096

    Exit Sub
    
Erro_BotaoImprimir_Click:

    Select Case gErr
    
        Case 76039
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDCOTACAO_IMPRESSAO", gErr)
            
        Case 76040, 76042
        
        Case 76041
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOTACAO_NAO_ENCONTRADO", gErr, objPedidoCotacao.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164720)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 53618
    
    Set gobjPedidoCotacao = New ClassPedidoCotacao
    Set gColMoedasUsadas = New Collection
    
    'Limpa a tela.
    Call Limpa_Tela_PedidoCotacao

    'Fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0
    
    gbPrecoGridAlterado = False

    Exit Sub

Erro_BotaoLimpar:

    Select Case gErr

        Case 53618  'Erro tratado na rotina chamada.

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164721)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimparGrid_Click()
'Remove da colecao a moeda selecionada

Dim lErro As Long

On Error GoTo Erro_BotaoLimparGrid_Click

    'Se existe uma moeda selecionada e linhas preenchidas no grid
    If Moeda.ListIndex >= 0 And objGridItens.iLinhasExistentes <> 0 Then
    
        'Remove da colecao
        lErro = RemoveCotacaoGlobal(Codigo_Extrai(Moeda.List(Moeda.ListIndex)), True)
        If lErro <> SUCESSO Then gError 108962
        
        gbPrecoGridAlterado = False
        
        Moeda.ListIndex = -1
        
        Call Grid_Limpa(objGridItens)
        
''''        lErro = Preenche_GridItens()
''''        If lErro <> SUCESSO Then gError 108970
        
        Call Totais_Calcula

    End If

    Exit Sub
    
Erro_BotaoLimparGrid_Click:

    Select Case gErr
    
        Case 108962, 108970
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164722)
            
    End Select

End Sub

Private Sub BotaoPedCompra_Click()

Dim objPedidoCompra As New ClassPedidoCompras
Dim iLinha As Integer
Dim iColuna As Integer
Dim iCondPagto As Integer
Dim lErro As Long
Dim colPedidoCompra As New Collection
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim objItemPedCotacao As New ClassItemPedCotacao
Dim objItensCotacao As New ClassItemCotacao
Dim bAchou As Boolean, dTaxa As Double
Dim colMoedasCotadas As New Collection

On Error GoTo Erro_BotaoPedCompra

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifca se há pedido de cotação na tela
    If (Len(Trim(Codigo.Text)) = 0) Or (gobjPedidoCotacao.colItens.Count = 0) Then gError 62869
    
    'Salva os valores da tela
    If Moeda.ListIndex > -1 Then
        Call SalvarGrid
        Call Moeda_Click
    End If
    
    Call Determina_Moedas_Usadas
    
    Call Determina_Moedas_Cotadas(colMoedasCotadas)

    If colMoedasCotadas.Count = 0 Then gError 106760

    If colMoedasCotadas.Count = 1 Then objPedidoCompra.iMoeda = Codigo_Extrai(colMoedasCotadas.Item(1))
    
    'Chama a tela de Condição de Pagamento para saber qual a Condição a ser utilizada.
    Call Chama_Tela_Modal("CondPagto", objPedidoCompra, colMoedasCotadas)

    If objPedidoCompra.iCondicaoPagto <> 0 Then

        Call Combo_Seleciona(Moeda, objPedidoCompra.iMoeda)
        If objPedidoCompra.iMoeda <> MOEDA_REAL Then dTaxa = StrParaDbl(TaxaConversao.Text)

        'Recolhe os dados da tela
        lErro = Move_PedidoCotacao_Memoria(gobjPedidoCotacao)
        If lErro <> SUCESSO Then gError 53703
            
        'Verifica se o Status do Pedido de Cotacao está atualizado
        If gobjPedidoCotacao.iStatus <> STATUS_ATUALIZADO Then gError 76429
        
        'Função que vai ler a tabela de CotacaoProdutoItemRC para verificar se o pedido de cotação é avulso.
        lErro = CF("CotacaoProdutoItemRC_Le", gobjPedidoCotacao)
        If lErro <> SUCESSO And lErro <> 53788 Then gError 53704
        
        'se a cotação for diferente de sucesso ==> é um pedido de cotação avulso ==> erro
        If lErro <> SUCESSO Then gError 53705
    
        'Se a condição for a prazo e a condicao de pagamento não estiver preenchida --> erro.
        If objPedidoCompra.iCondicaoPagto = CONDPAGTO_PRAZO Then
            objPedidoCompra.iCondicaoPagto = CondPagto_Extrai(CondicaoPagamento)
            If objPedidoCompra.iCondicaoPagto = 0 Then gError 53706
        End If
    
        'Verifica se para a condição de pagamento escolhida todos os
        'itens de cotação estão com o valor preenchido.
        If objPedidoCompra.iCondicaoPagto = CONDPAGTO_VISTA Then
            iColuna = iGrid_TotalVista_Col
        Else
            iColuna = iGrid_TotalPrazo_Col
        End If
    
        For iLinha = 1 To objGridItens.iLinhasExistentes
            If Len(Trim(GridProdutos.TextMatrix(iLinha, iColuna))) = 0 Then gError 53707
        Next
        
        iCondPagto = objPedidoCompra.iCondicaoPagto
        
        'Preenche o objPedidoCompra com os dados da tela.
        lErro = Move_Pedido_Memoria(colPedidoCompra, gobjPedidoCotacao, iCondPagto, objPedidoCompra.iMoeda, dTaxa)
        If lErro <> SUCESSO Then gError 53708
    
        'Atualiza o Pedido de Cotacao
        lErro = CF("PedidoCotacao_Atualiza", gobjPedidoCotacao)
        If lErro <> SUCESSO Then gError 76430
        
        'Gera Pedidos de Compra.
        lErro = CF("PedidoCotacao_Grava_PedidoCompra", colPedidoCompra)
        If lErro <> SUCESSO Then gError 53709
        
        'Limpa a Tela
        Call Limpa_Tela_PedidoCotacao
    
        'Fecha o sistema de setas
        lErro = ComandoSeta_Fechar(Me.Name)
    
        iAlterado = 0
        
        gbPrecoGridAlterado = False
        
        Call Rotina_Aviso(vbOKOnly, "AVISO_PEDIDOCOMPRA_GERADO")
    
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoPedCompra:

    Select Case gErr

        Case 53703, 53704, 53708, 53709, 76430, 62870
            'Erros tratados nas rotinas chamadas.

        Case 53705
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_COTACAO_AVULSO", gErr)

        Case 53706
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_PREENCHIDA", gErr)

        Case 53707
            Call Rotina_Erro(vbOKOnly, "ERRO_PRECOS_ITENS_CONDPAGTO_NAO_PREENCHIDOS", gErr)

        Case 62869
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOTACAO_NAO_SELECIONADO", gErr)
        
        Case 76429
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOTACAO_STATUS_NAOATUALIZADO_PC", gErr)
            
        Case 106760
            Call Rotina_Erro(vbOKOnly, "ERRO_FALTA_ATUALIZAR_PRECO_COTACAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164723)

    End Select
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

End Sub

Private Sub BotaoPedidosAtualizados_Click()

Dim objPedidoCotacao As New ClassPedidoCotacao
Dim colSelecao As New Collection

    'Se o código do pedido estiver preenchido, preenche o
    'objPedidoCotacao.lCodigo com o código da tela
    If Len(Trim(Codigo.Text)) > 0 Then
        objPedidoCotacao.lCodigo = StrParaLong(Codigo.Text)
    End If

    'Chama a tela PedidoCotAtualizadosLista
    Call Chama_Tela("PedidoCotAtualizadosLista", colSelecao, objPedidoCotacao, objEventoBotaoPedAtualizados)

End Sub

Private Sub BotaoPedidosAtualizar_Click()

Dim objPedidoCotacao As New ClassPedidoCotacao
Dim colSelecao As New Collection

    'Se o código do pedido estiver preenchido, preenche o
    'objPedidoCotacao.lCodigo com o código da tela
    If Len(Trim(Codigo.Text)) > 0 Then

        objPedidoCotacao.lCodigo = StrParaLong(Codigo.Text)

    End If

    'Chama a tela PedidoCotAtualizarLista
    Call Chama_Tela("PedidoCotAtualizarLista", colSelecao, objPedidoCotacao, objEventoBotaoPedAtualizar)

End Sub

Private Sub BotaoSalvarGrid_Click()
'Guarda no obj Global o grid

    'Se existe moeda selecionada ...
    If Moeda.ListIndex >= 0 And objGridItens.iLinhasExistentes <> 0 Then
    
        Call SalvarGrid
        
        Moeda.ListIndex = -1
        
        Call Grid_Limpa(objGridItens)
        
    End If

    Exit Sub

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodigoLabel_Click()

Dim objPedidoCotacao As New ClassPedidoCotacao
Dim colSelecao As New Collection

    'Coloca no objPedidoCotacao o código do pedido da tela
    objPedidoCotacao.lCodigo = StrParaLong(Codigo.Text)

    'Chama a tela de PedidoCotacaoLista
    Call Chama_Tela("PedidoCotacaoLista", colSelecao, objPedidoCotacao, objEventoCodigo)

    Exit Sub

End Sub

Private Sub Comprador_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Comprador_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CondicaoPagamento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CondicaoPagamento_Click()

Dim dPrecoPrazo As Double
Dim dPrecoVista As Double
Dim iLinha As Integer
Dim dTaxaFinanceira As Double
Dim lErro As Long

On Error GoTo Erro_CondicaoPagamento_Click

    iAlterado = REGISTRO_ALTERADO
    
    If CondicaoPagamento.ListIndex <> -1 Then

        For iLinha = 1 To objGridItens.iLinhasExistentes
    
            dPrecoPrazo = StrParaDbl(GridProdutos.TextMatrix(iLinha, iGrid_PrecoPrazo_Col))
            dPrecoVista = StrParaDbl(GridProdutos.TextMatrix(iLinha, iGrid_PrecoVista_Col))
            
            If dPrecoPrazo <> dPrecoVista And dPrecoVista > 0 And dPrecoPrazo > 0 Then
                lErro = Calcula_TaxaFinanceira(dPrecoPrazo, dPrecoVista, dTaxaFinanceira, CondPagto_Extrai(CondicaoPagamento))
                If lErro <> SUCESSO Then gError 83804
            End If
        
            'coloca a taxa financeira no grid
            GridProdutos.TextMatrix(iLinha, iGrid_TaxaFinanceira_Col) = Format(dTaxaFinanceira, "Percent")
    
        Next

    End If

    Exit Sub

Erro_CondicaoPagamento_Click:

    Select Case gErr

        Case 83804
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164724)

    End Select

    Exit Sub

End Sub

Private Sub CondicaoPagamento_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCondicaoPagamento As New ClassCondicaoPagto
Dim iCodigo As Integer
Dim iIndice As Integer
Dim dPrecoPrazo As Double
Dim dPrecoVista As Double
Dim iLinha As Integer
Dim dTaxaFinanceira As Double

On Error GoTo Erro_CondicaoPagto_Validate

    'Verifica se foi preenchida a ComboBox CondicaoPagamento
    If Len(Trim(CondicaoPagamento.Text)) = 0 Then
        For iIndice = 1 To objGridItens.iLinhasExistentes
            GridProdutos.TextMatrix(iIndice, iGrid_TotalPrazo_Col) = ""
            GridProdutos.TextMatrix(iIndice, iGrid_PrecoPrazo_Col) = ""
            GridProdutos.TextMatrix(iIndice, iGrid_TaxaFinanceira_Col) = ""
        Next

        Exit Sub
    
    End If
    
    'Verifica se a combo foi selecionada
    If CondicaoPagamento.ListIndex > -1 Then Exit Sub

    'Chama a função Combo_Seleciona que vai selecionar a condição
    'de pagamento caso ela esteja presente na combo.
    lErro = Combo_Seleciona(CondicaoPagamento, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 53631

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        'Preenche o objCondicaoPagamento com o código.
        objCondicaoPagamento.iCodigo = iCodigo

        'Tenta ler CondicaoPagto com esse código no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagamento)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 53632
        If lErro <> SUCESSO Then gError 53633 'Não encontrou CondicaoPagto no BD

        If lErro = SUCESSO Then
            
            If objCondicaoPagamento.iCodigo = COD_A_VISTA Then gError 63851
            
            'Coloca o "Código-NomeReduzido" na combo
            CondicaoPagamento.Text = CondPagto_Traz(objCondicaoPagamento) 'CStr(objCondicaoPagamento.iCodigo) & SEPARADOR & objCondicaoPagamento.sDescReduzida

            For iLinha = 1 To objGridItens.iLinhasExistentes
        
                dPrecoPrazo = StrParaDbl(GridProdutos.TextMatrix(iLinha, iGrid_PrecoPrazo_Col))
                dPrecoVista = StrParaDbl(GridProdutos.TextMatrix(iLinha, iGrid_PrecoVista_Col))
                
                If dPrecoPrazo <> dPrecoVista And dPrecoVista > 0 And dPrecoPrazo > 0 Then
                    lErro = Calcula_TaxaFinanceira(dPrecoPrazo, dPrecoVista, dTaxaFinanceira, CondPagto_Extrai(CondicaoPagamento))
                    If lErro <> SUCESSO Then gError 83840
                End If
            
                'coloca a taxa financeira no grid
                GridProdutos.TextMatrix(iLinha, iGrid_TaxaFinanceira_Col) = Format(dTaxaFinanceira, "Percent")
    
             Next

        End If

    End If

    'Não encontrou e foi retornada uma string --> Erro
    If lErro = 6731 Then gError 53634
        
    Exit Sub

Erro_CondicaoPagto_Validate:

    Cancel = True

    Select Case gErr

        Case 53631, 53632, 83840

        Case 53633  'Não encontrou CondicaoPagto no BD

            'Pergunta se deseja criar a condição de pagamento
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAO_PAGAMENTO")

            If vbMsgRes = vbYes Then
                'Chama a tela de CondicaoPagto
                Call Chama_Tela("CondicoesPagto", objCondicaoPagamento)

            Else
                'Segura o foco

            End If

        Case 53634
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", gErr, CondicaoPagamento.Text)
            
        Case 63851
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAOPAGTO_NAO_DISPONIVEL", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164725)

    End Select

    Exit Sub

End Sub

Private Sub CondicaoPagtoLabel_Click()

Dim objCondicaoPagto As New ClassCondicaoPagto
Dim colSelecao As New Collection

    'Se a condição de pagamento estiver preenchida,coloca o código
    'da Condição de pagamento no objCondicaoPagto.iCodigo
    If Len(Trim(CondicaoPagamento.Text)) > 0 Then
       objCondicaoPagto.iCodigo = CondPagto_Extrai(CondicaoPagamento)
    End If

    'Chama a tela CondicaoPagtoCPLista
    Call Chama_Tela("CondicaoPagtoCPLista", colSelecao, objCondicaoPagto, objEventoCondicaoPagto)

    Exit Sub

End Sub

Private Sub Contato_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Data_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Data_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEmissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEmissao_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataValidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataValidade_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataValidade, iAlterado)
        
End Sub

Private Sub DataValidade_Validate(Cancel As Boolean)
'Testa se a data inserida no campo Data Validade é válida
'Verifica se a data de validade é maior do que a data de geração do pedido de cotação

Dim lErro As Long

On Error GoTo Erro_DataValidade_Validate

    If Len(Trim(DataValidade.ClipText)) = 0 Then Exit Sub

    'Critica o valor data informado
    lErro = Data_Critica(DataValidade.Text)
    If lErro <> SUCESSO Then gError 53625
    
    '******************** Alteração feita por Luiz em 06/04/2001 ****************************
    'Inclusão do teste abaixo
    
    'Se a data de validade é menor do que a data de geração do pedido de cotação
    If StrParaDate(DataValidade.Text) < StrParaDate(Data.Caption) Then gError 79989
    '****************************************************************************************

    Exit Sub

Erro_DataValidade_Validate:

    Cancel = True


    Select Case gErr

        Case 53625 'Erro tratado na rotina chamada.

        Case 79989
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_VALIDADE_MENOR_DATA_COTACAO", gErr, Data.Caption, DataValidade.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164726)

    End Select

    Exit Sub

End Sub

Private Sub DescontoPrazo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescontoPrazo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DescontoPrazo_Validate

    'Se DescontoPrazo estiver preenchido, faz a crítica do valor informado
    'e coloca o valor formatado na tela.
    If Len(Trim(DescontoPrazo.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(DescontoPrazo.Text)
        If lErro <> SUCESSO Then gError 53816

        'Coloca o desconto a prazo na tela formatado.
        DescontoPrazo.Text = Format(DescontoPrazo.Text, "Standard")

    End If

    'Chama a função Totais_Calcula que calcula os totais do pedido
    lErro = Totais_Calcula()
    If lErro <> SUCESSO Then gError 53817

    Exit Sub

Erro_DescontoPrazo_Validate:

    Cancel = True


    Select Case gErr

        Case 53816, 53817 'Erros tratados nas rotinas chamadas.

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164727)

    End Select

    Exit Sub

End Sub

Private Sub DescontoVista_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescontoVista_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DescontoVista_Validate

    'Se DescontoVista estiver preenchido, faz a crítica do valor informado
    'e coloca o valor formatado na tela.
    If Len(Trim(DescontoVista.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(DescontoVista.Text)
        If lErro <> SUCESSO Then gError 53818

        'Coloca o desconto a prazo na tela formatado.
        DescontoVista.Text = Format(DescontoVista.Text, "Standard")

    End If

    'Chama a função Totais_Calcula que calcula os totais do pedido
    lErro = Totais_Calcula()
    If lErro <> SUCESSO Then gError 53819

    Exit Sub

Erro_DescontoVista_Validate:

    Cancel = True


    Select Case gErr

        Case 53818, 53819 'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164728)

    End Select

    Exit Sub

End Sub

Private Sub DescProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescProduto_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub DescProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub DescProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PrecoVista
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Filial_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialEmpresa_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialEmpresa_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialFornec_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialFornec_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_Activate()

Dim colSelecao As New Collection
Dim objPedidoCotacao As ClassPedidoCotacao


'    'Se no Trata_Parametros nenhuma Pedido de Cotacao foi passado
'    If iChamaTela = 1 Then
'
'        'Chama a tela PedidoCotAtualizarLista
'        Call Chama_Tela("PedidoCotAtualizarLista", colSelecao, objPedidoCotacao, objEventoBotaoPedAtualizar)
'        iChamaTela = 0
'
'    End If
    
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

        gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_Load()

Dim lErro As Long
Dim objComprador As New ClassComprador
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Form_Load

    gbCarregandoTela = True
    
    iFrameTipoDestinoAtual = 0

    giFrameAtual = 1
    
    bExibirColReal = True
    
    giMoedaAnterior = 0
     
    '##################################
    'Inserido por Wagner
    Call Formata_Controles
    '##################################
     
    Set objEventoCodigo = New AdmEvento
    Set objEventoCondicaoPagto = New AdmEvento
    
    Set objEventoBotaoPedAtualizados = New AdmEvento
    Set objEventoBotaoPedAtualizar = New AdmEvento
    
    Set gobjPedidoCotacao = New ClassPedidoCotacao
    Set gColMoedasUsadas = New Collection
    
    'Função que carrega a combo de Condição de Pagamento
    'lErro = Carrega_CondicaoPagamento()
    lErro = CF("Carrega_CondicaoPagamento", CondicaoPagamento, MODULO_CONTASAPAGAR)
    If lErro <> SUCESSO Then gError 53620

    'Inicializa a máscara do produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 53621

    Quantidade.Format = FORMATO_ESTOQUE
    QuantEntrega.Format = FORMATO_ESTOQUE

    objComprador.sCodUsuario = gsUsuario

    'Verifica se o usuário do sistema é um comprador
    lErro = CF("Comprador_Le_Usuario", objComprador)
    If lErro <> SUCESSO And lErro <> 50059 Then gError 53623
    'Se não achou o comprador --> erro.
    If lErro = 50059 Then gError 53624

    giPodeAumentarQuant = objComprador.iAumentaQuant

    objUsuarios.sCodUsuario = objComprador.sCodUsuario
    
    '** Ler NomeReduzido
    'Ok
    'Lê o usuário contido na tabela de Usuários
    lErro = CF("Usuarios_Le", objUsuarios)
    If lErro <> SUCESSO And lErro <> 40832 Then gError 53657
    If lErro <> SUCESSO Then gError 53658

    'Coloca nome reduzido do Comprador na tela
    Comprador.Caption = objUsuarios.sNomeReduzido
    
    Set objGridItens = New AdmGrid
    
    'Inicializa o grid de itens
    lErro = Inicializa_Grid_Itens(objGridItens)
    If lErro <> SUCESSO Then gError 53622
    
    Set objGridRequisicoes = New AdmGrid
    
    'Inicializa o grid de itens
    lErro = Inicializa_Grid_Req(objGridRequisicoes)
    If lErro <> SUCESSO Then gError 178861
    
    'Carrega a combo de moedas
    lErro = Carrega_Moeda()
    If lErro <> SUCESSO Then gError 108950
    
    gbCarregandoTela = False

    iAlterado = 0
    
    gbPrecoGridAlterado = False
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 53620, 53621, 53622, 53623, 53657, 108950, 178861

        Case 53624
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_COMPRADOR", gErr, objComprador.sCodUsuario)

        Case 53658
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", gErr, objUsuarios.sCodUsuario)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164729)

    End Select
    
    gbCarregandoTela = False

    iAlterado = 0
    
    gbPrecoGridAlterado = False
    
End Sub

'Function Calcula_PrecoPrazo(iNumeroParcelas As Integer, dPrecoTotVista As Double, dTaxaFinanceira As Double, dPrecoTotPrazo As Double) As Long
''Calcula o TotalPrazo em cima dos dados passados
'
'Dim iIndice As Integer
'Dim dTotalPrazo As Double
'Dim lErro As Long
'
'On Error GoTo Erro_Calcula_PrecoPrazo
'
'    'Recolhe o preço total a vista
'    dPrecoTotPrazo = dPrecoTotVista
'
'    'Se for positivo
'    If dPrecoTotVista > 0 Then
'
'        'De acordo com o número de parcelas
'        For iIndice = 1 To iNumeroParcelas
'
'            'Calcula o preço com a taxa financeira
'            dPrecoTotPrazo = dPrecoTotPrazo * (1 + dTaxaFinanceira)
'            dTotalPrazo = dTotalPrazo + dPrecoTotPrazo 'Acumula o preço total
'
'        Next
'
'    End If
'
'    'Divide a soma dos preços a prazo pelo número de parcelas
'    dPrecoTotPrazo = dTotalPrazo / iNumeroParcelas
'
'    Calcula_PrecoPrazo = SUCESSO
'
'    Exit Function
'
'Erro_Calcula_PrecoPrazo:
'
'    Calcula_PrecoPrazo = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164730)
'
'    End Select
'
'End Function



Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Libera as variáveis globais
    Set objGridItens = Nothing
    Set objGridRequisicoes = Nothing
    
    Set objEventoCodigo = Nothing
    Set objEventoCondicaoPagto = Nothing
    Set objEventoBotaoPedAtualizar = Nothing
    Set objEventoBotaoPedAtualizados = Nothing
    
    Set gobjPedidoCotacao = Nothing
    Set gColMoedasUsadas = Nothing
    
    'Libera o comando de setas
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Function Limpa_Tela_PedidoCotacao()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_PedidoCotacao

    Call Limpa_Tela(Me)
    
    gbPrecoGridAlterado = False
    
    giMoedaAnterior = 0
    
    'Limpa o restante da tela
    Codigo.Text = ""
    Fornecedor.Caption = ""
    Data.Caption = ""
    DataEmissao.Caption = ""
    TipoFrete.ListIndex = -1
    Filial.Caption = ""
    DataValidade.PromptInclude = False
    DataValidade.Text = ""
    DataValidade.PromptInclude = True
    CondicaoPagamento.Text = ""
    FilialEmpresa.Caption = ""
    TotalVista.Caption = ""
    TotalPrazo.Caption = ""
    IPIValorVista.Caption = ""
    IPIValorPrazo.Caption = ""
    FornecDestino.Caption = ""
    FilialFornec.Caption = ""
    TipoDestino(0).Value = False
    TipoDestino(1).Value = False

    Moeda.ListIndex = -1
    
    'Limpa o grid
    Call Grid_Limpa(objGridItens)

    'Limpa o GridRequisicoes
    Call Grid_Limpa(objGridRequisicoes)

    Limpa_Tela_PedidoCotacao = SUCESSO

    Exit Function

Erro_Limpa_Tela_PedidoCotacao:

    Limpa_Tela_PedidoCotacao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164731)

    End Select

    Exit Function

End Function

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
    objGrid.colColuna.Add ("Forn. Exclusivo")

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
    objGrid.colCampo.Add (Exclusivo.Name)
    
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
        iGrid_Exclusivo_Col = 17
        
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
        iGrid_Exclusivo_Col = 19
        
    End If
    
    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAX_ITENS_PEDIDO_COTACAO + 1

    'Linhas visíveis do grid
    objGrid.iLinhasVisiveis = 18
    
    'largura total do grid
    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    'É proibído excluir e incluir linhas do grid
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Reseta(objGrid)
    Call Grid_Inicializa(objGrid)

    Inicializa_Grid_Itens = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Itens:

    Inicializa_Grid_Itens = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164732)

    End Select

    Exit Function

End Function

Public Function Trata_Parametros(Optional objPedidoCotacao As ClassPedidoCotacao) As Long

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim colSelecao As New Collection
Dim gEncontrou As Boolean

On Error GoTo Erro_Trata_Parametros

    'Verifica se alguma nota foi passada por parametro
    If Not (objPedidoCotacao Is Nothing) Then
        
        gEncontrou = False
        
        If objPedidoCotacao.lCodigo > 0 Then
            objPedidoCotacao.iFilialEmpresa = giFilialEmpresa
            
            'Lê o Pedido de Cotação
            lErro = CF("PedidoCotacao_Le", objPedidoCotacao)
            If lErro <> SUCESSO And lErro <> 62867 Then gError 53671
            'Se o pedido existe --> exibir seus dados
            If lErro = SUCESSO Then gEncontrou = True
        
        ElseIf objPedidoCotacao.lNumIntDoc > 0 Then
            'Janaina
            lErro = CF("PedidoCotacao_Le_NumIntDoc", objPedidoCotacao)
            'Janaina
            If lErro <> SUCESSO And lErro <> 62867 Then gError 62868
            If lErro = SUCESSO Then gEncontrou = True
        End If
                
        If gEncontrou Then
            lErro = Traz_PedidoCotacao_Tela(objPedidoCotacao)
            If lErro <> SUCESSO Then gError 53672
        End If

    Else

        iChamaTela = 1
        
    End If

    iAlterado = 0
    
    gbPrecoGridAlterado = False

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 53671, 53672, 62867 'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164733)

    End Select

    Exit Function

End Function

'Private Function Carrega_CondicaoPagamento() As Long
''Carrega a combo de Condições de Pagamento com  as Condições lidas do BD
'
'Dim lErro As Long
'Dim colCod_DescReduzida As New AdmColCodigoNome
'Dim objCod_DescReduzida As AdmCodigoNome
'
'On Error GoTo Erro_Carrega_CondicaoPagamento
'
'    'Lê o código e a descrição reduzida de todas as Condições de Pagamento
'    lErro = CF("CondicoesPagto_Le_Pagamento", colCod_DescReduzida)
'    If lErro <> SUCESSO Then gError 31185
'
'    For Each objCod_DescReduzida In colCod_DescReduzida
'
'        'Vamos carregar a Combo com as Descrições de Pagamento com exceção da condição à vista
'        If objCod_DescReduzida.iCodigo <> COD_A_VISTA Then
'
'            'Adiciona novo ítem na List da Combo CondicaoPagamento
'            CondicaoPagamento.AddItem CInt(objCod_DescReduzida.iCodigo) & SEPARADOR & objCod_DescReduzida.sNome
'            CondicaoPagamento.ItemData(CondicaoPagamento.NewIndex) = objCod_DescReduzida.iCodigo
'
'        End If
'
'    Next
'
'    Carrega_CondicaoPagamento = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_CondicaoPagamento:
'
'    Carrega_CondicaoPagamento = gErr
'
'    Select Case gErr
'
'        Case 31185 'Erro tratado na rotina chamada
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164734)
'
'    End Select
'
'    Exit Function
'
'End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

    If (objControl.Name = PrecoPrazo.Name) Or (objControl.Name = TotPrazo.Name) Or (objControl.Name = TaxaFinanceira.Name) Then

        If Len(Trim(CondicaoPagamento.Text)) = 0 Then
            objControl.Enabled = False
        Else
            objControl.Enabled = True
        End If

    End If

    If objControl.Name = Quantidade.Name Then
            
        If iLinha = 0 Then
            objControl.Enabled = False
            Exit Sub
        End If
        
        'Permite alterar a qtde se o comprador tem este atributo em seu cadastro
        If giPodeAumentarQuant = MARCADO Then
            objControl.Enabled = True
        Else
            objControl.Enabled = False
        End If
        
    End If
    
End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim iIndice As Integer
Dim objPedidoCotacao As New ClassPedidoCotacao

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "PedidoCotacao"

    'Se o código estiver preenchido, carrega o obj com o código da tela
    lErro = Move_Tela_Memoria(objPedidoCotacao)
    If lErro <> SUCESSO Then gError 62641

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objPedidoCotacao.lCodigo, 0, "Codigo"
    colCampoValor.Add "Fornecedor", objPedidoCotacao.lFornecedor, 0, "Fornecedor"
    colCampoValor.Add "Filial", objPedidoCotacao.iFilial, 0, "Filial"

    'Adiciona filtro
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "Status", OP_DIFERENTE, STATUS_ATUALIZADO

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 53666, 62641 'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164735)

    End Select

    Exit Sub

End Sub

Private Sub FornecDestino_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FornecDestino_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Fornecedor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Fornecedor_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

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

Private Sub GridProdutos_LeaveCell()

    Call Saida_Celula(objGridItens)

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


Private Sub IPIValorPrazo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub IPIValorPrazo_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub IPIValorVista_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub IPIValorVista_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Moeda_Change()
    
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Moeda_Click()

Dim lErro As Long
Dim objPedidoCotacao As New ClassPedidoCotacao

On Error GoTo Erro_Moeda_Click

    'Se nao estiver carregando a tela
    If gbCarregandoTela = False Then
        
        'Verifica se algum preco foi alterado, se foi => pergunta se deseja salvar
        If gbPrecoGridAlterado = True Then
        
            If Rotina_Aviso(vbYesNo, "AVISO_COTACAO_MOEDA_SEM_SALVAR") = vbYes Then Call SalvarGrid
            
        End If
        
    End If
    
    'Limpa a cotacao
    TaxaConversao.Text = ""
    
    'Se a moeda selecionada for = REAL
    If Codigo_Extrai(Moeda.List(Moeda.ListIndex)) = MOEDA_REAL Then
        
        'Desabilita a cotacao
        TaxaConversao.Enabled = False
        BotaoTrazCotacao.Enabled = False
        LabelTaxa.Enabled = False
        
        bExibirColReal = False
        
    Else
            
        'Habilita a cotacao
        TaxaConversao.Enabled = True
        BotaoTrazCotacao.Enabled = True
        LabelTaxa.Enabled = True
        
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
        
        lErro = Preenche_GridRequisicoes
        If lErro <> SUCESSO Then gError 178872
        
    End If
    
    Call Totais_Calcula
    
    gbPrecoGridAlterado = False
    
    If Moeda.ListIndex >= 0 Then
        giMoedaAnterior = Codigo_Extrai(Moeda.Text)
    Else
        giMoedaAnterior = MOEDA_REAL
    End If
    
    Exit Sub
    
Erro_Moeda_Click:

    Select Case gErr
    
        Case 108951, 108960, 178872
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164736)
            
    End Select
    
End Sub

Private Sub objEventoBotaoPedAtualizados_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPedidoCotacao As New ClassPedidoCotacao

On Error GoTo Erro_objEventoBotaoPedAtualizados_evSelecao

    Set objPedidoCotacao = obj1

    'Chama Traz_PedidoCotacao_Tela
    lErro = Traz_PedidoCotacao_Tela(objPedidoCotacao)
    If lErro <> SUCESSO Then gError 53662

    'Fecha o sistema de setas
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoBotaoPedAtualizados_evSelecao:

    Select Case gErr

        Case 53662 'Erro tratado na rotina chamada.

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164737)

    End Select

    Exit Sub

End Sub

Private Sub objEventoBotaoPedAtualizar_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPedidoCotacao As New ClassPedidoCotacao

On Error GoTo Erro_objEventoBotaoPedAtualizar_evSelecao

    Set objPedidoCotacao = obj1

    'Chama Traz_PedidoCotacao_Tela
    lErro = Traz_PedidoCotacao_Tela(objPedidoCotacao)
    If lErro <> SUCESSO Then gError 53661

    'Fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoBotaoPedAtualizar_evSelecao:

    Select Case gErr

        Case 53661 'Erro tratado na rotina chamada.

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164738)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPedidoCotacao As New ClassPedidoCotacao

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objPedidoCotacao = obj1

    'Chama Traz_PedidoCotacao_Tela
    lErro = Traz_PedidoCotacao_Tela(objPedidoCotacao)
    If lErro <> SUCESSO Then gError 53660

    'Fecha o sistema de setas
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 53660 'Erro tratado na rotina chamada.

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164739)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCondicaoPagto_evSelecao(obj1 As Object)

Dim objCondicaoPagto As New ClassCondicaoPagto
Dim objPedidoCotacao As New ClassPedidoCotacao
Dim lErro As Long

On Error GoTo Erro_objEventoCondicaoPagto_evSelecao


    Set objCondicaoPagto = obj1

    If objCondicaoPagto.iCodigo <> COD_A_VISTA Then
        
        'Coloca o "Código-NomeReduzido" na combo
        CondicaoPagamento.Text = CStr(objCondicaoPagto.iCodigo) '& SEPARADOR & objCondicaoPagto.sDescReduzida
        Call CondicaoPagamento_Validate(bSGECancelDummy)

    End If
    
    If objCondicaoPagto.iCodigo = COD_A_VISTA Then gError 74938
    
    Me.Show

    Exit Sub
    
Erro_objEventoCondicaoPagto_evSelecao:

    Select Case gErr
    
        Case 74938
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAOPAGTO_NAO_DISPONIVEL", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164740)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Observacao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub Observacao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub Observacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Observacao
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub PrazoEntrega_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PrazoEntrega_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub PrazoEntrega_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub PrazoEntrega_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PrazoEntrega
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub PrecoPrazo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PrecoPrazo_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub PrecoPrazo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub PrecoPrazo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PrecoPrazo
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub PrecoVista_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PrecoVista_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub PrecoVista_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub PrecoVista_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PrecoVista
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantEntrega_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantEntrega_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub QuantEntrega_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub QuantEntrega_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = QuantEntrega
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TabStrip1_Click()
    
    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> giFrameAtual Then

        If TabStrip_PodeTrocarTab(giFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(giFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        giFrameAtual = TabStrip1.SelectedItem.Index

    End If

End Sub

Private Sub TaxaConversao_Change()
    iAlterado = REGISTRO_ALTERADO
    gbPrecoGridAlterado = True
End Sub

Private Sub TaxaConversao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_TaxaConversao_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(TaxaConversao.Text)) > 0 Then

        'Critica se é valor Positivo
        lErro = Valor_Positivo_Critica_Double(TaxaConversao.Text)
        If lErro <> SUCESSO Then gError 103378
    
        'Põe o valor formatado na tela
        TaxaConversao.Text = Format(TaxaConversao.Text, FORMATO_TAXA_CONVERSAO_MOEDA)
        
    End If
    
    For iLinha = 1 To objGridItens.iLinhasExistentes
        Call ComparativoMoedaReal_Calcula(StrParaDbl(TaxaConversao.Text), iLinha)
    Next

    gbPrecoGridAlterado = True
    
    Exit Sub
    
Erro_TaxaConversao_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 103378

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164741)
    
    End Select
    
End Sub

Private Sub TaxaFinanceira_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TaxaFinanceira_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub TaxaFinanceira_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub TaxaFinanceira_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = TaxaFinanceira
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub TipoDestino_Click(Index As Integer)


    If Index = iFrameTipoDestinoAtual Then Exit Sub

    'Torna Frame correspondente a Index visivel
    FrameTipo(Index).Visible = True

    'Torna Frame atual invisivel
    FrameTipo(iFrameTipoDestinoAtual).Visible = False

    'Armazena novo valor de iFrameTipoDestinoAtual
    iFrameTipoDestinoAtual = Index

   Exit Sub

End Sub

Private Sub TipoFrete_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoFrete_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TotalPrazo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TotalPrazo_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TotalVista_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TotalVista_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TotPrazo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TotPrazo_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub TotPrazo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub TotPrazo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = TotPrazo
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub TotVista_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TotVista_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub TotVista_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub TotVista_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = TotVista
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub UnidadeMed_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownDataValidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownDataValidade_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataValidade_DownClick

    lErro = Data_Up_Down_Click(DataValidade, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 53626

    Exit Sub

Erro_UpDownDataValidade_DownClick:

    Select Case gErr

        Case 53626 'Erro tratado na rotina chamada.

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164742)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataValidade_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataValidade_UpClick

    lErro = Data_Up_Down_Click(DataValidade, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 53627

    Exit Sub

Erro_UpDownDataValidade_UpClick:

    Select Case gErr

        Case 53627 'Erro tratado na rotina chamada.

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164743)

    End Select

    Exit Sub

End Sub

Private Sub ValorDesconto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is CondicaoPagamento Then
            Call CondicaoPagtoLabel_Click
        ElseIf Me.ActiveControl Is Codigo Then
            Call CodigoLabel_Click
        End If
    End If
    
End Sub

Private Sub UserControl_Terminate()
'If giDebug = 1 Then MsgBox ("Saiu")
End Sub

Private Sub ValorDespesas_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorDespesas_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorDespesas_Validate

    'Se ValorDespesas estiver preenchido, faz a crítica do valor informado
    'e coloca o valor formatado na tela.
    If Len(Trim(ValorDespesas.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(ValorDespesas.Text)
        If lErro <> SUCESSO Then gError 53630

        ValorDespesas.Text = Format(StrParaDbl(ValorDespesas.Text), "Standard")

    End If

    'Chama a função Totais_Calcula que calcula os totais do pedido
    lErro = Totais_Calcula()
    If lErro <> SUCESSO Then gError 53684

    Exit Sub

Erro_ValorDespesas_Validate:

    Cancel = True


    Select Case gErr

        Case 53630

        Case 53684 'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164744)

    End Select

    Exit Sub

End Sub

Private Sub ValorFrete_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorFrete_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorFrete_Change

    'Se ValorFrete estiver preenchido, faz a crítica do valor informado
    'e coloca o valor formatado na tela.
    If Len(Trim(ValorFrete.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(ValorFrete.Text)
        If lErro <> SUCESSO Then gError 53629

        ValorFrete.Text = Format(StrParaDbl(ValorFrete.Text), "Standard")

    End If

    'Chama a função Totais_Calcula que calcula os totais do pedido
    lErro = Totais_Calcula()
    If lErro <> SUCESSO Then gError 53685

    Exit Sub

Erro_ValorFrete_Change:

    Select Case gErr

        Case 53629
            ValorFrete.SetFocus

        Case 53685 'Erro tratado na rotina chamada.

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164745)

    End Select

    Exit Sub

End Sub

Private Sub ValorIPI_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorIPIPrazo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorIPIVista_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorSeguro_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorSeguro_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorSeguro_Validate

    'Se ValorSeguro estiver preenchido, faz a crítica do valor informado
    'e coloca o valor formatado na tela.
    If Len(Trim(ValorSeguro.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(ValorSeguro.Text)
        If lErro <> SUCESSO Then gError 53628

        ValorSeguro.Text = Format(StrParaDbl(ValorSeguro.Text), "Standard")

    End If

    'Chama a função Totais_Calcula que calcula os totais do pedido
    lErro = Totais_Calcula()
    If lErro <> SUCESSO Then gError 53686

    Exit Sub

Erro_ValorSeguro_Validate:

    Cancel = True


    Select Case gErr

        Case 53628

        Case 53686 'Erros tratados nas rotinas chamadas.

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164746)

    End Select

    Exit Sub

End Sub

Function Traz_PedidoCotacao_Tela(objPedidoCotacao As ClassPedidoCotacao) As Long

Dim lErro As Long
Dim iIndice2 As Integer
Dim objCotacao As New ClassCotacao
Dim objItemPedCotacao As New ClassItemPedCotacao
Dim iIndice As Integer
Dim objItensCotacao As ClassItemCotacao
Dim dValorFrete As Double
Dim dValorSeguro As Double
Dim dValorDespesa As Double
Dim objCliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objFornecedor As New ClassFornecedor
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Traz_PedidoCotacao_Tela

    'Guarda na variável global o pedido que está sendo trazido para a tela ...
    Set gobjPedidoCotacao = objPedidoCotacao
    
    'Limpa a tela
    Call Limpa_Tela_PedidoCotacao
    
    gbPrecoGridAlterado = False

    'Lê os itens do pedido.
    lErro = CF("ItensPedCotacao_Le", objPedidoCotacao)
    If lErro <> SUCESSO Then gError 53641

    objCotacao.lNumIntDoc = objPedidoCotacao.lCotacao

    'Lê a tabela de Cotação
    lErro = CF("Cotacao_Le", objCotacao)
    If lErro <> SUCESSO And lErro <> 53649 Then gError 53654
    If lErro <> SUCESSO Then gError 53655

    'Preenche a tela com os dados de PedidoCotacao e Cotacao
    Codigo.Text = objPedidoCotacao.lCodigo
    If objPedidoCotacao.iCondPagtoPrazo <> 0 And objPedidoCotacao.iCondPagtoPrazo <> COD_A_VISTA Then

        CondicaoPagamento.Text = objPedidoCotacao.iCondPagtoPrazo
        Call CondicaoPagamento_Validate(bSGECancelDummy)

    End If
    
    'Se a Cotação possui Tipo de Destino
    If objCotacao.iTipoDestino <> TIPO_DESTINO_AUSENTE Then
    
        'Preenche o TipoDestino
        TipoDestino(objCotacao.iTipoDestino).Value = True
    
        If iFrameTipoDestinoAtual = TIPO_DESTINO_EMPRESA Then
    
            objFilialEmpresa.iCodFilial = objCotacao.iFilialEmpresa
    
            'Lê a FilialEmpresa
            lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
            If lErro <> SUCESSO And lErro <> 27378 Then gError 53803
            If lErro = 27378 Then gError 53804
    
            'Coloca a FilialEmpresa na tela
            FilialEmpresa.Caption = objFilialEmpresa.sNome
    
        ElseIf iFrameTipoDestinoAtual = TIPO_DESTINO_FORNECEDOR Then
    
            objFornecedor.lCodigo = objCotacao.lFornCliDestino
    
            'Lê o Fornecedor
            lErro = CF("Fornecedor_Le", objFornecedor)
            If lErro <> SUCESSO And lErro <> 18272 Then gError 53805
            If lErro = 18272 Then gError 53806
    
            'Coloca o Fornecedor na tela.
            FornecDestino.Caption = objFornecedor.sNomeReduzido
    
            objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
            objFilialFornecedor.iCodFilial = objCotacao.iFilialDestino
    
            'Lê a FilialFornecedor
            lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", objFornecedor.sNomeReduzido, objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 18272 Then gError 53807
            'Se nao encontrou
            If lErro = 18272 Then gError 53808
    
            'Coloca a Filial na tela
            FilialFornec.Caption = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome
    
        End If
    
    End If
    
    'Passa o código do fornecedor de objpedidocotacao para o objfornecedor
    objFornecedor.lCodigo = objPedidoCotacao.lFornecedor

    'Lê o Fornecedor
    lErro = CF("Fornecedor_Le", objFornecedor)
    If lErro <> SUCESSO And lErro <> 18272 Then gError 53811
    If lErro = 18272 Then gError 53812

    'Coloca o NomeReduzido do Fornecedor na tela
    Fornecedor.Caption = objFornecedor.sNomeReduzido

    'Passa o CodFornecedor e o CodFilial para o objfilialfornecedor
    objFilialFornecedor.lCodFornecedor = objPedidoCotacao.lFornecedor
    objFilialFornecedor.iCodFilial = objPedidoCotacao.iFilial

    'Lê o filialforncedor
    lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", objFornecedor.sNomeReduzido, objFilialFornecedor)
    If lErro <> SUCESSO And lErro <> 18272 Then gError 53813
    'Se nao encontrou
    If lErro = 18272 Then gError 53814

    'Coloca a filial na tela
    Filial.Caption = objPedidoCotacao.iFilial & SEPARADOR & objFilialFornecedor.sNome

    Contato.Text = objPedidoCotacao.sContato
    
    If objPedidoCotacao.dtData <> DATA_NULA Then
        Data.Caption = Format(objPedidoCotacao.dtData, "dd/mm/yyyy")
    End If
    
    'Preenche a DataEmissao
    If objPedidoCotacao.dtDataEmissao <> DATA_NULA Then
        DataEmissao.Caption = Format(objPedidoCotacao.dtDataEmissao, "dd/mm/yyyy")
    End If

    'Preenche a DataValidade
    If objPedidoCotacao.dtDataValidade <> DATA_NULA Then
        DataValidade.PromptInclude = False
        DataValidade.Text = Format(objPedidoCotacao.dtDataValidade, "dd/mm/yy")
        DataValidade.PromptInclude = True
    End If
    
    'Preenche a combo TipoFrete
    For iIndice = 0 To TipoFrete.ListCount - 1
        If objPedidoCotacao.iTipoFrete = TipoFrete.ItemData(iIndice) Then
            TipoFrete.ListIndex = iIndice
        End If
    Next
    
    'Seleciona a Moeda
    For Each objItemPedCotacao In objPedidoCotacao.colItens
    
        If objItemPedCotacao.colItensCotacao.Count > 0 Then
        
            For iIndice2 = 0 To Moeda.ListCount - 1
            
                If Codigo_Extrai(Moeda.List(iIndice2)) = objItemPedCotacao.colItensCotacao.Item(1).iMoeda Then
                    Moeda.ListIndex = iIndice2
                    Exit For
                End If
                
            Next
            
            If Moeda.ListIndex >= 0 Then Exit For
        
        End If
        
    Next
    
    'Se nao selecionou => MOEDA_REAL
    If Moeda.ListIndex < 0 Then Moeda.ListIndex = MOEDA_REAL
    
    'Chama a função Totais_Calcula que calcula os totais do pedido
    lErro = Totais_Calcula()
    If lErro <> SUCESSO Then gError 53815

    iAlterado = 0
    
    gbPrecoGridAlterado = False

    Traz_PedidoCotacao_Tela = SUCESSO

    Exit Function

Erro_Traz_PedidoCotacao_Tela:

    Traz_PedidoCotacao_Tela = gErr

    Select Case gErr

        Case 53641, 53654, 53656, 53803, 53805, 53807, 53811, 53813, 53815   'Erros tratados nas rotinas chamadas.

        Case 53804
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr)

        Case 53806
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO_2", gErr)

        Case 53808, 53814
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_FORNECEDOR_INEXISTENTE", gErr, objFilialFornecedor.lCodFornecedor, objFilialFornecedor.iCodFilial)

        Case 53812
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)

        Case 53655
            Call Rotina_Erro(vbOKOnly, "ERRO_COTACAO_NAO_CADASTRADA", gErr, objPedidoCotacao.lCotacao)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164747)

    End Select

    Exit Function

End Function

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objPedidoCotacao As New ClassPedidoCotacao

On Error GoTo Erro_Tela_Preenche

    gbPrecoGridAlterado = False
    
    Set gobjPedidoCotacao = New ClassPedidoCotacao
    
    'Passa o código da coleção de campos valores para o objPedidoCotacao
    gobjPedidoCotacao.lCodigo = colCampoValor.Item("Codigo").vValor

    'Coloca giFilialEmpresa em objPedidoCotacao.giFilialEmpresa
    gobjPedidoCotacao.iFilialEmpresa = giFilialEmpresa

    'Lê o Pedido de Cotação
    lErro = CF("PedidoCotacao_Le", gobjPedidoCotacao)
    If lErro <> SUCESSO And lErro <> 62867 Then gError 53642
    If lErro = 62867 Then gError 53822

    'Chama Traz_PedidoCotacao_Tela
    lErro = Traz_PedidoCotacao_Tela(gobjPedidoCotacao)
    If lErro <> SUCESSO Then gError 53643

    iAlterado = 0
    
    gbPrecoGridAlterado = False

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 53642, 53643 'Erros tratados nas rotinas chamadas.

        Case 53822
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOTACAO_NAO_ENCONTRADO", gErr, objPedidoCotacao.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164748)

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
    dValorFrete = StrParaDbl(ValorFrete.Text)
    dSeguro = StrParaDbl(ValorSeguro.Text)
    dDescontoVista = StrParaDbl(DescontoVista.Text)
    dDescontoPrazo = StrParaDbl(DescontoPrazo.Text)

    'Coloca o IPI na tela
    IPIValorVista.Caption = Format(dValorIPITotalVista, "Standard")
    IPIValorPrazo.Caption = Format(dValorIPITotalPrazo, "Standard")

    'Depois de somar o TotalVista dos itens, somar o resultado
    'com o ValorFrete Seguro, Desconto e IPI.
    dValorTotalVista = dValorTotalVista + dValorFrete + dSeguro + dValorIPITotalVista - dDescontoVista + StrParaDbl(ValorDespesas.Text)

    'Coloca o TotalVista na tela
    TotalVista.Caption = Format(dValorTotalVista, TotalVIstaRS.Format) 'Alterado por Wagner

    'Depois de somar o TotalPrazo dos itens, somar o resultado
    'com o ValorFrete Seguro, Desconto e IPI.
    dValorTotalPrazo = dValorTotalPrazo + dValorFrete + dSeguro + dValorIPITotalPrazo - dDescontoPrazo + StrParaDbl(ValorDespesas.Text)

    'Coloca o TotalPrazo na tela
    TotalPrazo.Caption = Format(dValorTotalPrazo, TotalPrazoRS.Format) 'Alterado por Wagner

    Totais_Calcula = SUCESSO

    Exit Function

Erro_Totais_Calcula:

    Totais_Calcula = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164749)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PrecoVista(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim bPrecoAlterado As Boolean
Dim dPrecoVista As Double
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim dTotVista As Double
Dim dTaxaFinanceira As Double
Dim dTotPrazo As Double
Dim dQuantidade As Double
Dim dAliquotaIPI As Double
Dim dPrecoPrazo As Double
Dim dCotacao As Double

On Error GoTo Erro_Saida_Celula_PrecoVista

    Set objGridInt.objControle = PrecoVista

    bPrecoAlterado = False

    'Se o preço a vista foi preenchido, faz a critica do seu valor e o coloca na tela.
    If Len(Trim(PrecoVista.Text)) > 0 Then

        'Verifica se o valor é positivo
        lErro = Valor_Positivo_Critica(PrecoVista.Text)
        If lErro <> SUCESSO Then gError 53687

        'Formata o valor
        PrecoVista.Text = Format(PrecoVista.Text, gobjCOM.sFormatoPrecoUnitario) ' "STANDARD") 'Alterado por Wagner

        'Se o produto estiver preenchido
        If Len(Trim(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoPrazo_Col))) > 0 Then

            If StrParaDbl(PrecoVista) > (GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoPrazo_Col)) Then GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoPrazo_Col) = ""

        End If
        
    'Se o preço a Vista foi apagado
    Else

        'Se não foi preenchido limpa taxa financeira e o total a vista
        GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TaxaFinanceira_Col) = ""
        GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalVista_Col) = ""
        
    End If

    'Se o preco.text for diferente do preço a vista no grid, faz
    'bPrecoAlterado=True
    If StrParaDbl(PrecoVista.Text) <> StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoVista_Col)) Then

        bPrecoAlterado = True
        gbPrecoGridAlterado = True

    End If
    
    'Chama Grid_Abandona_Celula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 53688

    If bPrecoAlterado = True Then

        'Converte os valores do PrecoVista e da Quantidade
        dPrecoVista = StrParaDbl(PrecoVista.Text)
        dQuantidade = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Quantidade_Col))

        'Calcula o total a vista para o item.
        dTotVista = dPrecoVista * dQuantidade

        If dTotVista <> 0 Then
            'Coloca o Total a Vista no grid
            GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalVista_Col) = Format(dTotVista, TotalVIstaRS.Format) 'Alterado por Wagner
        End If

        If Len(Trim(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoPrazo_Col))) = 0 And dPrecoVista > 0 Then

            'Passa o que estiver escrito na combo para o objCondicaoPagto
            objCondicaoPagto.iCodigo = CondPagto_Extrai(CondicaoPagamento)

            'Se a Condição foi informada, lê a Condição de Pagamento
            lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
            If lErro <> SUCESSO And lErro <> 19205 Then gError 53689
            
            dTaxaFinanceira = objCondicaoPagto.dAcrescimoFinanceiro
            
            'Faz o cálculo do preço a prazo a partir da taxa financeira, PrecoVista e CondicaoPagto
            lErro = Calcula_PrecoPrazo(dPrecoPrazo, dPrecoVista, dTaxaFinanceira, objCondicaoPagto.iCodigo)
            If lErro <> SUCESSO Then gError 53689
            
            '##################################################
            'Inserido por Wagner 17/05/2006
            lErro = CF("PrecoPrazo_Customizado", dPrecoPrazo)
            If lErro <> SUCESSO Then gError 177422
            '##################################################

            'Calcula o PrecoPrazo
            dTotPrazo = dPrecoPrazo * dQuantidade
            
            'De acordo com a quantidade calcula o preço a prazo e
            'coloca na tela.
            If dPrecoPrazo > 0 Then
                GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoPrazo_Col) = Format(dPrecoPrazo, gobjCOM.sFormatoPrecoUnitario) ' "STANDARD") 'Alterado por Wagner
                GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalPrazo_Col) = Format(dTotPrazo, TotalPrazoRS.Format) 'Alterado por Wagner
            Else
                GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoPrazo_Col) = ""
                GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalPrazo_Col) = ""
            End If

            GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TaxaFinanceira_Col) = Format(dTaxaFinanceira, "Percent")
        
            'Atualiza o IPI a Prazo do item
            GridProdutos.TextMatrix(GridProdutos.Row, iGrid_ValorIPIPrazo_Col) = dTotPrazo * dAliquotaIPI

        ElseIf Len(Trim(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoPrazo_Col))) > 0 And dPrecoVista > 0 Then

            'Passa o que estiver escrito na combo para o objCondicaoPagto
            objCondicaoPagto.iCodigo = CondPagto_Extrai(CondicaoPagamento)

            dPrecoPrazo = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoPrazo_Col))

            lErro = Calcula_TaxaFinanceira(dPrecoPrazo, dPrecoVista, dTaxaFinanceira, objCondicaoPagto.iCodigo)
            If lErro <> SUCESSO Then gError 76208
            
            GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TaxaFinanceira_Col) = Format(dTaxaFinanceira, "Percent")

        End If

        'Atualiza o IPI a Vista do item
        dAliquotaIPI = PercentParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_AliquotaIPI_Col))
        GridProdutos.TextMatrix(GridProdutos.Row, iGrid_ValorIPIVista_Col) = dTotVista * dAliquotaIPI

        'Atualiza o Total a Prazo, a Vista do pedido e os valores do IPI do pedido
        lErro = Totais_Calcula()
        If lErro <> SUCESSO Then gError 53691

    End If
    
    If Len(Trim(TaxaConversao.Text)) > 0 Then Call ComparativoMoedaReal_Calcula(CDbl(TaxaConversao.Text), GridProdutos.Row)

    Saida_Celula_PrecoVista = SUCESSO

    Exit Function

Erro_Saida_Celula_PrecoVista:

    Saida_Celula_PrecoVista = gErr

    Select Case gErr

        Case 53687, 53688, 53691 'Erros tratados nas rotinas chamadas
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 53689, 76208, 177422
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 53690 'Não encontrou CondicaoPagto no BD
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 76205
            Call Rotina_Erro(vbOKOnly, "ERRO_PRECOPRAZO_MENOR_PRECOVISTA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164750)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Observacao(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Observacao

    Set objGridInt.objControle = Observacao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 53644

    Saida_Celula_Observacao = SUCESSO

    Exit Function

Erro_Saida_Celula_Observacao:

    Saida_Celula_Observacao = gErr

    Select Case gErr

        Case 53644 'Erro tratado na rotina chamada.
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164751)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PrazoEntrega(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_PrazoEntrega

    Set objGridInt.objControle = PrazoEntrega

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 53645

    Saida_Celula_PrazoEntrega = SUCESSO

    Exit Function

Erro_Saida_Celula_PrazoEntrega:

    Saida_Celula_PrazoEntrega = gErr

    Select Case gErr

        Case 53645 'Erro tratado na rotina chamada.
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164752)

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
        'essa observacao vem das requsicoes. Se objItensCotacao estiver preenchido para este produto
        'vai se sobrepor ao que veio das requisicoes
        GridProdutos.TextMatrix(iIndice, iGrid_Observacao_Col) = objItemPedCotacao.sObservacao
        
        'Janaina
        lErro = CF("ItemPedCotacao_Le_CotacaoProduto", objItemPedCotacao, objCotacaoProduto)
        'Janaina
        If lErro <> SUCESSO And lErro <> 76250 Then gError 76251
        If lErro = 76250 Then gError 76252
        
        If objCotacaoProduto.lFornecedor <> 0 And objCotacaoProduto.iFilial <> 0 Then
            GridProdutos.TextMatrix(iIndice, iGrid_Exclusivo_Col) = MARCADO
        Else
            GridProdutos.TextMatrix(iIndice, iGrid_Exclusivo_Col) = DESMARCADO
        End If
        
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
                If StrParaDbl(TaxaConversao.Text) <> objItensCotacao.dTaxa And objItensCotacao.iMoeda <> MOEDA_REAL Then TaxaConversao.Text = Format(objItensCotacao.dTaxa, TaxaConversao.Format)
                
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
                iCondPagto = CondPagto_Extrai(CondicaoPagamento)
                
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
                dTaxa = StrParaDbl(TaxaConversao.Text)
                If dTaxa > 0 Then
                    'Calcula a Taxa Financeira para o ItemCotacao
                    lErro = Calcula_TaxaFinanceira(dValorTotalPrazo * dTaxa, dValorTotalVista * dTaxa, dTaxaFinanceira, iCondPagto)
                    If lErro <> SUCESSO Then gError 76431
                Else
                    'Calcula a Taxa Financeira para o ItemCotacao
                    lErro = Calcula_TaxaFinanceira(dValorTotalPrazo, dValorTotalVista, dTaxaFinanceira, iCondPagto)
                    If lErro <> SUCESSO Then gError 76431
                End If
            
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
    ValorFrete.Text = Format(dValorFrete, "Standard")
    ValorSeguro.Text = Format(dValorSeguro, "Standard")
    ValorDespesas.Text = Format(dValorDespesa, "Standard")
    DescontoVista.Text = Format(dValorDescontoVista, "Standard")
    DescontoPrazo.Text = Format(dValorDescontoPrazo, "Standard")

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164753)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(objPedidoCotacao As ClassPedidoCotacao) As Long

Dim lErro As Long
Dim objCotacao As New ClassCotacao
Dim objItensCotacao As New ClassItemCotacao
Dim objItemPedCotacao As New ClassItemPedCotacao
Dim objFornecedor As New ClassFornecedor
Dim sNomeReduzido As String
Dim objFilialEmpresa As New AdmFiliais
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objCliente As New ClassCliente

On Error GoTo Erro_Move_Tela_Memoria

    'Move os dados de pedidocotacao para a memória.
    objPedidoCotacao.lCodigo = StrParaLong(Codigo.Text)
    objPedidoCotacao.dtData = StrParaDate(Data.Caption)
    objPedidoCotacao.dtDataEmissao = StrParaDate(DataEmissao.Caption)
    objPedidoCotacao.iCondPagtoPrazo = CondPagto_Extrai(CondicaoPagamento)

    If Len(Trim(DataValidade.ClipText)) > 0 Then
        objPedidoCotacao.dtDataValidade = StrParaDate(DataValidade.Text)
    Else
        objPedidoCotacao.dtDataValidade = DATA_NULA
    End If

    objPedidoCotacao.sContato = Contato.Text

    If Len(Trim(Fornecedor.Caption)) > 0 Then
        
        objFornecedor.sNomeReduzido = Fornecedor.Caption
        
        'Função que lê o fornecedor.
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 53802
        
        If lErro = 6681 Then gError 53823

        objPedidoCotacao.lFornecedor = objFornecedor.lCodigo
    
    End If
    
    objPedidoCotacao.iFilial = Codigo_Extrai(Filial.Caption)
    objPedidoCotacao.iFilialEmpresa = giFilialEmpresa

    If Len(Trim(TipoFrete.Text)) > 0 Then
        objPedidoCotacao.iTipoFrete = TipoFrete.ItemData(TipoFrete.ListIndex)
    End If

    'Frame Local de Entrega
    If TipoDestino(TIPO_DESTINO_EMPRESA) Then

        objCotacao.iTipoDestino = TIPO_DESTINO_EMPRESA
        objFilialEmpresa.sNome = FilialEmpresa.Caption

    ElseIf TipoDestino(TIPO_DESTINO_FORNECEDOR) Then

        objCotacao.iTipoDestino = TIPO_DESTINO_FORNECEDOR
        objFornecedor.sNomeReduzido = Fornecedor.Caption
        objFilialFornecedor.iCodFilial = Codigo_Extrai(Filial.Caption)

    End If
    
    Call Determina_Moedas_Usadas
    
    Call Determina_Status_PedCompra

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 53802 'Erro tratado na rotina chamada.

        Case 53823
           Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164754)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_AliquotaIPI(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dAliquotaIPI As Double
Dim dTotVista As Double
Dim dTotPrazo As Double
Dim dPrecoVista As Double
Dim dPrecoPrazo As Double
Dim dValorIPIVista As Double
Dim dValorIPIPrazo As Double

On Error GoTo Erro_Saida_Celula_AliquotaIPI

    Set objGridInt.objControle = AliquotaIPI

    'Verifica se a AlíquotaIPI está preenchida
    If Len(Trim(AliquotaIPI.Text)) > 0 Then

        'Critica o valor informado
        lErro = Porcentagem_Critica(AliquotaIPI.Text)
        If lErro <> SUCESSO Then gError 53673

        'Coloca em AliquotaIPI.Text a alíquota com o formato Fixed.
        AliquotaIPI.Text = Format(AliquotaIPI.Text, "Fixed")

        dAliquotaIPI = StrParaDbl(AliquotaIPI.Text)
        dTotVista = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalVista_Col))
        dTotPrazo = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalPrazo_Col))

        'Calcula o valor IPI em cima da alíquota informada
        If (dTotVista > 0 Or dTotPrazo > 0) Then

            dValorIPIVista = (dAliquotaIPI * dTotVista) / 100
            GridProdutos.TextMatrix(GridProdutos.Row, iGrid_ValorIPIVista_Col) = Format(dValorIPIVista, "Fixed")

            dValorIPIPrazo = (dAliquotaIPI * dTotPrazo) / 100
            GridProdutos.TextMatrix(GridProdutos.Row, iGrid_ValorIPIPrazo_Col) = Format(dValorIPIPrazo, "Fixed")

        End If

    Else
    
        GridProdutos.TextMatrix(GridProdutos.Row, iGrid_ValorIPIVista_Col) = ""
        GridProdutos.TextMatrix(GridProdutos.Row, iGrid_ValorIPIPrazo_Col) = ""
        
    End If

    'Atualiza o Total a Prazo, a Vista do pedido e os valores do IPI do pedido
    lErro = Totais_Calcula()
    If lErro <> SUCESSO Then gError 74943

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 53674

    Saida_Celula_AliquotaIPI = SUCESSO

    Exit Function

Erro_Saida_Celula_AliquotaIPI:

    Saida_Celula_AliquotaIPI = gErr

    Select Case gErr

        Case 53673, 53674 'Erro tratado na rotina chamada.
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164755)

    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da celula do grid que está deixando de ser a corrente /m

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        Select Case GridProdutos.Col

            Case iGrid_Observacao_Col

                lErro = Saida_Celula_Observacao(objGridInt)
                If lErro <> SUCESSO Then gError 53832

            Case iGrid_PrecoVista_Col

                lErro = Saida_Celula_PrecoVista(objGridInt)
                If lErro <> SUCESSO Then gError 53790

            Case iGrid_TotalVista_Col

                lErro = Saida_Celula_TotVista(objGridInt)
                If lErro <> SUCESSO Then gError 53791

            Case iGrid_PrecoPrazo_Col

                lErro = Saida_Celula_PrecoPrazo(objGridInt)
                If lErro <> SUCESSO Then gError 53792

            Case iGrid_TaxaFinanceira_Col

                lErro = Saida_Celula_TaxaFinanceira(objGridInt)
                If lErro <> SUCESSO Then gError 53793

            Case iGrid_PrazoEntrega_Col

                lErro = Saida_Celula_PrazoEntrega(objGridInt)
                If lErro <> SUCESSO Then gError 53794

            Case iGrid_Quantidade_Col
                lErro = Saida_Celula_Quantidade(objGridInt)
                If lErro <> SUCESSO Then gError 53795
            
            Case iGrid_QuantEntrega_Col

                lErro = Saida_Celula_QuantEntrega(objGridInt)
                If lErro <> SUCESSO Then gError 53795

            Case iGrid_AliquotaIPI_Col

                lErro = Saida_Celula_AliquotaIPI(objGridInt)
                If lErro <> SUCESSO Then gError 53796

            Case iGrid_AliquotaICMS_Col

                lErro = Saida_Celula_AliquotaICMS(objGridInt)
                If lErro <> SUCESSO Then gError 53797

            Case iGrid_TotalPrazo_Col

                lErro = Saida_Celula_TotPrazo(objGridInt)
                If lErro <> SUCESSO Then gError 53831

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 53798

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 53790, 53791, 53792, 53793, 53794, 53795, 53796, 53797, 53798, 53831, 53832, 86174
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164756)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PrecoPrazo(objGridInt As AdmGrid) As Long

Dim bPrecoAlterado As Boolean
Dim lErro As Long
Dim dPrecoPrazo As Double
Dim dTotPrazo As Double
Dim dTotVista As Double
Dim dQuantidade As Double
Dim dAliquotaIPI As Double
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim dTaxaFinanceira As Double
Dim dPrecoVista As Double
Dim iIndice As Integer
Dim iQuantIntervalos As Integer

On Error GoTo Erro_Saida_Celula_PrecoPrazo

    Set objGridInt.objControle = PrecoPrazo

    bPrecoAlterado = False

    If Len(Trim(PrecoPrazo.Text)) > 0 Then

        'Faz a critica do valor informado.
        lErro = Valor_Positivo_Critica(PrecoPrazo.Text)
        If lErro <> SUCESSO Then gError 53663

        'Coloca o valor na tela já Formatao.
        PrecoPrazo.Text = Format(PrecoPrazo.Text, gobjCOM.sFormatoPrecoUnitario) ' "STANDARD") 'Alterado por Wagner
            
        If Len(Trim(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoVista_Col))) > 0 Then
            
            'Se o PrecoPrazo é menor que o PrecoVista ==> erro
            If StrParaDbl(PrecoPrazo.Text) < (GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoVista_Col)) Then gError 76204
        
        End If
        
    Else
        
        'Limpa taxa Financeira
        GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TaxaFinanceira_Col) = ""

    End If

    'Se PrecoPrazo.text for diferente do Preço a Prazo no Grid faz bPrecoAlterado true
    If StrParaDbl(PrecoPrazo.Text) <> StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoPrazo_Col)) Then

        bPrecoAlterado = True
        gbPrecoGridAlterado = True

    End If
    
    'Chama Grid_Abandona_Celula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 53664

    If bPrecoAlterado = True Then
        
        'Calcula o Total a Prazo para o item
        dPrecoPrazo = StrParaDbl(PrecoPrazo.Text)
        dQuantidade = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Quantidade_Col))

        If Len(PrecoPrazo.Text) = 0 Or Len(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Quantidade_Col)) = 0 Then
            GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalPrazo_Col) = ""
        Else
            GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalPrazo_Col) = Format((dPrecoPrazo * dQuantidade), TotalPrazoRS.Format) 'Alterado por Wagner
        End If

        'Passa o que estiver escrito na combo para o objCondicaoPagto
        objCondicaoPagto.iCodigo = CondPagto_Extrai(CondicaoPagamento)

        'Se a Condição foi informada, lê a Condição de Pagamento
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 53771
        If lErro <> SUCESSO Then gError 53772

        dTotVista = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalVista_Col))
        dTotPrazo = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalPrazo_Col))

        'Atualiza o IPI a Vista do item
        dAliquotaIPI = PercentParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_AliquotaIPI_Col))
        GridProdutos.TextMatrix(GridProdutos.Row, iGrid_ValorIPIVista_Col) = dTotVista * dAliquotaIPI

        'Atualiza o IPI a Prazo do item
        GridProdutos.TextMatrix(GridProdutos.Row, iGrid_ValorIPIPrazo_Col) = Format((dTotPrazo * dAliquotaIPI), "Fixed")

        'Atualiza o Total a Prazo, a Vista do pedido e os valores do IPI do pedido
        lErro = Totais_Calcula()
        If lErro <> SUCESSO Then gError 53694

        dPrecoVista = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoVista_Col))
            
        'Calcula taxa Financeira
        If Len(Trim(PrecoPrazo.Text)) > 0 And Len(Trim(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoVista_Col))) > 0 Then
        
            dPrecoPrazo = StrParaDbl(PrecoPrazo.Text)

            If dPrecoPrazo <> dPrecoVista Then
                lErro = Calcula_TaxaFinanceira(dPrecoPrazo, dPrecoVista, dTaxaFinanceira, objCondicaoPagto.iCodigo)
                If lErro <> SUCESSO Then gError 76209
            End If
            'coloca a taxa financeira no grid
            GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TaxaFinanceira_Col) = Format(dTaxaFinanceira, "Percent")
        
        'Senão calcula preço a vista
        ElseIf Len(Trim(PrecoPrazo.Text)) > 0 And PercentParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TaxaFinanceira_Col)) > 0 Then
        
            dTaxaFinanceira = PercentParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TaxaFinanceira_Col))

            lErro = Calcula_PrecoVista_TaxaFinanceira(objCondicaoPagto, dPrecoVista, dTaxaFinanceira, dPrecoPrazo)
            If lErro <> SUCESSO Then gError 86285
        
            GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoVista_Col) = Format(dPrecoVista, gobjCOM.sFormatoPrecoUnitario) ' "STANDARD") 'Alterado por Wagner
            GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalVista_Col) = Format(dPrecoVista * dQuantidade, TotalVIstaRS.Format) 'Alterado por Wagner
        
        End If
        
        
    End If
   
    'Coloca o valor na tela já Formatao.
    PrecoPrazo.Text = Format(PrecoPrazo.Text, gobjCOM.sFormatoPrecoUnitario) ' "STANDARD") 'Alterado por Wagner

    If Len(Trim(TaxaConversao.Text)) > 0 Then Call ComparativoMoedaReal_Calcula(CDbl(TaxaConversao.Text), GridProdutos.Row)
        
    Saida_Celula_PrecoPrazo = SUCESSO

    Exit Function

Erro_Saida_Celula_PrecoPrazo:

    Saida_Celula_PrecoPrazo = gErr

    Select Case gErr

        Case 53771, 76209, 86285, 53772, 53663, 53664, 53665, 53694
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 76204
            Call Rotina_Erro(vbOKOnly, "ERRO_PRECOPRAZO_MENOR_PRECOVISTA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164757)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TaxaFinanceira(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim dTotVista As Double
Dim dTaxaFinanceira As Double
Dim dTotPrazo As Double
Dim dQuantidade As Double
Dim dAliquotaIPI As Double
Dim dPreco As Double
Dim dPrecoVista As Double
Dim dPrecoPrazo As Double

On Error GoTo Erro_Saida_Celula_TaxaFinanceira
    
    Set objGridInt.objControle = TaxaFinanceira

    dTaxaFinanceira = (PercentParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TaxaFinanceira_Col)))
        
    'Verifica se a Taxa Financeira está preenchida.
    If Len(Trim(TaxaFinanceira.Text)) > 0 Then

        'Critica o valor informado.
        lErro = Valor_NaoNegativo_Critica(TaxaFinanceira.Text)
        If lErro <> SUCESSO Then gError 53695

    End If
        
    'Chama a função Grid_Abandona_Celula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 53696

    dQuantidade = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Quantidade_Col))

    'Verifica se a Quantidade é maior que zero
    If dQuantidade > 0 Then
            
        'Verifica  se a Taxa Financeira foi alterada.
        If Abs((StrParaDbl(TaxaFinanceira.Text) / 100) - dTaxaFinanceira) > 0.00001 Then
        
            'Verifica se a Condição de Pagamento foi informada
            If Len(Trim(CondicaoPagamento.Text)) > 0 Then

                dTotVista = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalVista_Col))
                
                dPrecoVista = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoVista_Col))
                
                dTaxaFinanceira = PercentParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TaxaFinanceira_Col))

                If dPrecoVista > 0 Then

                    lErro = Calcula_PrecoPrazo(dPrecoPrazo, dPrecoVista, dTaxaFinanceira, CondPagto_Extrai(CondicaoPagamento))
                    If lErro <> SUCESSO Then gError 53699

                    If dPrecoPrazo > 0 Then
                        GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoPrazo_Col) = Format(dPrecoPrazo, gobjCOM.sFormatoPrecoUnitario) ' "STANDARD") 'Alterado por Wagner
                        GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalPrazo_Col) = Format(dPrecoPrazo * dQuantidade, TotalPrazoRS.Format) 'Alterado por Wagner
                    Else
                        GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoPrazo_Col) = ""
                        GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalPrazo_Col) = ""
                    End If
                Else
                    
                    
                    objCondicaoPagto.iCodigo = CondPagto_Extrai(CondicaoPagamento)
                    dPrecoPrazo = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoPrazo_Col))
                    
                    lErro = Calcula_PrecoVista_TaxaFinanceira(objCondicaoPagto, dPrecoVista, dTaxaFinanceira, dPrecoPrazo)
                    If lErro <> SUCESSO Then gError 53699

                    If dPrecoVista > 0 Then
                        GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoVista_Col) = Format(dPrecoVista, gobjCOM.sFormatoPrecoUnitario) ' "STANDARD") 'Alterado por Wagner
                        GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalVista_Col) = Format(dPrecoVista * dQuantidade, TotalVIstaRS.Format) 'Alterado por Wagner
                    Else
                    
                    End If
                End If
                
                dTotPrazo = dPrecoPrazo * dQuantidade

                If dTotPrazo > 0 Then
                    GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalPrazo_Col) = Format(dTotPrazo, TotalPrazoRS.Format) 'Alterado por Wagner
                Else
                    GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalPrazo_Col) = ""
                End If

                'Atualiza o IPI a Vista do item
                dAliquotaIPI = PercentParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_AliquotaIPI_Col))
                
                GridProdutos.TextMatrix(GridProdutos.Row, iGrid_ValorIPIVista_Col) = Format((dTotVista * dAliquotaIPI), "Standard")

                'Atualiza o IPI a Prazo do item
                GridProdutos.TextMatrix(GridProdutos.Row, iGrid_ValorIPIPrazo_Col) = Format((dTotPrazo * dAliquotaIPI), "Standard")

                'Atualiza o Total a Prazo, a Vista do pedido e os valores do IPI do pedido
                lErro = Totais_Calcula()
                If lErro <> SUCESSO Then gError 53700

            End If
        End If
    End If

    If Len(Trim(TaxaConversao.Text)) > 0 Then Call ComparativoMoedaReal_Calcula(CDbl(TaxaConversao.Text), GridProdutos.Row)
        
    Saida_Celula_TaxaFinanceira = SUCESSO

    Exit Function

Erro_Saida_Celula_TaxaFinanceira:

    Saida_Celula_TaxaFinanceira = gErr

    Select Case gErr

        Case 53695, 53696, 53698, 53699, 53700 'Erros tratados nas rotinas chamadas.
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 53697
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164758)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantEntrega(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantEntrega As Double

On Error GoTo Erro_Saida_Celula_QuantEntrega

    Set objGridInt.objControle = QuantEntrega

    'Verifica se QuantEntrega está preenchida
    If Len(Trim(QuantEntrega.Text)) > 0 Then

        'Critica o valor informado.
        lErro = Valor_NaoNegativo_Critica(QuantEntrega.Text)
        If lErro <> SUCESSO Then gError 53701

        'Converte QuantEntrega pasa Double
        dQuantEntrega = StrParaDbl(QuantEntrega.Text)

        'Coloca QuantEntrega na tela já formatada
        QuantEntrega.Text = Formata_Estoque(dQuantEntrega)

        'Verifica se QuantEntrega é menor ou igual a Quantidade
        'de Cotação.
        If StrParaDbl(QuantEntrega.Text) > StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Quantidade_Col)) Then gError 53677

    End If

    'Chama a função Grid_Abandona_Celula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 53678

    Saida_Celula_QuantEntrega = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantEntrega:

    Saida_Celula_QuantEntrega = gErr

    Select Case gErr

        Case 53677
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTENTREGA_MAIOR_QUANTCOTACAO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 53678, 53701 'Erros tratados nas rotinas chamadas.
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164759)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_AliquotaICMS(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_AliquotaICMS

    Set objGridInt.objControle = AliquotaICM

    'Verifica se a AlíquotaICMS está preenchida
    If Len(Trim(AliquotaICM.Text)) > 0 Then

        'Critica o valor informado
        lErro = Porcentagem_Critica(AliquotaICM.Text)
        If lErro <> SUCESSO Then gError 53575

        'Coloca em AliquotaICMS.Text a alíquota com o formato Fixed.
        AliquotaICM.Text = Format(AliquotaICM.Text, "Fixed")

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 53576

    Saida_Celula_AliquotaICMS = SUCESSO

    Exit Function

Erro_Saida_Celula_AliquotaICMS:

    Saida_Celula_AliquotaICMS = gErr

    Select Case gErr

        Case 53575, 53576 'Erros tratados nas rotinas chamadas
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164760)

    End Select

    Exit Function

End Function

Private Function Move_GridItens_Memoria(objPedidoCotacao As ClassPedidoCotacao) As Long
'Recolhe do Grid os dados do item pedido no parametro

Dim lErro As Long
Dim sProduto As String
Dim objItemPedCotacao As New ClassItemPedCotacao
Dim iPreenchido As Integer
Dim iIndice As Integer
Dim objItensCotacao As New ClassItemCotacao
Dim objItensCotacao1 As New ClassItemCotacao
Dim dValor As Double
Dim iContVista As Integer
Dim iContPrazo As Integer

On Error GoTo Erro_Move_GridItens_Memoria

    For iIndice = 1 To objGridItens.iLinhasExistentes

        Set objItemPedCotacao = New ClassItemPedCotacao

        'Verifica se o Produto está preenchido
        If Len(Trim(GridProdutos.TextMatrix(iIndice, iGrid_Produto_Col))) > 0 Then

            'Formata o produto
            lErro = CF("Produto_Formata", GridProdutos.TextMatrix(iIndice, iGrid_Produto_Col), sProduto, iPreenchido)
            If lErro <> SUCESSO Then gError 53773

            'Armazena o produto no obj.
            objItemPedCotacao.sProduto = sProduto

        End If
        
        'Armazena os dados do item
        objItemPedCotacao.sUM = GridProdutos.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
        objItemPedCotacao.dQuantidade = StrParaDbl(GridProdutos.TextMatrix(iIndice, iGrid_Quantidade_Col))
        objItemPedCotacao.iExclusivo = StrParaInt(GridProdutos.TextMatrix(iIndice, iGrid_Exclusivo_Col))
        objItemPedCotacao.sObservacao = GridProdutos.TextMatrix(iIndice, iGrid_Observacao_Col)

        Set objItensCotacao = New ClassItemCotacao

        'Armazena os dados dos itens.
        objItensCotacao.dQuantEntrega = StrParaDbl(GridProdutos.TextMatrix(iIndice, iGrid_QuantEntrega_Col))
        objItensCotacao.dAliquotaICMS = PercentParaDbl(GridProdutos.TextMatrix(iIndice, iGrid_AliquotaICMS_Col))
        objItensCotacao.iPrazoEntrega = StrParaInt(GridProdutos.TextMatrix(iIndice, iGrid_PrazoEntrega_Col))
        objItensCotacao.sObservacao = (GridProdutos.TextMatrix(iIndice, iGrid_Observacao_Col))
        objItensCotacao.dAliquotaIPI = PercentParaDbl(GridProdutos.TextMatrix(iIndice, iGrid_AliquotaIPI_Col))
        objItensCotacao.dtDataReferencia = DATA_NULA
        
        If giMoedaAnterior <> Codigo_Extrai(Moeda.Text) Then
            objItensCotacao.iMoeda = giMoedaAnterior
        Else
            objItensCotacao.iMoeda = Codigo_Extrai(Moeda.Text)
        End If
        
        objItensCotacao.dTaxa = StrParaDbl(TaxaConversao.Text)
        
        'Se a condição de pagamento não for a vista, mas se o preço a vista for preenchido o item possui dois registros: um a vista e o outro a prazo.
        If StrParaDbl(GridProdutos.TextMatrix(iIndice, iGrid_TotalVista_Col)) > 0 Then

            Set objItensCotacao1 = New ClassItemCotacao

            'Função que move os dados referentes a condição de pagamento a vista do objItensCotacao para o objItensCotacao1.
            Call Move_Dados(objItensCotacao, objItensCotacao1)

            iContVista = iContVista + 1
            objItensCotacao1.dValorIPI = StrParaDbl(GridProdutos.TextMatrix(iIndice, iGrid_ValorIPIVista_Col))
            objItensCotacao1.dPrecoUnitario = StrParaDbl(GridProdutos.TextMatrix(iIndice, iGrid_PrecoVista_Col))
            objItensCotacao1.dValorTotal = StrParaDbl(GridProdutos.TextMatrix(iIndice, iGrid_TotalVista_Col))
            objItensCotacao1.iCondPagto = COD_A_VISTA
            
            If giMoedaAnterior <> Codigo_Extrai(Moeda.Text) Then
                objItensCotacao.iMoeda = giMoedaAnterior
            Else
                objItensCotacao.iMoeda = Codigo_Extrai(Moeda.Text)
            End If
            
            objItensCotacao1.dTaxa = StrParaDbl(TaxaConversao.Text)
            objItensCotacao1.sObservacao = (GridProdutos.TextMatrix(iIndice, iGrid_Observacao_Col))

            objItemPedCotacao.colItensCotacao.Add objItensCotacao1
        
        End If
            
        If StrParaDbl(GridProdutos.TextMatrix(iIndice, iGrid_TotalPrazo_Col)) > 0 Then
            'Preenche os valores a prazo e usa o iContPrazo para saber quantos registros possuem condição de pagamento a prazo.
            iContPrazo = iContPrazo + 1
            objItensCotacao.dValorIPI = StrParaDbl(GridProdutos.TextMatrix(iIndice, iGrid_ValorIPIPrazo_Col))
            objItensCotacao.dPrecoUnitario = StrParaDbl(GridProdutos.TextMatrix(iIndice, iGrid_PrecoPrazo_Col))
            objItensCotacao.dValorTotal = StrParaDbl(GridProdutos.TextMatrix(iIndice, iGrid_TotalPrazo_Col))
            objItensCotacao.iCondPagto = CondPagto_Extrai(CondicaoPagamento)
            If giMoedaAnterior <> Codigo_Extrai(Moeda.Text) Then
                objItensCotacao.iMoeda = giMoedaAnterior
            Else
                objItensCotacao.iMoeda = Codigo_Extrai(Moeda.Text)
            End If
            objItensCotacao.dTaxa = StrParaDbl(TaxaConversao.Text)
            objItensCotacao.sObservacao = (GridProdutos.TextMatrix(iIndice, iGrid_Observacao_Col))

            'Adiciona o item na colecao de itens do itempedcotacao
            objItemPedCotacao.colItensCotacao.Add objItensCotacao

        End If

        'Adiciona o item na colecao de itens do pedido de cotacao
        objPedidoCotacao.colItens.Add objItemPedCotacao
    
    Next
    
    'Faz a validacao e pondera os Valores Total, Frete, Seguro, Despesas, Desconto e IPI
    'para as condicoes de pagto à vista e a prazo para os Itens de Cotacao
    lErro = Valida_Proporcao_Valores(objGridItens, iContVista, iContPrazo, objPedidoCotacao)
    If lErro <> SUCESSO Then gError 76487
    
    'Se a condição de pagamento estiver preenchida, verifica
    'o preço a vista dos itens para armazenar o STATUS.
    If Len(Trim(CondicaoPagamento.Text)) = 0 Then

        If iContVista = 0 Then

            objPedidoCotacao.iStatus = STATUS_GERADO_NAO_ATUALIZADO
        ElseIf iContVista = objGridItens.iLinhasExistentes Then
            objPedidoCotacao.iStatus = STATUS_ATUALIZADO
        Else
            objPedidoCotacao.iStatus = STATUS_PARCIALMENTE_ATUALIZADO
        End If

    Else

        If iContPrazo = objGridItens.iLinhasExistentes Then

            If iContVista = 0 Or iContVista = objGridItens.iLinhasExistentes Then
                objPedidoCotacao.iStatus = STATUS_ATUALIZADO
            Else
                objPedidoCotacao.iStatus = STATUS_PARCIALMENTE_ATUALIZADO
            End If
        Else
            objPedidoCotacao.iStatus = STATUS_PARCIALMENTE_ATUALIZADO
        End If

    End If

    Move_GridItens_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItens_Memoria:

    Move_GridItens_Memoria = gErr

    Select Case gErr

        Case 53773, 76487
            'Erro tratado na rotina chamada.

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164761)

    End Select

    Exit Function

End Function

Private Function Valida_Proporcao_Valores(objGridItens As AdmGrid, iContVista As Integer, iContPrazo As Integer, objPedidoCotacao As ClassPedidoCotacao) As Long
'Faz a validacao e pondera os Valores Total, Frete, Seguro, Despesas, Desconto e IPI
'para as condicoes de pagto à vista e a prazo para os Itens de Cotacao
    
Dim lErro As Long
Dim objItensCotacao As New ClassItemCotacao
Dim objItemPedCotacao As New ClassItemPedCotacao
Dim dValor As Double

On Error GoTo Erro_Valida_Proporcao_Valores

    Set objItensCotacao = New ClassItemCotacao
    Set objItemPedCotacao = New ClassItemPedCotacao
    
    If iContVista = objGridItens.iLinhasExistentes Then
        
        For Each objItemPedCotacao In objPedidoCotacao.colItens

            For Each objItensCotacao In objItemPedCotacao.colItensCotacao
                
                'Soma o ValorTotal dos Itens com condpagto à vista, se for a mesma moeda
                If objItensCotacao.iCondPagto = COD_A_VISTA And objItensCotacao.iMoeda = Codigo_Extrai(Moeda.Text) Then dValor = dValor + objItensCotacao.dValorTotal
                
            Next
        Next

        For Each objItemPedCotacao In objPedidoCotacao.colItens

            For Each objItensCotacao In objItemPedCotacao.colItensCotacao

                If objItensCotacao.iCondPagto = COD_A_VISTA And objItensCotacao.iMoeda = Codigo_Extrai(Moeda.Text) Then
                    
                    'Calcula os valores de Frete,Desconto,IPI,Despesa,Seguro proporcionalmente para ItemCotacao
                    lErro = Calcula_Proporcao_Valores(objItensCotacao, dValor)
                    If lErro <> SUCESSO Then gError 76432
                    
                End If
            Next
        Next
    End If

    dValor = 0

    If iContPrazo = objGridItens.iLinhasExistentes Then

        For Each objItemPedCotacao In objPedidoCotacao.colItens

            For Each objItensCotacao In objItemPedCotacao.colItensCotacao
                'soma o ValorTotal de cada ItemCotacao com condicaopagto diferente de codigo à vista
                If objItensCotacao.iCondPagto <> COD_A_VISTA And objItensCotacao.iMoeda = Codigo_Extrai(Moeda.Text) Then dValor = dValor + objItensCotacao.dValorTotal
                
            Next
        Next

        For Each objItemPedCotacao In objPedidoCotacao.colItens

            For Each objItensCotacao In objItemPedCotacao.colItensCotacao

                If objItensCotacao.iCondPagto <> COD_A_VISTA And objItensCotacao.iMoeda = Codigo_Extrai(Moeda.Text) Then
                    
                    'Calcula os valores proporcionalmente para cada ItemCotacao
                    lErro = Calcula_Proporcao_Valores(objItensCotacao, dValor)
                    If lErro <> SUCESSO Then gError 76433
                    
                End If
            Next
        Next

    End If

    Valida_Proporcao_Valores = SUCESSO
    
    Exit Function
    
Erro_Valida_Proporcao_Valores:

    Valida_Proporcao_Valores = gErr
    
    Select Case gErr
    
        Case 76432, 76433
            'Erros tratados nas rotinas chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164762)
    
    End Select
    
    Exit Function
        
End Function

Private Function Calcula_Proporcao_Valores(objItensCotacao As ClassItemCotacao, dValor As Double) As Long
'Calcula os valores de Desconto,Frete, IPI,Seguro e Despesas proporcionais a cada ItemCotacao

Dim lErro As Long
Dim dPercentual As Double
Dim dValorFrete As Double
Dim dValorSeguro As Double
Dim dValorDespesas As Double
Dim dValorDescontoVista As Double
Dim dValorDescontoPrazo As Double
Dim dValorIPIVista As Double
Dim dValorIPIPrazo As Double
Dim dValorTotalVista As Double
Dim dValorTotalPrazo As Double
Dim dOutrasDespesas As Double

On Error GoTo Erro_Calcula_Proporcao_Valores

    'Recolhe os valores da tela
    dValorFrete = StrParaDbl(ValorFrete.Text)
    dValorSeguro = StrParaDbl(ValorSeguro.Text)
    dValorDespesas = StrParaDbl(ValorDespesas.Text)
    dValorDescontoVista = StrParaDbl(DescontoVista.Text)
    dValorDescontoPrazo = StrParaDbl(DescontoPrazo.Text)
    dValorIPIVista = StrParaDbl(IPIValorVista.Caption)
    dValorIPIPrazo = StrParaDbl(IPIValorPrazo.Caption)
    dValorTotalVista = StrParaDbl(TotalVista.Caption)
    dValorTotalPrazo = StrParaDbl(TotalPrazo.Caption)
    dOutrasDespesas = StrParaDbl(ValorDespesas.Text)
        
    'Define o Percentual do ItemCotacao
    dPercentual = objItensCotacao.dValorTotal / dValor
    
    If objItensCotacao.iCondPagto = COD_A_VISTA Then
        objItensCotacao.dValorDesconto = dPercentual * dValorDescontoVista
'        objItensCotacao.dValorIPI = dPercentual * dValorIPIVista
    Else
        objItensCotacao.dValorDesconto = dPercentual * dValorDescontoPrazo
'        objItensCotacao.dValorIPI = dPercentual * dValorIPIPrazo
    End If
    
    'Define os valores para cada ItemCotacao
    objItensCotacao.dValorFrete = dPercentual * dValorFrete
    objItensCotacao.dValorSeguro = dPercentual * dValorSeguro
    objItensCotacao.dOutrasDespesas = dPercentual * dOutrasDespesas

    Calcula_Proporcao_Valores = SUCESSO
    
    Exit Function
    
Erro_Calcula_Proporcao_Valores:

    Calcula_Proporcao_Valores = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164763)
            
    End Select
    
    Exit Function
    
End Function

Private Function Move_PedidoCotacao_Memoria(objPedidoCotacao As ClassPedidoCotacao) As Long
'Função que chama as funções para mover a tela e os dados do grid para a memória.

Dim lErro As Long

On Error GoTo Erro_Move_PedidoCotacao_Memoria

    lErro = Move_Tela_Memoria(objPedidoCotacao)
    If lErro <> SUCESSO Then gError 53774

    Move_PedidoCotacao_Memoria = SUCESSO

    Exit Function

Erro_Move_PedidoCotacao_Memoria:

    Move_PedidoCotacao_Memoria = gErr

    Select Case gErr

        Case 53774 ', 53775
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164764)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objPedidoCotacao As New ClassPedidoCotacao

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se existe pedido na tela
    If Len(Trim(Codigo.Text)) = 0 Then gError 53776
    
    'Salva os valores da tela
     If Moeda.ListIndex > -1 Then Call SalvarGrid

    'Recolhe os dados da tela
    lErro = Move_PedidoCotacao_Memoria(gobjPedidoCotacao)
    If lErro <> SUCESSO Then gError 53778

    gobjPedidoCotacao.lCodigo = StrParaLong(Codigo.Text)
    gobjPedidoCotacao.iFilialEmpresa = giFilialEmpresa
    
    lErro = CF("PedidoCotacao_Atualiza", gobjPedidoCotacao)
    If lErro <> SUCESSO Then gError 53779

    'Limpa a Tela
    Call Limpa_Tela_PedidoCotacao

    'Fecha o sistema de setas
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0
    
    gbPrecoGridAlterado = False

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 53776
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOTACAO_NAO_SELECIONADO", gErr)

        Case 53778, 53779
            'Erros tratados nas rotinas chamadas

        Case 53780
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ITENS_PEDIDOCOTACAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164765)

    End Select
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

End Function

Private Function Move_Dados(objItensCotacao As ClassItemCotacao, objItensCotacao1 As ClassItemCotacao) As Long
'Função que move os dados independentes da condição de pagamento do objItensCotacao para o objItensCotacao1

    objItensCotacao1.dQuantEntrega = objItensCotacao.dQuantEntrega
    objItensCotacao1.dAliquotaICMS = objItensCotacao.dAliquotaICMS
    objItensCotacao1.iPrazoEntrega = objItensCotacao.iPrazoEntrega
    objItensCotacao1.sObservacao = objItensCotacao.sObservacao
    objItensCotacao1.dAliquotaIPI = objItensCotacao.dAliquotaIPI
    objItensCotacao1.dtDataReferencia = objItensCotacao.dtDataReferencia
    objItensCotacao1.iMoeda = objItensCotacao.iMoeda

End Function

Private Function Saida_Celula_TotVista(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim bTotalAlterado As Boolean, dTotVistaAnterior As Double
Dim dPrecoVista As Double
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim dTotVista As Double
Dim dTaxaFinanceira As Double
Dim dTotPrazo As Double
Dim dQuantidade As Double
Dim dAliquotaIPI As Double
Dim dPreco As Double
Dim dPrecoPrazo  As Double

On Error GoTo Erro_Saida_Celula_TotVista

    Set objGridInt.objControle = TotVista

    bTotalAlterado = False

    dQuantidade = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Quantidade_Col))
    dTotVistaAnterior = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalVista_Col))

    'Se o total a vista foi preenchido, critica o seu valor e coloca na tela.
    If Len(Trim(TotVista.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TotVista.Text)
        If lErro <> SUCESSO Then gError 53635

        TotVista.Text = Format(TotVista.Text, TotVista.Format) 'Alterado por Wagner

        dTotVista = StrParaDbl(TotVista.Text)

        If Len(Trim(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalPrazo_Col))) > 0 Then
            
            If StrParaDbl(TotVista.Text) > StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalPrazo_Col)) Then gError 76207
            
        End If
        
        'Verifica se o Total Vista foi alterado.
        If (dQuantidade > 0) And dTotVista <> StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalVista_Col)) Then

            dPreco = (dTotVista / dQuantidade)
            
            If dPreco > 0 Then
                GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoVista_Col) = Format(dPreco, gobjCOM.sFormatoPrecoUnitario) ' "STANDARD") 'Alterado por Wagner
            Else
                GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoVista_Col) = ""
            End If

        End If

    Else

        GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalVista_Col) = ""
        GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TaxaFinanceira_Col) = ""
        GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoVista_Col) = ""
        
    End If

    'Se o TotVista.text for diferente do total a vista no grid, faz
    'bPrecoAlterado=True
    If StrParaDbl(TotVista.Text) <> dTotVistaAnterior Then

        bTotalAlterado = True
        gbPrecoGridAlterado = True

    End If

    'Chama Grid_Abandona_Celula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 53636

    If bTotalAlterado = True Then

        'Calcula o Preço a Vista do item
        If (dTotVista / dQuantidade) > 0 Then
            GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoVista_Col) = Format((dTotVista / dQuantidade), gobjCOM.sFormatoPrecoUnitario) ' "STANDARD") 'Alterado por Wagner
        Else
            GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoVista_Col) = ""
        End If

        If ((dQuantidade > 0) And Len(Trim(CondicaoPagamento.Text)) > 0 And (dTotVista > 0)) Then

            'Passa o que estiver escrito na combo para o objCondicaoPagto
            objCondicaoPagto.iCodigo = CondPagto_Extrai(CondicaoPagamento)

            'Se a Condição foi informada, lê a Condição de Pagamento
            lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
            If lErro <> SUCESSO And lErro <> 19205 Then gError 53782
            If lErro <> SUCESSO Then gError 53783

            'Calcula o PrecoPrazo
            If Len(Trim(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalPrazo_Col))) = 0 Then
            
                dTotPrazo = dTotVista
                dTotPrazo = dTotPrazo * (1 + objCondicaoPagto.dAcrescimoFinanceiro)
                dPrecoPrazo = dTotPrazo / objCondicaoPagto.iNumeroParcelas
                dPrecoVista = dTotVista / dQuantidade
            
                GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalPrazo_Col) = Format(dTotPrazo, TotalPrazoRS.Format) 'Alterado por Wagner
            Else
                dTotPrazo = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalPrazo_Col))
                dPrecoPrazo = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalPrazo_Col)) / dQuantidade
                dPrecoVista = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalVista_Col)) / dQuantidade
            End If

            If dPrecoVista < 0.01 Then
                dPrecoVista = 0.01
            Else
                dPrecoVista = StrParaDbl(Format(dPrecoVista, "Fixed"))
                
                lErro = Calcula_TaxaFinanceira(dPrecoPrazo, dPrecoVista, dTaxaFinanceira, objCondicaoPagto.iCodigo)
                If lErro <> SUCESSO Then gError 76208
                
                GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TaxaFinanceira_Col) = Format(dTaxaFinanceira, "Percent")
            End If
            
            'De acordo com a quantidade calcula o preço a prazo e
            'coloca na tela.
            If (dTotPrazo / dQuantidade) > 0 Then
                GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoPrazo_Col) = Format((dTotPrazo / dQuantidade), PrecoPrazo.Format) 'Alterado por Wagner
            Else
                GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoPrazo_Col) = ""
            End If

            'Atualiza o IPI a vista do item.
            dAliquotaIPI = PercentParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_AliquotaIPI_Col))
            GridProdutos.TextMatrix(GridProdutos.Row, iGrid_ValorIPIVista_Col) = dTotVista * dAliquotaIPI

            'Atualiza o IPI a Prazo do item
            GridProdutos.TextMatrix(GridProdutos.Row, iGrid_ValorIPIPrazo_Col) = dTotPrazo * dAliquotaIPI

            'Atualiza o Total a Prazo, a Vista do pedido e os valores do IPI do pedido
            lErro = Totais_Calcula()
            If lErro <> SUCESSO Then gError 53821

        Else
        
            'Atualiza o IPI a vista do item.
            dAliquotaIPI = PercentParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_AliquotaIPI_Col))
            GridProdutos.TextMatrix(GridProdutos.Row, iGrid_ValorIPIVista_Col) = dTotVista * dAliquotaIPI

            'Atualiza o Total a Prazo, a Vista do pedido e os valores do IPI do pedido
            lErro = Totais_Calcula()
            If lErro <> SUCESSO Then gError 53821

        End If

    End If

    If Len(Trim(TaxaConversao.Text)) > 0 Then Call ComparativoMoedaReal_Calcula(CDbl(TaxaConversao.Text), GridProdutos.Row)

    Saida_Celula_TotVista = SUCESSO

    Exit Function

Erro_Saida_Celula_TotVista:

    Saida_Celula_TotVista = gErr

    Select Case gErr

        Case 53635, 53636, 53821 'Erros tratados nas rotinas chamadas
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 53782, 76208
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 53783 'Não encontrou condição de pagamento.
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 76207
            Call Rotina_Erro(vbOKOnly, "ERRO_PRECOPRAZO_MENOR_PRECOVISTA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164766)

    End Select

    Exit Function

End Function

Private Function Move_Pedido_Memoria(colPedidoCompras As Collection, objPedidoCotacao As ClassPedidoCotacao, iCondPagto As Integer, ByVal iMoeda As Integer, ByVal dTaxa As Double) As Long
'Função que move os dados da tela para colPedidoCompra.

Dim lErro As Long
Dim objItemPedCotacao As New ClassItemPedCotacao
Dim iIndice As Integer
Dim objPedidoCompra As New ClassPedidoCompras
Dim objReqCompras As New ClassRequisicaoCompras
Dim objItemRC As New ClassItemReqCompras
Dim colPedCompraExclu As New Collection
Dim colPedCompraGeral As New Collection
Dim bAchou  As Boolean
Dim iIndice1 As Integer
Dim objItemPC  As New ClassItemPedCompra
Dim colReqCompras As New Collection

On Error GoTo Erro_Move_Pedido_Memoria
        
'    'Verifica se a Quantidade do ItemPedCotacao é igual a soma da
'    'QuantidadeCotar de CotacaoProdutoItemRC
'    lErro = CF("Valida_Quantidade_ItemPedCotacao", objPedidoCotacao)
'    If lErro <> SUCESSO Then gError 76146

    For Each objItemPedCotacao In objPedidoCotacao.colItens
                
        Set colReqCompras = New Collection
        
        'Lê as Requisicoes de Compra associadas ao Item do Pedido de Cotacao
        'Janaina
        lErro = CF("ItemPedCotacao_Le_RequisicaoCompra", objItemPedCotacao, colReqCompras, objPedidoCotacao)
        'Janaina
        If lErro <> SUCESSO Then gError 76149
        
        For Each objReqCompras In colReqCompras
            
            For Each objItemRC In objReqCompras.colItens
            
                'Verifica se o Item é exclusivo
                If objItemRC.iExclusivo = MARCADO Then
                    Set colPedidoCompras = colPedCompraExclu
                Else
                    Set colPedidoCompras = colPedCompraGeral
                End If
                
                iIndice1 = 0
                bAchou = False
                    
                For Each objPedidoCompra In colPedidoCompras
                        
                    iIndice1 = iIndice1 + 1
                    'Verifica se já existe um Pedido de Compra com o mesmo TipoDestino
                    If (objPedidoCompra.iTipoDestino = objReqCompras.iTipoDestino) And (objPedidoCompra.lFornCliDestino = objReqCompras.lFornCliDestino) And (objPedidoCompra.iFilialDestino = objReqCompras.iFilialDestino) Then
                        bAchou = True
                        Exit For
                    End If
                    
                Next
                
                'Se já existe Pedido de compra com o mesmo destino
                If bAchou Then
                    'seleciona o pedido
                    Set objPedidoCompra = colPedidoCompras(iIndice1)
                    
                'Senão houver pedido de compra com o mesmo destino
                Else
                    'Cria um novo Pedido de compras
                    Set objPedidoCompra = New ClassPedidoCompras
                                                            
                    'Cria um  novo Pedido de Compra
                    lErro = PedidoCompra_Cria(objPedidoCompra, objItemRC, objPedidoCotacao, objReqCompras, objItemPedCotacao, iCondPagto, iMoeda, dTaxa)
                    If lErro <> SUCESSO Then gError 76245
                    
                    'Adiciona o Pedido de Compra na colecao de PedCompra
                    colPedidoCompras.Add objPedidoCompra
                    
                End If
    
                'Adiciona  ItemPedCompra no Pedido de Compra ou atualiza um ItemPedCompra já existente
                lErro = PedidoCompra_Adiciona_ItemPedCompra(objPedidoCompra, objItemRC, objItemPedCotacao, iCondPagto)
                If lErro <> SUCESSO Then gError 76244
                
            Next
        Next
    Next
    
    lErro = Atualiza_Valores_Pedido(colPedidoCompras, objPedidoCotacao.colItens)
    If lErro <> SUCESSO Then gError 178588
    
    Set colPedidoCompras = New Collection

    'Gera uma única colecao de Pedidos de Compra, a partir das colecoes colPedCompraExclu e colPedCompraGeral já criadas
    lErro = PedidoCompra_Define_Colecao(colPedCompraExclu, colPedCompraGeral, colPedidoCompras)
    If lErro <> SUCESSO Then gError 76246
    
    Move_Pedido_Memoria = SUCESSO

    Exit Function

Erro_Move_Pedido_Memoria:

    Move_Pedido_Memoria = gErr

    Select Case gErr
    
        Case 76146, 76149, 76244, 76245, 76246, 178588
            'Erros tratados nas rotinas chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164767)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TotPrazo(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim bTotalAlterado As Boolean
Dim dPrecoVista As Double
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim dTotPrazo As Double
Dim dTaxaFinanceira As Double
Dim dQuantidade As Double
Dim dAliquotaIPI As Double
Dim dPreco As Double
Dim dPrecoPrazo  As Double

On Error GoTo Erro_Saida_Celula_TotPrazo

    Set objGridInt.objControle = TotPrazo

    bTotalAlterado = False

    dQuantidade = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Quantidade_Col))

    'Se o total a prazo for preenchido faz a critica do seu valor e coloca na tela.
    If Len(Trim(TotPrazo.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TotPrazo.Text)
        If lErro <> SUCESSO Then gError 53801

        TotPrazo.Text = Format(TotPrazo.Text, TotPrazo.Format) 'Alterado por Wagner

        dTotPrazo = StrParaDbl(TotPrazo.Text)

        If Len(Trim(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalVista_Col))) > 0 Then
            If StrParaDbl(TotPrazo.Text) < StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalVista_Col)) Then gError 76206
        End If
        
        'Verifica se o Total foi alterado.
        If (dQuantidade > 0) And dTotPrazo <> StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalPrazo_Col)) Then

            If (dTotPrazo / dQuantidade) > 0 Then
                GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoPrazo_Col) = Format((dTotPrazo / dQuantidade), gobjCOM.sFormatoPrecoUnitario) ' "STANDARD") 'Alterado por Wagner
            Else
                GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoPrazo_Col) = ""
            End If

        End If

    Else

        GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoPrazo_Col) = ""
        
        'Limpa taxa Financeira
        GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TaxaFinanceira_Col) = ""
    
    End If
    
'    If dQuantidade > 0 Then
'
'        dPreco = StrParaDbl(TotPrazo.Text) / StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Quantidade_Col))
'
'        If dPreco > 0 Then
'            GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoPrazo_Col) = Format(dPreco, gobjCOM.sFormatoPrecoUnitario) ' "fixed") 'Alterado por Wagner
'        Else
'            GridProdutos.TextMatrix(GridProdutos.Row, iGrid_PrecoPrazo_Col) = ""
'        End If
'
'    End If
    
    'Se o preco.text for diferente do preço a vista no grid, faz
    'bPrecoAlterado=True
    If StrParaDbl(TotPrazo.Text) <> StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalPrazo_Col)) Then

        bTotalAlterado = True
        gbPrecoGridAlterado = True
    
    End If
        
    'Chama Grid_Abandona_Celula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 53799

    If bTotalAlterado = True Then

        'Passa o que estiver escrito na combo para o objCondicaoPagto
        objCondicaoPagto.iCodigo = CondPagto_Extrai(CondicaoPagamento)

        'Se a Condição foi informada, lê a Condição de Pagamento
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 76218
        If lErro <> SUCESSO Then gError 76219

        'Calcula taxa Financeira
        If Len(Trim(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalPrazo_Col))) > 0 And Len(Trim(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalVista_Col))) > 0 Then
        
            dPrecoPrazo = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalPrazo_Col)) / dQuantidade
            dPrecoVista = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TotalVista_Col)) / dQuantidade
            
            If dPrecoPrazo <> dPrecoVista And dPrecoVista > 0 Then
                dPrecoPrazo = Round(dPrecoPrazo, 2)
                lErro = Calcula_TaxaFinanceira(dPrecoPrazo, dPrecoVista, dTaxaFinanceira, objCondicaoPagto.iCodigo)
                If lErro <> SUCESSO Then gError 76209
            End If
            
            'coloca a taxa financeira no grid
            GridProdutos.TextMatrix(GridProdutos.Row, iGrid_TaxaFinanceira_Col) = Format(dTaxaFinanceira, "Percent")
            
        End If
        
        'Atualiza o IPI a Prazo do item
        dAliquotaIPI = PercentParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_AliquotaIPI_Col))
        GridProdutos.TextMatrix(GridProdutos.Row, iGrid_ValorIPIPrazo_Col) = Format((dTotPrazo * dAliquotaIPI), "Fixed")

        'Chama a função Totais_Calcula que calcula os totais do pedido.
        lErro = Totais_Calcula()
        If lErro <> SUCESSO Then gError 53800

    End If

    Saida_Celula_TotPrazo = SUCESSO

    Exit Function

Erro_Saida_Celula_TotPrazo:

    Saida_Celula_TotPrazo = gErr

    Select Case gErr

        Case 53799, 53800, 53801, 76218 'Erros tratados nas rotinas chamadas
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 76206
            Call Rotina_Erro(vbOKOnly, "ERRO_PRECOPRAZO_MENOR_PRECOVISTA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 76219
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164768)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Pedido de Cotação"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "PedidoCotacao"
    
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


Private Sub Label15_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label15(Index), Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15(Index), Button, Shift, X, Y)
End Sub


Private Sub FornecDestinoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecDestinoLabel, Source, X, Y)
End Sub

Private Sub FornecDestinoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecDestinoLabel, Button, Shift, X, Y)
End Sub

Private Sub FornecedorLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecedorLabel, Source, X, Y)
End Sub

Private Sub FornecedorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel, Button, Shift, X, Y)
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
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
    'Alterado por Wagner, se tiver uma parcela só paga a vista também deve considerar mesmo se estiver marcado iMensal=1
    If ((objCondicaoPagto.iMensal = 0 And objCondicaoPagto.iIntervaloParcelas = 0) Or objCondicaoPagto.iNumeroParcelas = 1) And objCondicaoPagto.iDiasParaPrimeiraParcela = 0 Then
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
    
    '        'Calcula o ValorPresente para o limite superior do intervalo da taxa financeira
    '        lErro = Calcula_ValorPresente(dTaxaFim, dPrecoVista, iCondPagto, dValorPresente1)
    '        If lErro <> SUCESSO Then gError 76211
    
    '        'Calcula o ValorPresente para o limite inferior do intervalo da taxa financeira
    '        lErro = Calcula_ValorPresente(dTaxaIni, dPrecoVista, iCondPagto, dValorPresente2)
    '        If lErro <> SUCESSO Then gError 76210
    
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
    
    '            lErro = Calcula_ValorPresente(dTaxaMeio, dPrecoVista, iCondPagto, dValorPresente3)
    '            If lErro <> SUCESSO Then gError 76212
    
                lErro = Calcula_PrecoPrazo_TaxaFinanceira(objCondicaoPagto, dPrecoVista, dTaxaMeio, dValorPresente3)
                If lErro <> SUCESSO Then gError 76212
    
    
    '            dValorPresente3 = Round(dValorPresente3, 2)
    
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164769)

    End Select

    Exit Function

End Function

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164770)

    End Select

    Exit Function

End Function

'Function CalculaDias_CondPagamento(objCondPagto As ClassCondicaoPagto, iDias As Integer) As Long
''devolve o numero de dias médio em que o pagamento será feito para a condição passada como parametro
'
'Dim iDiaParcelaAtual As Integer
'Dim iTotalDias As Integer
'Dim iNumeroParcelas As Integer
'Dim iPeso As Integer
'Dim iIntervalo As Integer
'
'On Error GoTo Erro_CalculaDias_CondPagamento
'
'    'Se a condição de pagamento for à vista
'    If objCondPagto.iCodigo = COD_A_VISTA Or objCondPagto.iCodigo = 0 Then
'
'        iDias = 0
'
'    Else
'
'        'Se a condição de pagamento for Mensal
'        If objCondPagto.iMensal = 1 Then
'            iIntervalo = 30
'        Else
'            iIntervalo = objCondPagto.iIntervaloParcelas
'        End If
'
'        'Guarda o número de parcelas tirando a primeira
'        iNumeroParcelas = objCondPagto.iNumeroParcelas - 1
'
'        'Calcula total de dias da condição de pagamento
'        Do While iNumeroParcelas >= 0
'
'            'Calcula o número de dias que faltam para chegar a parcela em questão
'            iDiaParcelaAtual = objCondPagto.iDiasParaPrimeiraParcela + (iIntervalo * iNumeroParcelas)
'
'            'Acumula o número de dias de todas as parcelas
'            iTotalDias = iTotalDias + iDiaParcelaAtual
'
'            'Decrementa o número de parcelas
'            iNumeroParcelas = iNumeroParcelas - 1
'
'        Loop
'
'        'Calcula a média ponderada de dias
'        iDias = iTotalDias / objCondPagto.iNumeroParcelas
'
'
'    End If
'
'    CalculaDias_CondPagamento = SUCESSO
'
'    Exit Function
'
'Erro_CalculaDias_CondPagamento:
'
'    CalculaDias_CondPagamento = gErr
'
'    Select Case gErr
'
'         Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164771)
'
'   End Select
'
'    Exit Function
'
'End Function

'Function Calcula_TaxaFinanceira(ByVal dPrecoPrazo As Double, ByVal dPrecoVista As Double, dTaxaFinanceira As Double, ByVal iCondPagto As Integer) As Long
''Faz o cálculo da taxa financeira, a partir do PrecoPrazo, PrecoVista e CondicaoPagto
'
'Dim lErro As Long
'Dim objCondicaoPagto As New ClassCondicaoPagto
'Dim iDias As Integer
'Dim dTaxaFinanceiraDiaria As Double
'
'On Error GoTo Erro_Calcula_TaxaFinanceira
'
'    objCondicaoPagto.iCodigo = iCondPagto
'
'    'Lê a condicao de Pagamento
'    lErro = CF("CondicaoPagto_Le",objCondicaoPagto)
'    If lErro <> SUCESSO And lErro <> 19205 Then gError 83831
'
'    If lErro <> SUCESSO Then
'        dTaxaFinanceira = 0
'    End If
'
'    'devolve o numero de dias médio em que o pagamento será feito para a condição passada como parametro
'    lErro = CalculaDias_CondPagamento(objCondicaoPagto, iDias)
'    If lErro <> SUCESSO Then gError 83832
'
'    If iDias = 0 Then
'        dTaxaFinanceiraDiaria = 0
'    Else
'        dTaxaFinanceiraDiaria = ((dPrecoPrazo / dPrecoVista) ^ (1 / iDias)) - 1
'    End If
'
'    dTaxaFinanceira = (1 + dTaxaFinanceiraDiaria) ^ 30 - 1
'
'    Calcula_TaxaFinanceira = SUCESSO
'
'    Exit Function
'
'Erro_Calcula_TaxaFinanceira:
'
'    Calcula_TaxaFinanceira = gErr
'
'    Select Case gErr
'
'        Case 83831, 83832
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164772)
'
'    End Select
'
'    Exit Function
'
'End Function

'Function Calcula_PrecoPrazo(dPrecoPrazo As Double, ByVal dPrecoVista As Double, ByVal dTaxaFinanceira As Double, ByVal iCondPagto As Integer) As Long
''Faz o cálculo do preço a prazo a partir da taxa financeira, PrecoVista e CondicaoPagto
'
'Dim lErro As Long
'Dim objCondicaoPagto As New ClassCondicaoPagto
'Dim iDias As Integer
'Dim dTaxaFinanceiraDiaria As Double
'
'On Error GoTo Erro_Calcula_PrecoPrazo
'
'    objCondicaoPagto.iCodigo = iCondPagto
'
'    'Lê a condicao de Pagamento
'    lErro = CF("CondicaoPagto_Le",objCondicaoPagto)
'    If lErro <> SUCESSO And lErro <> 19205 Then gError 83834
'
'    If lErro <> SUCESSO Then
'        dPrecoPrazo = 0
'    End If
'
'    'devolve o numero de dias médio em que o pagamento será feito para a condição passada como parametro
'    lErro = CalculaDias_CondPagamento(objCondicaoPagto, iDias)
'    If lErro <> SUCESSO Then gError 83836
'
'    dTaxaFinanceiraDiaria = (1 + dTaxaFinanceira) ^ (1 / 30) - 1
'
'    dPrecoPrazo = dPrecoVista * ((1 + dTaxaFinanceiraDiaria) ^ iDias)
'
'    Calcula_PrecoPrazo = SUCESSO
'
'    Exit Function
'
'Erro_Calcula_PrecoPrazo:
'
'    Calcula_PrecoPrazo = gErr
'
'    Select Case gErr
'
'        Case 83834, 83836
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164773)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Calcula_PrecoVista(ByVal dPrecoPrazo As Double, dPrecoVista As Double, ByVal dTaxaFinanceira As Double, ByVal iCondPagto As Integer) As Long
''Faz o cálculo do preço a vista a partir da taxa financeira, PrecoVista e CondicaoPagto
'
'Dim lErro As Long
'Dim objCondicaoPagto As New ClassCondicaoPagto
'Dim iDias As Integer
'Dim dTaxaFinanceiraDiaria As Double
'
'On Error GoTo Erro_Calcula_PrecoVista
'
'    objCondicaoPagto.iCodigo = iCondPagto
'
'    'Lê a condicao de Pagamento
'    lErro = CF("CondicaoPagto_Le",objCondicaoPagto)
'    If lErro <> SUCESSO And lErro <> 19205 Then gError 83837
'
'    If lErro <> SUCESSO Then
'        dPrecoVista = 0
'    End If
'
'    'devolve o numero de dias médio em que o pagamento será feito para a condição passada como parametro
'    lErro = CalculaDias_CondPagamento(objCondicaoPagto, iDias)
'    If lErro <> SUCESSO Then gError 83839
'
'    dTaxaFinanceiraDiaria = dTaxaFinanceira ^ (1 / 30)
'
'    If ((1 + dTaxaFinanceiraDiaria) ^ iDias) = 0 Then
'        dPrecoVista = 0
'    Else
'        dPrecoVista = dPrecoPrazo / ((1 + dTaxaFinanceiraDiaria) ^ iDias)
'    End If
'
'    Calcula_PrecoVista = SUCESSO
'
'    Exit Function
'
'Erro_Calcula_PrecoVista:
'
'    Calcula_PrecoVista = gErr
'
'    Select Case gErr
'
'        Case 83837, 83839
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164774)
'
'    End Select
'
'    Exit Function
'
'End Function

Function Calcula_PrecoPrazo(dPrecoPrazo As Double, ByVal dPrecoVista As Double, ByVal dTaxaFinanceira As Double, ByVal iCondPagto As Integer) As Long
'Faz o cálculo do preço a prazo a partir da taxa financeira, PrecoVista e CondicaoPagto

Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim iDias As Integer
Dim dTaxaFinanceiraDiaria As Double

On Error GoTo Erro_Calcula_PrecoPrazo

    objCondicaoPagto.iCodigo = iCondPagto

    'Lê a condicao de Pagamento
    lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
    If lErro <> SUCESSO And lErro <> 19205 Then gError 83841

    If lErro <> SUCESSO Then
        dPrecoPrazo = 0
    Else

        'Calcula o Preco a Prazo a partir da taxa financeira informada
        lErro = Calcula_PrecoPrazo_TaxaFinanceira(objCondicaoPagto, dPrecoVista, dTaxaFinanceira, dPrecoPrazo)
        If lErro <> SUCESSO Then gError 83842

    End If

    Calcula_PrecoPrazo = SUCESSO

    Exit Function

Erro_Calcula_PrecoPrazo:

    Calcula_PrecoPrazo = gErr

    Select Case gErr

        Case 83841, 83842

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164775)

    End Select

    Exit Function

End Function


Function PedidoCompra_Cria(objPedidoCompra As ClassPedidoCompras, objItemRC As ClassItemReqCompras, objPedidoCotacao As ClassPedidoCotacao, objReqCompras As ClassRequisicaoCompras, objItemPedCotacao As ClassItemPedCotacao, iCondPagto As Integer, ByVal iMoeda As Integer, ByVal dTaxa As Double) As Long
'Cria um novo Pedido de Compra

Dim lErro As Long
Dim objItemPC As New ClassItemPedCompra
Dim objUsuario As New ClassUsuario
Dim objComprador As New ClassComprador

On Error GoTo Erro_PedidoCompra_Cria
    
    objPedidoCompra.dtData = gdtDataAtual
    objPedidoCompra.dtDataEmissao = DATA_NULA
    objPedidoCompra.dtDataBaixa = DATA_NULA
    objPedidoCompra.dtDataAlteracao = DATA_NULA
    objPedidoCompra.dtDataEnvio = DATA_NULA
    objPedidoCompra.iFilial = objPedidoCotacao.iFilial
    objPedidoCompra.iFilialEmpresa = objPedidoCotacao.iFilialEmpresa
    objPedidoCompra.lFornecedor = objPedidoCotacao.lFornecedor
    objPedidoCompra.sContato = objPedidoCotacao.sContato
    objPedidoCompra.iTipoDestino = objReqCompras.iTipoDestino
    objPedidoCompra.iFilialDestino = objReqCompras.iFilialDestino
    objPedidoCompra.lFornCliDestino = objReqCompras.lFornCliDestino
    objPedidoCompra.iMoeda = iMoeda
    If iMoeda <> MOEDA_REAL Then objPedidoCompra.dTaxa = dTaxa
    
    objUsuario.sNomeReduzido = Comprador.Caption
    'Lê o usuario
    lErro = CF("Usuario_Le_NomeRed", objUsuario)
    If lErro <> SUCESSO And lErro <> 57269 Then gError 72500
    If lErro = 57269 Then gError 72502
    
    objComprador.sCodUsuario = objUsuario.sCodUsuario
    objComprador.iFilialEmpresa = giFilialEmpresa
    
    'Lê o Comprador
    lErro = CF("Comprador_Le_Usuario", objComprador)
    If lErro <> SUCESSO And lErro <> 50059 Then gError 72501
    If lErro = 50059 Then gError 72503
    
    objPedidoCompra.iComprador = objComprador.iCodigo
    objPedidoCompra.sTipoFrete = objPedidoCotacao.iTipoFrete
    objPedidoCompra.dOutrasDespesas = StrParaDbl(ValorDespesas.Text)
    objPedidoCompra.dValorFrete = StrParaDbl(ValorFrete.Text)
    objPedidoCompra.dValorSeguro = StrParaDbl(ValorSeguro.Text)
    
    'Se a condição de pagamento for a vista, armazena os valores a vista.
    objPedidoCompra.iCondicaoPagto = iCondPagto
    
    If iCondPagto = COD_A_VISTA Then
    
        objPedidoCompra.dValorDesconto = StrParaDbl(DescontoVista.Text)
        objPedidoCompra.dValorIPI = StrParaDbl(IPIValorVista.Caption)
        objPedidoCompra.dValorTotal = StrParaDbl(TotalVista.Caption)
    
    'Senão, armazena os valores a prazo.
    Else
    
        objPedidoCompra.dValorDesconto = StrParaDbl(DescontoPrazo.Text)
        objPedidoCompra.dValorIPI = StrParaDbl(IPIValorPrazo.Caption)
        objPedidoCompra.dValorTotal = StrParaDbl(TotalPrazo.Caption)
    
    End If
    
    objPedidoCompra.dValorProdutos = objPedidoCompra.dValorTotal - objPedidoCompra.dValorFrete - objPedidoCompra.dValorSeguro - objPedidoCompra.dOutrasDespesas - objPedidoCompra.dValorIPI + objPedidoCompra.dValorDesconto
        
    PedidoCompra_Cria = SUCESSO
    
    Exit Function
    
Erro_PedidoCompra_Cria:
    
    PedidoCompra_Cria = gErr
    
    Select Case gErr
    
        Case 72500, 72501
            'Erros tratados nas rotinas chamadas
            
        Case 72502
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO2", gErr, objUsuario.sNomeReduzido)
                
        Case 72503
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPRADOR_NAO_CADASTRADO1", gErr, objComprador.sCodUsuario)
                        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164776)
            
    End Select
    
    Exit Function
    
    
End Function

Private Function PedidoCompra_Adiciona_ItemPedCompra(objPedidoCompra As ClassPedidoCompras, objItemRC As ClassItemReqCompras, objItemPedCotacao As ClassItemPedCotacao, iCondPagto As Integer) As Long
'Adiciona  ItemPedCompra no Pedido de Compra ou atualiza um ItemPedCompra já existente

Dim lErro As Long
Dim objItemPC As New ClassItemPedCompra
Dim bAchouItem As Boolean
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim dQuantidade As Double
Dim dDesconto As Double
Dim sProduto As String
Dim iPreenchido As Integer
Dim objItem As New ClassItemCotacao
Dim objItemRCAux As ClassItemReqCompras

On Error GoTo Erro_PedidoCompra_Adiciona_ItemPedCompra

    iIndice2 = 0
    bAchouItem = False
    
    'Verifica se há um Item de Pedido de Compra com o mesmo produto do ItemRC
    For Each objItemPC In objPedidoCompra.colItens
    
        iIndice2 = iIndice2 + 1
        
        'Se existe um Item de Pedido de compra com o Produto igual ao item de Requisicao de Compra
        If objItemPC.sProduto = objItemRC.sProduto Then
        
            objProduto.sCodigo = objItemRC.sProduto

            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 76186

            'Se não encontrou o produto ==> erro
            If lErro = 28030 Then gError 76191

            'Converte a UM do Produto de UMCompra para UM do ItemReqCompra
            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemPC.sUM, objItemRC.sUM, dFator)
            If lErro <> SUCESSO Then gError 76187

            'Pega a quantidade do Item de Pedido de Compras
            dQuantidade = objItemPC.dQuantidade
            
            'Subtrai da quantidade quanto já alocado em outros itens de requisições
            For Each objItemRCAux In objItemPC.colItemReqCompras
                dQuantidade = dQuantidade - objItemRCAux.dQuantNoPedido
            Next
            
            'Converte para a U.M. do Item de Requisição em questão.
            dQuantidade = dQuantidade * dFator
            
            objItemRC.dQuantComprar = IIf(objItemRC.dQuantNaCotacao - dQuantidade < 0, objItemRC.dQuantNaCotacao, dQuantidade)
            
            objItemRC.dQuantNoPedido = objItemRC.dQuantComprar / dFator
            
            objItemPC.colItemReqCompras.Add objItemRC
            objItemPC.lNumIntOrigem = objItemPedCotacao.lNumIntDoc
        
            bAchouItem = True
            Exit For
        
        End If
        
    Next
    
    'Se não encontrou Item com o mesmo produto
    If bAchouItem = False Then
    
        'cria um novo item para o pedido de compras
        Set objItemPC = New ClassItemPedCompra
    
        objItemPC.sProduto = objItemRC.sProduto
        
        For iIndice = 1 To objGridItens.iLinhasExistentes
        
            'Verifica se o Produto está preenchido
            If Len(Trim(GridProdutos.TextMatrix(iIndice, iGrid_Produto_Col))) > 0 Then
        
                'Formata o produto
                lErro = CF("Produto_Formata", GridProdutos.TextMatrix(iIndice, iGrid_Produto_Col), sProduto, iPreenchido)
                If lErro <> SUCESSO Then gError 53702
        
            End If
            
            If objItemPC.sProduto = sProduto And StrParaInt(GridProdutos.TextMatrix(iIndice, iGrid_Exclusivo_Col)) = objItemRC.iExclusivo Then
                'Move os dados do grid para o objItemPC
                objItemPC.dQuantidade = objItemPedCotacao.dQuantidade
'                objItemPC.dQuantidade = objItemRC.dQuantNaCotacao
                objItemPC.dAliquotaICMS = PercentParaDbl(GridProdutos.TextMatrix(iIndice, iGrid_AliquotaICMS_Col))
                objItemPC.sObservacao = GridProdutos.TextMatrix(iIndice, iGrid_Observacao_Col)
                objItemPC.dAliquotaIPI = PercentParaDbl(GridProdutos.TextMatrix(iIndice, iGrid_AliquotaIPI_Col))
                objItemPC.sDescProduto = GridProdutos.TextMatrix(iIndice, iGrid_DescProduto_Col)
                objItemPC.iStatus = ITEM_PED_COMPRAS_ABERTO
                objItemPC.sUM = GridProdutos.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
                objItemPC.iTipoOrigem = TIPO_ITEMPEDCOTACAO
                objItemPC.dtDataLimite = DATA_NULA
                objItemPC.lNumIntOrigem = objItemPedCotacao.lNumIntDoc
            
                For Each objItem In objItemPedCotacao.colItensCotacao
                
                    If objItem.iCondPagto = iCondPagto Then Exit For
                    
                Next
                
                If iCondPagto = COD_A_VISTA Then
                    objItemPC.dPrecoUnitario = StrParaDbl(GridProdutos.TextMatrix(iIndice, iGrid_PrecoVista_Col))
                    objItemPC.dValorIPI = (objItemPC.dAliquotaIPI) * (objItemPC.dQuantidade * objItemPC.dPrecoUnitario)
                    
                    objItemPC.dValorDesconto = objItem.dValorDesconto + dDesconto
                    
                Else
                    objItemPC.dPrecoUnitario = StrParaDbl(GridProdutos.TextMatrix(iIndice, iGrid_PrecoPrazo_Col))
                    objItemPC.dValorIPI = objItemPC.dAliquotaIPI * (objItemPC.dQuantidade * objItemPC.dPrecoUnitario)
                    
                    objItemPC.dValorDesconto = objItem.dValorDesconto
                    
                    
                End If
        
                objProduto.sCodigo = objItemRC.sProduto
        
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 76215
        
                'Se não encontrou o produto ==> erro
                If lErro = 28030 Then gError 76217
        
                objItemPC.dPercentMaisReceb = objProduto.dPercentMaisReceb
                objItemPC.dPercentMenosReceb = objProduto.dPercentMenosReceb
        
                If objProduto.iTemFaixaReceb = 1 Then
                    objItemPC.iRebebForaFaixa = 0
                Else
                    objItemPC.iRebebForaFaixa = objProduto.iRecebForaFaixa + 1
                End If
        
                'Converte a UM do Produto de UMCompra para UM do ItemReqCompra
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemPC.sUM, objItemRC.sUM, dFator)
                If lErro <> SUCESSO Then gError 76216

                dQuantidade = objItemPC.dQuantidade * dFator
    
                objItemRC.dQuantComprar = IIf(objItemRC.dQuantNaCotacao - dQuantidade < 0, objItemRC.dQuantNaCotacao, dQuantidade)
        
                objItemRC.dQuantNoPedido = objItemRC.dQuantComprar / dFator
                
                'Adiciona na colecao de ItemReqCompras
                objItemPC.colItemReqCompras.Add objItemRC
        
                'Adiciona na colecao de ItensPedCompra
                objPedidoCompra.colItens.Add objItemPC
                
           End If
           
        Next
   
    End If
    
    PedidoCompra_Adiciona_ItemPedCompra = SUCESSO
    
    Exit Function
    
Erro_PedidoCompra_Adiciona_ItemPedCompra:

    PedidoCompra_Adiciona_ItemPedCompra = gErr
    
    Select Case gErr
    
        Case 53702, 76186, 76187, 76215, 76216
            'Erro tratado na rotina chamada
            
        Case 76191, 76217
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", objProduto.sCodigo, gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164777)
            
    End Select
    
    Exit Function
    
End Function

Function PedidoCompra_Define_Colecao(colPedCompraExclu As Collection, colPedCompraGeral As Collection, colPedidoCompras As Collection) As Long
'A partir das colecoes de Pedidos de Compra Exclusivos e de Pedidos de Compra Não Exclusivos,
'define uma coleção única para todos os Pedidos de Compra criados

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim bProdutoIgual As Boolean
Dim objPCGeral As New ClassPedidoCompras
Dim objPCExclu As New ClassPedidoCompras
Dim objItemPCExclu As New ClassItemPedCompra
Dim objItemPCGeral As New ClassItemPedCompra
Dim objPedidoCompra As New ClassPedidoCompras

On Error GoTo Erro_PedidoCompra_Define_Colecao

    'Verifica se existem Pedidos de Compra nas duas colecoes criadas
    If colPedCompraExclu.Count > 0 And colPedCompraGeral.Count > 0 Then
    
        bProdutoIgual = False
        For iIndice = colPedCompraExclu.Count To 1 Step -1
        
            Set objPCExclu = colPedCompraExclu.Item(iIndice)
            For Each objPCGeral In colPedCompraGeral
            
                'Verifica se os Pedidos tem o mesmo TipoDestino
                If objPCExclu.iTipoDestino = objPCGeral.iTipoDestino And objPCExclu.iFilialDestino = objPCGeral.iFilialDestino And objPCExclu.lFornCliDestino = objPCGeral.lFornCliDestino Then
                
                    For iIndice2 = objPCExclu.colItens.Count To 1 Step -1
                        
                        Set objItemPCExclu = objPCExclu.colItens.Item(iIndice2)
                        
                        For Each objItemPCGeral In objPCGeral.colItens
                         
                            'Verifica se o produto do Item Exclusivo está presente na colecao de Itens nao exclusivos
                            If objItemPCExclu.sProduto = objItemPCGeral.sProduto Then
                                bProdutoIgual = True
                                Exit For
                            End If
                        Next
                    Next
                    'Se nao encontrou produto igual nas colecoes de Itens pesquisadas
                    If bProdutoIgual = False Then
                        
                        For iIndice2 = objPCExclu.colItens.Count To 1 Step -1
                            'Adiciona o item exclusivo na colecao de itens nao exclusivos
                            objPCGeral.colItens.Add objPCExclu.colItens.Item(iIndice2)
                            'Remove o Item
                            objPCExclu.colItens.Remove (iIndice2)
                        Next
                        
                        'Remove o Pedido
                        colPedCompraExclu.Remove (iIndice)
                        
                    End If
                End If
            Next
        Next
    End If
    
    'Coloca todos os pedidos em uma única coleção
    For Each objPedidoCompra In colPedCompraExclu
        colPedidoCompras.Add objPedidoCompra
    Next
    For Each objPedidoCompra In colPedCompraGeral
        colPedidoCompras.Add objPedidoCompra
    Next
    
    PedidoCompra_Define_Colecao = SUCESSO
    
    Exit Function
    
Erro_PedidoCompra_Define_Colecao:

    PedidoCompra_Define_Colecao = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164778)
            
    End Select
    
    Exit Function
    
End Function

Function Carrega_Moeda() As Long

Dim lErro As Long
Dim objMoeda As ClassMoedas
Dim colMoedas As New Collection
Dim iPosMoedaReal As Integer
Dim iIndice As Integer

On Error GoTo Erro_Carrega_Moeda
    
    lErro = CF("Moedas_Le_Todas", colMoedas)
    If lErro <> SUCESSO Then gError 103371
    
    'se não existem moedas cadastradas
    If colMoedas.Count = 0 Then gError 103372
    
    For Each objMoeda In colMoedas
    
        Moeda.AddItem objMoeda.iCodigo & SEPARADOR & objMoeda.sNome
        If objMoeda.iCodigo = MOEDA_REAL Then iPosMoedaReal = iIndice
        Moeda.ItemData(Moeda.NewIndex) = objMoeda.iCodigo
        
        iIndice = iIndice + 1
    
    Next
    
    Moeda.ListIndex = iPosMoedaReal

    Carrega_Moeda = SUCESSO
    
    Exit Function
    
Erro_Carrega_Moeda:

    Carrega_Moeda = gErr
    
    Select Case gErr
    
        Case 103371
        
        Case 103372
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDAS_NAO_CADASTRADAS", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164779)
    
    End Select

End Function

Private Sub BotaoTrazCotacao_Click()
'Traz a última cotação da moeda selecionada

Dim lErro As Long
Dim iLinha As Integer
Dim objCotacao As New ClassCotacaoMoeda
Dim objCotacaoAnterior As New ClassCotacaoMoeda

On Error GoTo Erro_BotaoTrazCotacao_Click

    'Carrega objCotacao
    objCotacao.dtData = gdtDataAtual
    
    'Se a moeda não foi selecionada => Erro
    If Len(Trim(Moeda.Text)) = 0 Then gError 108943
        
    'Preeche com a Moeda selecionada
    objCotacao.iMoeda = Codigo_Extrai(Moeda.List(Moeda.ListIndex))
    objCotacaoAnterior.iMoeda = Codigo_Extrai(Moeda.List(Moeda.ListIndex))

    'Chama função de leitura
    lErro = CF("CotacaoMoeda_Le_UltimasCotacoes", objCotacao, objCotacaoAnterior)
    If lErro <> SUCESSO Then gError 108944
    
    'Se nao existe cotacao para a data informada => Mostra a última. Se mesmo assim nao existir => Colocar 1,00
    TaxaConversao.Text = IIf(objCotacaoAnterior.dValor = 0, 1, IIf(objCotacao.dValor <> 0, Format(objCotacao.dValor, "#.0000"), Format(objCotacaoAnterior.dValor, "#.0000")))
    
    Call TaxaConversao_Validate(bSGECancelDummy)
    
    Exit Sub
    
Erro_BotaoTrazCotacao_Click:

    Select Case gErr
    
        Case 108943
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDA_NAO_PREENCHIDA", gErr)
            
        Case 108944
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164780)
    
    End Select
    
End Sub

Private Sub ComparativoMoedaReal_Calcula(ByVal dTaxa As Double, ByVal iLinha As Integer)
'Preenche as colunas INFORMATIVAS de proporção da moeda R$.

    If iLinha > 0 Then

        'Preço A vista em R$
        GridProdutos.TextMatrix(iLinha, iGrid_TotalVista_RS_Col) = Format(StrParaDbl(GridProdutos.TextMatrix(iLinha, iGrid_TotalVista_Col)) * dTaxa, TotalVIstaRS.Format) 'Alterado por Wagner
        
        'Preço A prazo em R$
        GridProdutos.TextMatrix(iLinha, iGrid_TotalPrazo_RS_Col) = Format(StrParaDbl(GridProdutos.TextMatrix(iLinha, iGrid_TotalPrazo_Col)) * dTaxa, TotalPrazoRS.Format) 'Alterado por Wagner
    End If
    
    Exit Sub

End Sub

Private Function RemoveCotacaoGlobal(ByVal iMoeda As Integer, ByVal bAvisaExclusao As Boolean) As Long

Dim iIndice As Integer
Dim iIndice2 As Integer
Dim objItemPedCotacao As New ClassItemPedCotacao
Dim objItensCotacao As New ClassItemCotacao

On Error GoTo Erro_RemoveCotacaoGlobal

    If bAvisaExclusao = True Then
        
        If Rotina_Aviso(vbYesNo, "AVISO_REMOVE_COTACAO_MOEDA") = vbNo Then gError 108966
        
    End If
    
    'Busca na colecao global a moeda para remocao
    For Each objItemPedCotacao In gobjPedidoCotacao.colItens
     
        iIndice = iIndice + 1
        
            For Each objItensCotacao In objItemPedCotacao.colItensCotacao
            
                iIndice2 = iIndice2 + 1
            
                'Se a moeda for a mesma => Remove da colecao Global
                If objItensCotacao.iMoeda = Codigo_Extrai(Moeda.List(Moeda.ListIndex)) Then
                
                        objItemPedCotacao.colItensCotacao.Remove (iIndice2)
                        iIndice2 = iIndice2 - 1
                    
                End If
                
            Next
            
            iIndice2 = 0
               
    Next
    
    RemoveCotacaoGlobal = SUCESSO
    
    Exit Function
    
Erro_RemoveCotacaoGlobal:

    Select Case gErr
    
        Case 108966
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164781)

    End Select

End Function

Private Function AdicionaCotacaoGlobal() As Long
'Guarda na colecao global o atual grid

Dim lErro As Long
Dim iIndice As Integer
Dim objPedidoCotacao As New ClassPedidoCotacao
Dim objItemPedCotacaoGlobal As New ClassItemPedCotacao
Dim objItensCotacao As New ClassItemCotacao

Dim objItemPedCotacaoAux As New ClassItemPedCotacao

On Error GoTo Erro_AdicionaCotacaoGlobal

    'Move os dados do grid para o obj
    lErro = Move_GridItens_Memoria(objPedidoCotacao)
    If lErro <> SUCESSO Then gError 108967
    
    If gbPrecoGridAlterado = False Then gError 108971
        
    'Para cada Item do obj Global
    For Each objItemPedCotacaoGlobal In gobjPedidoCotacao.colItens
        
        'Para cada Item recolhido do Grid
        For Each objItemPedCotacaoAux In objPedidoCotacao.colItens
        
            'Se encontrou o item na coleção global
            If objItemPedCotacaoGlobal.sProduto = objItemPedCotacaoAux.sProduto Then
                
                objItemPedCotacaoGlobal.dQuantidade = objItemPedCotacaoAux.dQuantidade
                objItemPedCotacaoGlobal.iExclusivo = objItemPedCotacaoAux.iExclusivo
                
                'Para cada Item recolhido no Grid
                For Each objItensCotacao In objItemPedCotacaoAux.colItensCotacao
                
                    'Inclui a nova informação no obj Global
                    objItemPedCotacaoGlobal.colItensCotacao.Add objItensCotacao
                
                Next
                
            End If
        Next
    Next
                        
    'Coloca os valores de frete seguro e despesas na tela
    ValorFrete.Text = ""
    ValorSeguro.Text = ""
    ValorDespesas.Text = ""
    DescontoVista.Text = ""
    DescontoPrazo.Text = ""
    
    AdicionaCotacaoGlobal = SUCESSO
    
    Exit Function

Erro_AdicionaCotacaoGlobal:
    
    AdicionaCotacaoGlobal = gErr
    
    Select Case gErr
    
        Case 108967, 108968, 108971
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164782)
            
    End Select

End Function

Private Sub Determina_Status_PedCompra()
'Determina se o pc está ou nao atualizado.

Dim iIndiceMoeda As Integer
Dim objItemPedCotacao As New ClassItemPedCotacao
Dim objItensCotacao As New ClassItemCotacao
Dim iContCotacao As Integer
Dim iContProdutos As Integer

On Error GoTo Erro_Determina_Status_PedCompra

    gobjPedidoCotacao.iStatus = STATUS_GERADO_NAO_ATUALIZADO

    'Para cada moeda utilizada ...
    For iIndiceMoeda = 1 To gColMoedasUsadas.Count
        
        'Para cada produto ...
        For Each objItemPedCotacao In gobjPedidoCotacao.colItens
            
            'Guarda a quantidade de produtos pesquisados.
            iContProdutos = iContProdutos + 1
            
            'Verifica as n cotacoes atribuidas ...
            For Each objItensCotacao In objItemPedCotacao.colItensCotacao
            
                'Se a moeda da cotacao for a pesquisada e o produto tem preco atribuido ...
                If objItensCotacao.iMoeda = Codigo_Extrai(gColMoedasUsadas.Item(iIndiceMoeda)) And objItensCotacao.dValorTotal <> 0 Then
                    
                    'Atualiza o contador ...
                    iContCotacao = iContCotacao + 1
                    Exit For
                    
                End If
                
            Next
        
        Next
        
        'Se o numero de cotacoes com preco na moeda em questao para esse produto for igual a quantidade de produtos _
        Logo, existe para essa moeda, precos em todos os produtos
        If iContCotacao = iContProdutos Then
        
            'Atualiza o Status
            gobjPedidoCotacao.iStatus = STATUS_ATUALIZADO
            Exit For
            
        Else
        
            If iContCotacao > 0 Then gobjPedidoCotacao.iStatus = STATUS_PARCIALMENTE_ATUALIZADO
            
            'Zera os contadores para o próximo produto
            iContProdutos = 0
            iContCotacao = 0
        
        End If
            
    Next
    
    Exit Sub
    
Erro_Determina_Status_PedCompra:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164783)
    
    End Select

End Sub

Private Sub Determina_Moedas_Cotadas(colMoedasCotadas As Collection)
'Determina se o pc está ou nao atualizado.

Dim iIndiceMoeda As Integer
Dim iIndice As Integer
Dim objItemPedCotacao As New ClassItemPedCotacao
Dim objItensCotacao As New ClassItemCotacao
Dim iContCotacao As Integer
Dim iContProdutos As Integer

On Error GoTo Erro_Determina_Moedas_Cotadas

    'Para cada moeda utilizada ...
    For iIndiceMoeda = 1 To gColMoedasUsadas.Count
        
        'Para cada produto ...
        For Each objItemPedCotacao In gobjPedidoCotacao.colItens
            
            'Guarda a quantidade de produtos pesquisados.
            iContProdutos = iContProdutos + 1
            
            'Verifica as n cotacoes atribuidas ...
            For Each objItensCotacao In objItemPedCotacao.colItensCotacao
            
                'Se a moeda da cotacao for a pesquisada e o produto tem preco atribuido ...
                If objItensCotacao.iMoeda = Codigo_Extrai(gColMoedasUsadas.Item(iIndiceMoeda)) And objItensCotacao.dValorTotal <> 0 Then
                    
                    'Atualiza o contador ...
                    iContCotacao = iContCotacao + 1
                    Exit For
                    
                End If
                
            Next
        
        Next
        
        'Se o numero de cotacoes com preco na moeda em questao para esse produto for igual a quantidade de produtos _
        Logo, existe para essa moeda, precos em todos os produtos
        If iContCotacao = iContProdutos Then
        
            colMoedasCotadas.Add gColMoedasUsadas.Item(iIndiceMoeda)
            
        Else
            
            'Zera os contadores para o próximo produto
            iContProdutos = 0
            iContCotacao = 0
        
        End If
            
    Next
    
    Exit Sub
    
Erro_Determina_Moedas_Cotadas:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164784)
    
    End Select

    Exit Sub

End Sub

Private Sub Determina_Moedas_Usadas()

Dim iIndice As Integer
Dim iIndice2 As Integer
Dim objItemPedCotacao As New ClassItemPedCotacao
Dim objItensCotacao As New ClassItemCotacao
Dim bAchou As Boolean

On Error GoTo Erro_Determina_Moedas_Usadas

    'Para cada produto ...
    For Each objItemPedCotacao In gobjPedidoCotacao.colItens
        
        iIndice = iIndice + 1
            
            'Verifica as n cotacoes atribuidas ...
            For Each objItensCotacao In objItemPedCotacao.colItensCotacao
            
                bAchou = False
                
                'E caso ainda nao esteja na colecao de moedas usadas => Adiciona
                If gColMoedasUsadas.Count > 0 Then
                
                    'Para cada item já existente na colecao, verifica a existencia
                    For iIndice2 = 1 To gColMoedasUsadas.Count
                        If Codigo_Extrai(gColMoedasUsadas.Item(iIndice2)) = objItensCotacao.iMoeda Then
                            bAchou = True
                            Exit For
                        End If
                    Next
                    
                End If
                    
                If bAchou = False Then
                    
                    iIndice2 = 0
                
                    Do While Codigo_Extrai(Moeda.List(iIndice2)) <> objItensCotacao.iMoeda
                        iIndice2 = iIndice2 + 1
                    Loop
                    
                    If Codigo_Extrai(Moeda.List(iIndice2)) = objItensCotacao.iMoeda Then
                        gColMoedasUsadas.Add Moeda.List(iIndice2)
                    End If
                    
                End If
                
            Next
            
    Next

    Exit Sub
    
Erro_Determina_Moedas_Usadas:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164785)
    
    End Select

End Sub

Public Sub SalvarGrid()
        
Dim lErro As Long

On Error GoTo Erro_SalvarGrid
        
    gbPrecoGridAlterado = True
    
    'Verifica se já existe na colecao. Se existir => Exclui e armazena a nova.
    lErro = RemoveCotacaoGlobal(giMoedaAnterior, False)
    If lErro <> SUCESSO Then gError 108965
            
    'Guarda o Grid na Colecao
    lErro = AdicionaCotacaoGlobal()
    If lErro <> SUCESSO Then gError 108965
    
    gbPrecoGridAlterado = False

    Exit Sub
    
Erro_SalvarGrid:

    Select Case gErr
    
        Case 108965
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164786)
    
    End Select
    
    Exit Sub
    
End Sub

Function Calcula_PrecoVista_TaxaFinanceira(objCondPagto As ClassCondicaoPagto, dPrecoVista As Double, dTaxaFinanceira As Double, dPrecoPrazo As Double) As Long
'Calcula o Preco a Vista a partir da taxa financeira e do Preço a prazo


Dim lErro As Long
Dim dPrecoIni As Double
Dim dPrecoFim As Double
Dim dPrecoMeio As Double
Dim dValorPresente1 As Double
Dim dValorPresente2 As Double
Dim dValorPresente3 As Double
Dim dValor As Double
Dim bAchou As Boolean
Dim dValorPrazoIni As Double
Dim dValorPrazoFim As Double

On Error GoTo Erro_Calcula_PrecoVista_TaxaFinanceira

    'Lê a condicao de Pagamento
    lErro = CF("CondicaoPagto_Le", objCondPagto)
    If lErro <> SUCESSO And lErro <> 19205 Then gError 83841
    If lErro <> SUCESSO Then gError 86284

    'Define os intervalos onde preco vista será tolerado (para menos e para mais)
    dValorPrazoIni = dPrecoPrazo - 0.001
    dValorPrazoFim = dPrecoPrazo + 0.001

    'Define o intervalo do Preço a vista
    dPrecoIni = 0
    dPrecoFim = dPrecoPrazo

    bAchou = False

    'Enquanto o Preco Vista nao estiver no intervalo definido
    Do While Not bAchou

        lErro = Calcula_PrecoPrazo_TaxaFinanceira(objCondPagto, dPrecoIni, dTaxaFinanceira, dValorPresente1)
        If lErro <> SUCESSO Then gError 76211

        lErro = Calcula_PrecoPrazo_TaxaFinanceira(objCondPagto, dPrecoFim, dTaxaFinanceira, dValorPresente2)
        If lErro <> SUCESSO Then gError 76210

        If dPrecoPrazo = dValorPresente1 Then
            dPrecoVista = dPrecoIni
            Exit Do
        ElseIf dPrecoPrazo = dValorPresente2 Then
            dPrecoVista = dPrecoFim
            Exit Do
        End If

        'Define o preço Intermediário
        dPrecoMeio = ((dPrecoIni + dPrecoFim) / 2)

        'Se o preço real está entre os dois calculados
        If (dValorPresente1 < dPrecoPrazo) And (dPrecoPrazo < dValorPresente2) Then

            'Tenta com o preco mediano
            lErro = Calcula_PrecoPrazo_TaxaFinanceira(objCondPagto, dPrecoMeio, dTaxaFinanceira, dValorPresente3)
            If lErro <> SUCESSO Then gError 76212


            If (dValorPrazoFim >= dValorPresente3) And (dValorPrazoIni <= dValorPresente3) Then
                bAchou = True
                dPrecoVista = dPrecoMeio

            ElseIf (dValorPrazoFim < dValorPresente3) Then

                'altera a taxa inicial
                dPrecoFim = dPrecoMeio

            Else

                dPrecoIni = dPrecoMeio

            End If

        Else

            dPrecoIni = dPrecoFim
            dPrecoFim = dPrecoFim + 100

        End If

    Loop

    Calcula_PrecoVista_TaxaFinanceira = SUCESSO

    Exit Function


Erro_Calcula_PrecoVista_TaxaFinanceira:

    Calcula_PrecoVista_TaxaFinanceira = gErr

    Select Case gErr

        Case 76210 To 76214, 83303, 86284
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164787)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantidade As Double

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    'Verifica se Quantidade está preenchida
    If Len(Trim(Quantidade.Text)) > 0 Then

        'Critica o valor informado.
        lErro = Valor_NaoNegativo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 53701

        'Converte Quantidade pasa Double
        dQuantidade = StrParaDbl(Quantidade.Text)

        'Coloca Quantidade na tela já formatada
        Quantidade.Text = Formata_Estoque(dQuantidade)

    End If

    'Chama a função Grid_Abandona_Celula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 53678

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 53678, 53701 'Erros tratados nas rotinas chamadas.
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164788)

    End Select

    Exit Function

End Function

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

'##############################################
'Inserido por Wagner
Private Sub Formata_Controles()

    PrecoPrazo.Format = gobjCOM.sFormatoPrecoUnitario
    PrecoVista.Format = gobjCOM.sFormatoPrecoUnitario

End Sub
'##############################################

Function Atualiza_Valores_Pedido(colPedidoCompras As Collection, colItensPedCotacao As Collection) As Long
'Aproveita os valores das cotações utilizadas
'caso o pedido tenha sido gerado com itens da mesma cotação
         
Dim lErro As Long
Dim objItemPC As ClassItemPedCompra
Dim objItemPedCotacao As ClassItemPedCotacao
Dim objCotItemConc As ClassCotacaoItemConc
Dim objPedidoCompra As ClassPedidoCompras
Dim objPedidoCotacao As New ClassPedidoCotacao
Dim objItemConcorrencia As ClassItemConcorrencia
Dim objItensCotacao As ClassItemCotacao
    
    
On Error GoTo Erro_Atualiza_Valores_Pedido

    'Atualiza o valor dos produtos no pedido de venda
    For Each objPedidoCompra In colPedidoCompras

        'Zera os acumuladores dos valores
        objPedidoCompra.dValorDesconto = 0
        objPedidoCompra.dValorFrete = 0
        objPedidoCompra.dValorIPI = 0
        objPedidoCompra.dValorProdutos = 0
        objPedidoCompra.dValorSeguro = 0
        objPedidoCompra.dOutrasDespesas = 0

'        'Se o pedido foi gerado com itens de um só ped Cotação
'        If objPedidoCompra.lPedCotacao <> 0 Then
'
'            objPedidoCotacao.lCodigo = objPedidoCompra.lPedCotacao
'            objPedidoCotacao.iFilialEmpresa = giFilialEmpresa
'
'            'Lê o Pedido de Cotacao
'            lErro = CF("PedidoCotacao_Le", objPedidoCotacao)
'            If lErro <> SUCESSO And lErro <> 53670 Then gError 62728
'            If lErro <> SUCESSO Then gError 62729 'Não encontrou
'
'            objPedidoCompra.sTipoFrete = objPedidoCotacao.iTipoFrete
            
            'Para cada item de pedido de compra
        For Each objItemPC In objPedidoCompra.colItens
                
'                'Busca nos itens de concorrencia os dados do item de cotação
'                For Each objItemConcorrencia In gcolItemConcorrencia
'
'                    For Each objCotItemConc In objItemConcorrencia.colCotacaoItemConc
                        
'                        'Se a cotação foi a utilizada pelo item de Pedido de Compras
'                        If objItemPC.lNumIntOrigem = objCotItemConc.lNumIntDoc Then
            For Each objItemPedCotacao In colItensPedCotacao


                If objItemPC.lNumIntOrigem = objItemPedCotacao.lNumIntDoc Then

                    For Each objItensCotacao In objItemPedCotacao.colItensCotacao
                
                        'Soma o ValorTotal dos Itens com condpagto à vista, se for a mesma moeda
                        If objItensCotacao.iCondPagto = objPedidoCompra.iCondicaoPagto Then



'                            'Guarda o número do item de cotação
'                            Set objItemCotacao = colItensCotacao(CStr(objCotItemConc.lItemCotacao))
                                                         
                            objPedidoCompra.dOutrasDespesas = objPedidoCompra.dOutrasDespesas + (objItensCotacao.dOutrasDespesas * (objItemPC.dQuantidade * objItemPC.dPrecoUnitario) / (objItensCotacao.dValorTotal))
                            objPedidoCompra.dValorDesconto = objPedidoCompra.dValorDesconto + (objItensCotacao.dValorDesconto * (objItemPC.dQuantidade * objItemPC.dPrecoUnitario) / (objItensCotacao.dValorTotal))
                            objPedidoCompra.dValorFrete = objPedidoCompra.dValorFrete + (objItensCotacao.dValorFrete * (objItemPC.dQuantidade * objItemPC.dPrecoUnitario) / (objItensCotacao.dValorTotal))
                            objPedidoCompra.dValorSeguro = objPedidoCompra.dValorSeguro + (objItensCotacao.dValorSeguro * (objItemPC.dQuantidade * objItemPC.dPrecoUnitario) / (objItensCotacao.dValorTotal))
                            objItemPC.dAliquotaICMS = objItensCotacao.dAliquotaICMS
                            objItemPC.dAliquotaIPI = objItensCotacao.dAliquotaIPI
                            objItemPC.dValorIPI = (objItensCotacao.dValorIPI * (objItemPC.dQuantidade * objItemPC.dPrecoUnitario) / (objItensCotacao.dValorTotal))
                            objPedidoCompra.dValorIPI = objPedidoCompra.dValorIPI + objItemPC.dValorIPI
                            
                        End If
                        
                    Next
                End If
            Next
        Next
        
        'Atualiza o valor dos produtos no Pedido de compras
        For Each objItemPC In objPedidoCompra.colItens
            objPedidoCompra.dValorProdutos = objPedidoCompra.dValorProdutos + (objItemPC.dPrecoUnitario * objItemPC.dQuantidade)
        Next
        
        objPedidoCompra.dValorTotal = objPedidoCompra.dValorFrete + objPedidoCompra.dValorIPI + objPedidoCompra.dValorProdutos + objPedidoCompra.dValorSeguro - objPedidoCompra.dValorDesconto
    Next
    
    Atualiza_Valores_Pedido = SUCESSO
    
    Exit Function
    
Erro_Atualiza_Valores_Pedido:

    Atualiza_Valores_Pedido = gErr
    
    Select Case gErr
    
        Case 62728
    
        Case 62729
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOTACAO_NAO_ENCONTRADO", gErr, objPedidoCotacao.lCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161268)
            
    End Select
    
    Exit Function

End Function

Private Sub GridRequisicoes_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridRequisicoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRequisicoes, iAlterado)
    End If

End Sub

Private Sub GridRequisicoes_EnterCell()

    Call Grid_Entrada_Celula(objGridRequisicoes, iAlterado)

End Sub

Private Sub GridRequisicoes_GotFocus()

    Call Grid_Recebe_Foco(objGridRequisicoes)

End Sub

Private Sub GridRequisicoes_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridRequisicoes, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRequisicoes, iAlterado)
    End If

End Sub

Private Sub GridRequisicoes_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridRequisicoes)

End Sub

Private Sub GridRequisicoes_LeaveCell()

    Call Saida_Celula(objGridRequisicoes)

End Sub

Private Sub GridRequisicoes_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridRequisicoes)

End Sub

Private Sub GridRequisicoes_RowColChange()

    Call Grid_RowColChange(objGridRequisicoes)

End Sub

Private Sub GridRequisicoes_Scroll()

    Call Grid_Scroll(objGridRequisicoes)

End Sub

Private Function Inicializa_Grid_Req(objGridInt As AdmGrid) As Long

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Filial Req.")
    objGridInt.colColuna.Add ("Requisição")
    objGridInt.colColuna.Add ("Pedido Venda")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (FilialReq.Name)
    objGridInt.colCampo.Add (CodigoReq.Name)
    objGridInt.colCampo.Add (CodPV.Name)
    
    'Colunas do Grid
    iGrid_FilialReq_Col = 1
    iGrid_CodigoReq_Col = 2
    iGrid_CodPV_Col = 3

    'Grid do GridInterno
    objGridInt.objGrid = GridRequisicoes

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_REQUISICOES + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 30

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Req = SUCESSO

    Exit Function

End Function

Function Preenche_GridRequisicoes() As Long

Dim lErro As Long
Dim colReq As New Collection
Dim objReq As ClassRequisicaoCompras
Dim lCodigoPV As Long
Dim iFilialPV As Integer
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim objItemOP As ClassItemOP
Dim iLinha As Integer

On Error GoTo Erro_Preenche_GridRequisicoes

    If gobjPedidoCotacao.lCotacao > 0 Then

    'Limpa o Grid de Requisições
    Call Grid_Limpa(objGridRequisicoes)

    lErro = CF("Cotacao_Le_Req_Associadas", gobjPedidoCotacao.lCotacao, gobjPedidoCotacao.lFornecedor, gobjPedidoCotacao.iFilial, colReq)
    If lErro <> SUCESSO Then gError 178867
    
    For Each objReq In colReq
    
        iLinha = iLinha + 1
        
        GridRequisicoes.TextMatrix(iLinha, iGrid_FilialReq_Col) = objReq.iFilialEmpresa
        GridRequisicoes.TextMatrix(iLinha, iGrid_CodigoReq_Col) = objReq.lCodigo
    
        If Len(Trim(objReq.sOPCodigo)) <> 0 Then
        
            objOrdemProducao.iFilialEmpresa = objReq.iFilialEmpresa
            objOrdemProducao.sCodigo = objReq.sOPCodigo
        
            lErro = CF("ItensOrdemProducao_Le", objOrdemProducao)
            If lErro <> SUCESSO And lErro <> 30401 Then gError 178868
    
            If lErro <> SUCESSO Then
            
                lErro = CF("ItensOP_Baixada_Le", objOrdemProducao)
                If lErro <> SUCESSO And lErro <> 178689 Then gError 178869
            
            End If
            
            If lErro = SUCESSO Then
            
                For Each objItemOP In objOrdemProducao.colItens
                    
                    If objItemOP.lCodPedido <> 0 Then
                        GridRequisicoes.TextMatrix(iLinha, iGrid_CodPV_Col) = objItemOP.lCodPedido
                        Exit For
                    End If
                    
                    If objItemOP.lNumIntDocPai <> 0 Then
                    
                        lErro = CF("ItensOP_Le_PV", objItemOP.lNumIntDocPai, lCodigoPV, iFilialPV)
                        If lErro <> SUCESSO And lErro <> 178696 And lErro <> 178697 Then gError 178870
                
                    End If
                
                    If lCodigoPV <> 0 Then
                        GridRequisicoes.TextMatrix(iLinha, iGrid_CodPV_Col) = lCodigoPV
                        Exit For
                    End If
                
                Next
        
            End If
        
        End If
    
    Next
    
    End If
    
    Preenche_GridRequisicoes = SUCESSO
    
    Exit Function
    
Erro_Preenche_GridRequisicoes:

    Preenche_GridRequisicoes = gErr
    
    Select Case gErr
    
        Case 178867 To 178870

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178871)

    End Select
        
End Function

