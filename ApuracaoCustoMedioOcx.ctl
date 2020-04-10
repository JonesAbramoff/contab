VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ApuracaoCustoMedioOcx 
   ClientHeight    =   4605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9495
   LockControls    =   -1  'True
   ScaleHeight     =   4605
   ScaleWidth      =   9495
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   2895
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   8985
      Begin VB.Frame Frame3 
         Caption         =   "Faixas de Classificação"
         Height          =   765
         Left            =   210
         TabIndex        =   51
         Top             =   3960
         Width           =   5475
         Begin MSMask.MaskEdBox FaixaA 
            Height          =   285
            Left            =   1155
            TabIndex        =   52
            Top             =   330
            Visible         =   0   'False
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   3
            Format          =   "0\%"
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FaixaB 
            Height          =   285
            Left            =   2925
            TabIndex        =   53
            Top             =   315
            Visible         =   0   'False
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   3
            Format          =   "0\%"
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin VB.Label FaixaC 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4650
            TabIndex        =   54
            Top             =   330
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Faixa B:"
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
            Left            =   2160
            TabIndex        =   57
            Top             =   360
            Width           =   705
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Faixa A:"
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
            Left            =   375
            TabIndex        =   56
            Top             =   360
            Width           =   705
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Faixa C:"
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
            Left            =   3885
            TabIndex        =   55
            Top             =   360
            Width           =   705
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Valores"
         Height          =   1185
         Left            =   150
         TabIndex        =   46
         Top             =   750
         Width           =   8445
         Begin MSMask.MaskEdBox MatPrima 
            Height          =   315
            Index           =   1
            Left            =   1485
            TabIndex        =   7
            Top             =   285
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
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
            Format          =   "#,##0.0000"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Embalagem 
            Height          =   315
            Index           =   1
            Left            =   1470
            TabIndex        =   9
            Top             =   720
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
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
            Format          =   "#,##0.0000"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdIntermediarios 
            Height          =   315
            Index           =   1
            Left            =   6405
            TabIndex        =   8
            Top             =   285
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
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
            Format          =   "#,##0.0000"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorProduto 
            Height          =   315
            Index           =   1
            Left            =   6420
            TabIndex        =   10
            Top             =   720
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
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
            Format          =   "#,##0.0000"
            PromptChar      =   " "
         End
         Begin VB.Label LabelCustoStd 
            AutoSize        =   -1  'True
            Caption         =   "Matéria Prima:"
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
            Left            =   225
            TabIndex        =   50
            Top             =   345
            Width           =   1230
         End
         Begin VB.Label Label5 
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
            Left            =   405
            TabIndex        =   49
            Top             =   780
            Width           =   1035
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Produtos Intermediários:"
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
            Left            =   4335
            TabIndex        =   48
            Top             =   345
            Width           =   2070
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Produto:"
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
            Left            =   5670
            TabIndex        =   47
            Top             =   780
            Width           =   735
         End
      End
      Begin MSMask.MaskEdBox HorasMaquina 
         Height          =   315
         Index           =   1
         Left            =   6720
         TabIndex        =   6
         Top             =   240
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox GastosDiretos 
         Height          =   315
         Index           =   1
         Left            =   1860
         TabIndex        =   11
         Top             =   2025
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
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
         Format          =   "#,##0.0000"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Total 
         Height          =   315
         Index           =   1
         Left            =   3870
         TabIndex        =   14
         Top             =   2490
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
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
         Format          =   "#,##0.0000"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox GastosIndiretos 
         Height          =   315
         Index           =   1
         Left            =   6720
         TabIndex        =   12
         Top             =   2025
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
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
         Format          =   "#,##0.0000"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CustoMedio 
         Height          =   315
         Index           =   1
         Left            =   6735
         TabIndex        =   15
         Top             =   2490
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
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
         Format          =   "#,##0.0000"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox QuantProduto 
         Height          =   315
         Index           =   1
         Left            =   2235
         TabIndex        =   5
         Top             =   240
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
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
         Format          =   "#,##0.0000"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CustosAnteriores 
         Height          =   315
         Index           =   1
         Left            =   1875
         TabIndex        =   13
         Top             =   2505
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         _Version        =   393216
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
         Format          =   "#,##0.0000"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Custos Anteriores:"
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
         Left            =   315
         TabIndex        =   65
         Top             =   2565
         Width           =   1560
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade do Produto:"
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
         Left            =   195
         TabIndex        =   64
         Top             =   285
         Width           =   2040
      End
      Begin VB.Label LabelMinutos 
         Caption         =   "minutos"
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
         Left            =   7395
         TabIndex        =   63
         Top             =   285
         Width           =   810
      End
      Begin VB.Label LabelHorasMaq 
         AutoSize        =   -1  'True
         Caption         =   "Horas de Máquina:"
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
         Left            =   5055
         TabIndex        =   62
         Top             =   285
         Width           =   1620
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Gastos Diretos:"
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
         Left            =   525
         TabIndex        =   61
         Top             =   2085
         Width           =   1335
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
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
         Left            =   3360
         TabIndex        =   60
         Top             =   2550
         Width           =   510
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Gastos Indiretos:"
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
         Left            =   5235
         TabIndex        =   59
         Top             =   2085
         Width           =   1455
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Custo Médio:"
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
         Left            =   5580
         TabIndex        =   58
         Top             =   2520
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   2865
      Index           =   2
      Left            =   180
      TabIndex        =   16
      Top             =   1440
      Visible         =   0   'False
      Width           =   9015
      Begin VB.Frame Frame5 
         Caption         =   "Valores"
         Height          =   1185
         Left            =   150
         TabIndex        =   92
         Top             =   750
         Width           =   8445
         Begin MSMask.MaskEdBox MatPrima 
            Height          =   315
            Index           =   2
            Left            =   1485
            TabIndex        =   19
            Top             =   285
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
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
            Format          =   "#,##0.0000"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Embalagem 
            Height          =   315
            Index           =   2
            Left            =   1470
            TabIndex        =   21
            Top             =   720
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
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
            Format          =   "#,##0.0000"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdIntermediarios 
            Height          =   315
            Index           =   2
            Left            =   6405
            TabIndex        =   20
            Top             =   285
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
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
            Format          =   "#,##0.0000"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorProduto 
            Height          =   315
            Index           =   2
            Left            =   6420
            TabIndex        =   22
            Top             =   720
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
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
            Format          =   "#,##0.0000"
            PromptChar      =   " "
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Produto:"
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
            Left            =   5670
            TabIndex        =   96
            Top             =   780
            Width           =   735
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Produtos Intermediários:"
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
            Left            =   4335
            TabIndex        =   95
            Top             =   345
            Width           =   2070
         End
         Begin VB.Label Label19 
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
            Left            =   405
            TabIndex        =   94
            Top             =   780
            Width           =   1035
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Matéria Prima:"
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
            Left            =   225
            TabIndex        =   93
            Top             =   345
            Width           =   1230
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Faixas de Classificação"
         Height          =   765
         Left            =   210
         TabIndex        =   85
         Top             =   3960
         Width           =   5475
         Begin MSMask.MaskEdBox MaskEdBox10 
            Height          =   285
            Left            =   1155
            TabIndex        =   86
            Top             =   330
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   3
            Format          =   "0\%"
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaskEdBox11 
            Height          =   285
            Left            =   2925
            TabIndex        =   87
            Top             =   315
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   3
            Format          =   "0\%"
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin VB.Label Label17 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4650
            TabIndex        =   91
            Top             =   330
            Width           =   555
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Faixa C:"
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
            Left            =   3885
            TabIndex        =   90
            Top             =   360
            Width           =   705
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Faixa A:"
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
            Left            =   375
            TabIndex        =   89
            Top             =   360
            Width           =   705
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Faixa B:"
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
            Left            =   2160
            TabIndex        =   88
            Top             =   360
            Width           =   705
         End
      End
      Begin MSMask.MaskEdBox HorasMaquina 
         Height          =   315
         Index           =   2
         Left            =   6720
         TabIndex        =   18
         Top             =   240
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox GastosDiretos 
         Height          =   315
         Index           =   2
         Left            =   1860
         TabIndex        =   23
         Top             =   2025
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
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
         Format          =   "#,##0.0000"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Total 
         Height          =   315
         Index           =   2
         Left            =   1830
         TabIndex        =   25
         Top             =   2415
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
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
         Format          =   "#,##0.0000"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox GastosIndiretos 
         Height          =   315
         Index           =   2
         Left            =   6720
         TabIndex        =   24
         Top             =   2025
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
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
         Format          =   "#,##0.0000"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CustoMedio 
         Height          =   315
         Index           =   2
         Left            =   6735
         TabIndex        =   26
         Top             =   2490
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
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
         Format          =   "#,##0.0000"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox QuantProduto 
         Height          =   315
         Index           =   2
         Left            =   2235
         TabIndex        =   17
         Top             =   240
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
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
         Format          =   "#,##0.0000"
         PromptChar      =   " "
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Custo Médio:"
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
         Left            =   5535
         TabIndex        =   103
         Top             =   2520
         Width           =   1125
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Gastos Indiretos:"
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
         Left            =   5190
         TabIndex        =   102
         Top             =   2085
         Width           =   1455
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
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
         Left            =   1320
         TabIndex        =   101
         Top             =   2475
         Width           =   510
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Gastos Diretos:"
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
         Left            =   525
         TabIndex        =   100
         Top             =   2085
         Width           =   1335
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Horas de Máquina:"
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
         Left            =   5055
         TabIndex        =   99
         Top             =   285
         Width           =   1620
      End
      Begin VB.Label Label24 
         Caption         =   "minutos"
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
         Left            =   7395
         TabIndex        =   98
         Top             =   285
         Width           =   810
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade do Produto:"
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
         Left            =   195
         TabIndex        =   97
         Top             =   285
         Width           =   2040
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   2880
      Index           =   3
      Left            =   225
      TabIndex        =   27
      Top             =   1410
      Visible         =   0   'False
      Width           =   9000
      Begin VB.Frame Frame7 
         Caption         =   "Faixas de Classificação"
         Height          =   765
         Left            =   210
         TabIndex        =   71
         Top             =   3960
         Width           =   5475
         Begin MSMask.MaskEdBox MaskEdBox26 
            Height          =   285
            Left            =   1155
            TabIndex        =   72
            Top             =   330
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   3
            Format          =   "0\%"
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaskEdBox27 
            Height          =   285
            Left            =   2925
            TabIndex        =   73
            Top             =   315
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   3
            Format          =   "0\%"
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Faixa B:"
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
            Left            =   2160
            TabIndex        =   77
            Top             =   360
            Width           =   705
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Faixa A:"
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
            Left            =   375
            TabIndex        =   76
            Top             =   360
            Width           =   705
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Faixa C:"
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
            Left            =   3885
            TabIndex        =   75
            Top             =   360
            Width           =   705
         End
         Begin VB.Label Label34 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4650
            TabIndex        =   74
            Top             =   330
            Width           =   555
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Valores"
         Height          =   1185
         Left            =   150
         TabIndex        =   66
         Top             =   750
         Width           =   8445
         Begin MSMask.MaskEdBox MatPrima 
            Height          =   315
            Index           =   3
            Left            =   1485
            TabIndex        =   30
            Top             =   285
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
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
            Format          =   "#,##0.0000"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Embalagem 
            Height          =   315
            Index           =   3
            Left            =   1470
            TabIndex        =   32
            Top             =   720
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
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
            Format          =   "#,##0.0000"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdIntermediarios 
            Height          =   315
            Index           =   3
            Left            =   6405
            TabIndex        =   31
            Top             =   285
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
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
            Format          =   "#,##0.0000"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorProduto 
            Height          =   315
            Index           =   3
            Left            =   6420
            TabIndex        =   33
            Top             =   720
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
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
            Format          =   "#,##0.0000"
            PromptChar      =   " "
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Matéria Prima:"
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
            Left            =   225
            TabIndex        =   70
            Top             =   345
            Width           =   1230
         End
         Begin VB.Label Label31 
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
            Left            =   405
            TabIndex        =   69
            Top             =   780
            Width           =   1035
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Produtos Intermediários:"
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
            Left            =   4335
            TabIndex        =   68
            Top             =   345
            Width           =   2070
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Produto:"
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
            Left            =   5670
            TabIndex        =   67
            Top             =   780
            Width           =   735
         End
      End
      Begin MSMask.MaskEdBox HorasMaquina 
         Height          =   315
         Index           =   3
         Left            =   6720
         TabIndex        =   29
         Top             =   240
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox GastosDiretos 
         Height          =   315
         Index           =   3
         Left            =   1860
         TabIndex        =   34
         Top             =   2025
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
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
         Format          =   "#,##0.0000"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Total 
         Height          =   315
         Index           =   3
         Left            =   1830
         TabIndex        =   36
         Top             =   2415
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
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
         Format          =   "#,##0.0000"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox GastosIndiretos 
         Height          =   315
         Index           =   3
         Left            =   6720
         TabIndex        =   35
         Top             =   2025
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
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
         Format          =   "#,##0.0000"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CustoMedio 
         Height          =   315
         Index           =   3
         Left            =   6735
         TabIndex        =   37
         Top             =   2490
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
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
         Format          =   "#,##0.0000"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox QuantProduto 
         Height          =   315
         Index           =   3
         Left            =   2235
         TabIndex        =   28
         Top             =   240
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
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
         Format          =   "#,##0.0000"
         PromptChar      =   " "
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade do Produto:"
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
         Left            =   195
         TabIndex        =   84
         Top             =   285
         Width           =   2040
      End
      Begin VB.Label Label43 
         Caption         =   "minutos"
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
         Left            =   7395
         TabIndex        =   83
         Top             =   285
         Width           =   810
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Horas de Máquina:"
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
         Left            =   5055
         TabIndex        =   82
         Top             =   285
         Width           =   1620
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "Gastos Diretos:"
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
         Left            =   525
         TabIndex        =   81
         Top             =   2085
         Width           =   1335
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
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
         Left            =   1320
         TabIndex        =   80
         Top             =   2475
         Width           =   510
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "Gastos Indiretos:"
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
         Left            =   5190
         TabIndex        =   79
         Top             =   2085
         Width           =   1455
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "Custo Médio:"
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
         Left            =   5535
         TabIndex        =   78
         Top             =   2520
         Width           =   1125
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7230
      ScaleHeight     =   495
      ScaleWidth      =   1665
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   165
      Width           =   1725
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ApuracaoCustoMedioOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   615
         Picture         =   "ApuracaoCustoMedioOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1125
         Picture         =   "ApuracaoCustoMedioOcx.ctx":068C
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3435
      Left            =   75
      TabIndex        =   0
      Top             =   1050
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   6059
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Concluídos no mês/Iniciados antes"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Iniciados/Concluídos no mês"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Iniciados/Não concluídos no mês"
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
   Begin MSMask.MaskEdBox Ano 
      Height          =   315
      Left            =   810
      TabIndex        =   1
      Top             =   165
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Produto 
      Height          =   315
      Left            =   1125
      TabIndex        =   3
      Top             =   630
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      AllowPrompt     =   -1  'True
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Mes 
      Height          =   315
      Left            =   2115
      TabIndex        =   2
      Top             =   165
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin VB.Label DescProduto 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2790
      TabIndex        =   45
      Top             =   630
      Width           =   3015
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      Caption         =   "Mês:"
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
      Left            =   1650
      TabIndex        =   44
      Top             =   225
      Width           =   420
   End
   Begin VB.Label LabelProduto 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   375
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   43
      Top             =   675
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Ano:"
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
      Left            =   345
      TabIndex        =   42
      Top             =   225
      Width           =   405
   End
End
Attribute VB_Name = "ApuracaoCustoMedioOcx"
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

Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Inicializa o Frame atual com 1
    iFrameAtual = 1

    Set objEventoProduto = New AdmEvento
    
    'Inicializa as máscaras de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd",Produto)
    If lErro <> SUCESSO Then gError 76477

    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case Err

        Case 76477
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143000)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Private Sub Mes_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Mes_Validate

    'Verifica se o mês foi preenchido
    If Len(Trim(Mes.Text)) > 0 Then
    
        'Verifica se é um mês válido
        If StrParaInt(Mes.Text) > 12 Then gError 76486
        
    End If
    
    Exit Sub
    
Erro_Mes_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 76486
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INTEIRO_NAO_MES", gErr, StrParaInt(Mes.Text))
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143001)
            
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

Function MovEstoque_Le_ApropInsumosProduto(iMes As Integer, iAno As Integer, iFilialEmpresa As Integer) As Long
'Lê os Movimentos de Estoque(de requisicao de producao)para o mês informado e que estejam relacionados à tabela ApropriacaoInsumosProd
'Guarda os valores de Quantidade e HorasMaquina lidos em MovimentoEstoque

Dim lErro As Long
Dim lComando As Long
Dim tMovEstoque As typeMovEst
Dim objProduto As New ClassProduto
Dim dtDataInicial As Date
Dim dtDataFinal As Date
Dim lNumIntDoc As Long
Dim sProdAprop As String
Dim sProdMovEst As String
Dim dCusto As Double
Dim dQuantidadeProd As Double
Dim dQuantProdAprop As Double
Dim dQuantMateriaPrima As Double
Dim dQuantEmbalagem As Double
Dim dQuantOutros As Double
Dim dQuantidadeProprio As Double
Dim lHorasMaquina As Long
Dim lTotalHorasMaq As Long
Dim iDias As Integer
Dim iNatureza As Integer
Dim lNumIntDocLido As Long
Dim bMovEstoqueAlterado As Boolean

On Error GoTo Erro_MovEstoque_Le_ApropInsumosProduto

    bMovEstoqueAlterado = False
    
    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 76472
    
    'Calcula o numero de dias do mês atual
    iDias = Dias_Mes(iMes, iAno)
    
    dtDataInicial = CDate("01/" & iMes & "/" & iAno)
    dtDataFinal = CDate(iDias & "/" & iMes & "/" & iAno)
    
    sProdMovEst = String(STRING_PRODUTO, 0)
    sProdAprop = String(STRING_PRODUTO, 0)

    'Busca os Insumos utilizados em Producao de Entrada para o periodo informado
    '??substituir pela view MovEstoque_ApropriacaoInsumosProd (SGEDados26)
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Codigo, NumIntDoc, Custo, Quantidade, QuantAprop, ProdAprop, Natureza, Produto," _
    & "HorasMaquina, DataInicioProducao FROM MovEstoque_ApropriacaoInsumosProd WHERE DataInicioProducao>=? AND DataInicioProducao<=?" _
    & "AND FilialEmpresa=? ORDER BY NumIntDoc", tMovEstoque.iFilialEmpresa, tMovEstoque.lCodigo, lNumIntDoc, dCusto, dQuantidadeProd, _
    dQuantProdAprop, sProdAprop, iNatureza, sProdMovEst, dtDataInicial, dtDataFinal, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 76473
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76474
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        If lNumIntDoc <> lNumIntDocLido Then bMovEstoqueAlterado = True
        
        'Verifica se o produto utilizado para a producao não é ele mesmo
        If sProdAprop <> sProdMovEst Then
            
            'Verifica se a Natureza do produto é matéria prima
            If iNatureza = NATUREZA_PROD_MATERIA_PRIMA Then
            
                'Guarda a quantidade de materia prima utilizada
                dQuantMateriaPrima = dQuantMateriaPrima + dQuantProdAprop
            
            'Se o produto é embalagem
            ElseIf iNatureza = NATUREZA_PROD_EMBALAGENS Then
                
                'Guarda a quantidade de embalagem utilizada
                dQuantEmbalagem = dQuantEmbalagem + dQuantProdAprop
            
            'Se não é embalagem e nem matéria prima
            ElseIf iNatureza <> NATUREZA_PROD_MATERIA_PRIMA And iNatureza <> NATUREZA_PROD_EMBALAGENS Then
                
                'Guarda a quantidade dos outros produtos utilizados
                dQuantOutros = dQuantOutros + dQuantProdAprop
                
            End If
            
        'Se o próprio produto está sendo utilizado na sua producao
        Else
        
            'Guarda a quantidade do proprio produto e as horas de maquina utilizadas
            dQuantidadeProprio = dQuantidadeProprio + dQuantProdAprop
            
        End If
        
        lNumIntDocLido = lNumIntDoc
        
        If bMovEstoqueAlterado = True Then
        
            lTotalHorasMaq = lTotalHorasMaq + lHorasMaquina
            bMovEstoqueAlterado = False
            
        End If
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76476
        
    Loop
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    MovEstoque_Le_ApropInsumosProduto = SUCESSO
    
    Exit Function
    
Erro_MovEstoque_Le_ApropInsumosProduto:

    MovEstoque_Le_ApropInsumosProduto = gErr
    
    Select Case gErr
    
        Case 76472
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 76473, 76474, 76476
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)
            
        Case 76475
            'Erro tratado na rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143002)
            
    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Public Function Trata_Parametros(Optional objProduto As ClassProduto) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143003)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""
''''''????Esperando
''''''Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'''''''Extrai os campos da tela que correspondem aos campos no BD
''''''
''''''Dim lErro As Long
''''''Dim objProduto As New ClassProduto
''''''Dim colTabelaPrecoItem As New Collection
''''''
''''''On Error GoTo Erro_Tela_Extrai
''''''
''''''    'Informa tabela associada à Tela
''''''    sTabela = "Produtos"
''''''
''''''    'Lê os dados da Tela Notas Fiscais a Pagar
''''''    lErro = Move_Tela_Memoria(objProduto, colTabelaPrecoItem)
''''''    If lErro <> SUCESSO Then Error 64482
''''''
''''''    'Preenche a coleção colCampoValor, com nome do campo,
''''''    'valor atual (com a tipagem do BD), tamanho do campo
''''''    'no BD no caso de STRING e Key igual ao nome do campo
''''''    colCampoValor.Add "Codigo", objProduto.sCodigo, STRING_PRODUTO, "Codigo"
''''''    colCampoValor.Add "Tipo", objProduto.iTipo, 0, "Tipo"
''''''    colCampoValor.Add "Descricao", objProduto.sDescricao, STRING_PRODUTO_DESCRICAO, "Descricao"
''''''    colCampoValor.Add "NomeReduzido", objProduto.sNomeReduzido, STRING_PRODUTO_NOME_REDUZIDO, "NomeReduzido"
''''''    colCampoValor.Add "Modelo", objProduto.sModelo, STRING_PRODUTO_MODELO, "Modelo"
''''''    colCampoValor.Add "Gerencial", objProduto.iGerencial, 0, "Gerencial"
''''''    colCampoValor.Add "Nivel", objProduto.iNivel, 0, "Nivel"
''''''    colCampoValor.Add "Substituto1", objProduto.sSubstituto1, STRING_PRODUTO_SUBSTITUTO1, "Substituto1"
''''''    colCampoValor.Add "Substituto2", objProduto.sSubstituto2, STRING_PRODUTO_SUBSTITUTO2, "Substituto2"
''''''    colCampoValor.Add "PrazoValidade", objProduto.iPrazoValidade, 0, "PrazoValidade"
''''''    colCampoValor.Add "CodigoBarras", objProduto.sCodigoBarras, STRING_PRODUTO_CODIGO_BARRAS, "CodigoBarras"
''''''    colCampoValor.Add "EtiquetasCodBarras", objProduto.iEtiquetasCodBarras, 0, "EtiquetasCodBarras"
''''''    colCampoValor.Add "PesoLiq", objProduto.dPesoLiq, 0, "PesoLiq"
''''''    colCampoValor.Add "PesoBruto", objProduto.dPesoBruto, 0, "PesoBruto"
''''''    colCampoValor.Add "Comprimento", objProduto.dComprimento, 0, "Comprimento"
''''''    colCampoValor.Add "Espessura", objProduto.dEspessura, 0, "Espessura"
''''''    colCampoValor.Add "Largura", objProduto.dLargura, 0, "Largura"
''''''    colCampoValor.Add "Cor", objProduto.sCor, STRING_PRODUTO_COR, "Cor"
''''''    colCampoValor.Add "ObsFisica", objProduto.sObsFisica, STRING_PRODUTO_OBS_FISICA, "ObsFisica"
''''''    colCampoValor.Add "ClasseUM", objProduto.iClasseUM, 0, "ClasseUM"
''''''    colCampoValor.Add "SiglaUMCompra", objProduto.sSiglaUMCompra, STRING_PRODUTO_SIGLAUMCOMPRA, "SiglaUMCompra"
''''''    colCampoValor.Add "SiglaUMEstoque", objProduto.sSiglaUMEstoque, STRING_PRODUTO_SIGLAUMESTOQUE, "SiglaUMEstoque"
''''''    colCampoValor.Add "SiglaUMVenda", objProduto.sSiglaUMVenda, STRING_PRODUTO_SIGLAUMVENDA, "SiglaUMVenda"
''''''    colCampoValor.Add "Ativo", objProduto.iAtivo, 0, "Ativo"
''''''    colCampoValor.Add "Faturamento", objProduto.iFaturamento, 0, "Faturamento"
''''''    colCampoValor.Add "Compras", objProduto.iCompras, 0, "Compras"
''''''    colCampoValor.Add "PCP", objProduto.iPCP, 0, "PCP"
''''''    colCampoValor.Add "KitBasico", objProduto.iKitBasico, 0, "KitBasico"
''''''    colCampoValor.Add "KitInt", objProduto.iKitInt, 0, "KitInt"
''''''    colCampoValor.Add "IPIAliquota", objProduto.dIPIAliquota, 0, "IPIAliquota"
''''''    colCampoValor.Add "IPICodigo", objProduto.sIPICodigo, STRING_PRODUTO_IPI_CODIGO, "IPICodigo"
''''''    colCampoValor.Add "IPICodDIPI", objProduto.sIPICodDIPI, STRING_PRODUTO_IPI_COD_DIPI, "IPICodDIPI"
''''''    colCampoValor.Add "Apropriacao", objProduto.iApropriacaoCusto, 0, "Apropriacao"
''''''    colCampoValor.Add "ContaContabil", objProduto.sContaContabil, STRING_CONTA, "ContaContabil"
''''''    colCampoValor.Add "Natureza", objProduto.iNatureza, 0, "Natureza"
''''''    colCampoValor.Add "ContaContabilProducao", objProduto.sContaContabilProducao, STRING_CONTA, "ContaContabilProducao"
''''''    colCampoValor.Add "PercentMaisQuantCotacaoAnterior", objProduto.dPercentMaisQuantCotacaoAnterior, 0, "PercentMaisQuantCotacaoAnterior"
''''''    colCampoValor.Add "PercentMenosQuantCotacaoAnterior", objProduto.dPercentMenosQuantCotacaoAnterior, 0, "PercentMenosQuantCotacaoAnterior"
''''''    colCampoValor.Add "ConsideraQuantCotacaoAnterior", objProduto.iConsideraQuantCotacaoAnterior, 0, "ConsideraQuantCotacaoAnterior"
''''''    colCampoValor.Add "TemFaixaReceb", objProduto.iTemFaixaReceb, 0, "TemFaixaReceb"
''''''    colCampoValor.Add "PercentMaisReceb", objProduto.dPercentMaisReceb, 0, "PercentMaisReceb"
''''''    colCampoValor.Add "PercentMenosReceb", objProduto.dPercentMenosReceb, 0, "PercentMenosReceb"
''''''    colCampoValor.Add "RecebForaFaixa", objProduto.iRecebForaFaixa, 0, "RecebForaFaixa"
''''''    colCampoValor.Add "CreditoICMS", objProduto.iCreditoICMS, 0, "CreditoICMS"
''''''    colCampoValor.Add "CreditoIPI", objProduto.iCreditoIPI, 0, "CreditoIPI"
''''''    colCampoValor.Add "Residuo", objProduto.dResiduo, 0, "Residuo"
''''''    colCampoValor.Add "CustoReposicao", objProduto.dCustoReposicao, 0, "CustoReposicao"
''''''    colCampoValor.Add "OrigemMercadoria", objProduto.iOrigemMercadoria, 0, "OrigemMercadoria"
''''''    colCampoValor.Add "TempoProducao", objProduto.iTempoProducao, 0, "TempoProducao"
''''''    colCampoValor.Add "Rastro", objProduto.iRastro, 0, "Rastro"
''''''    colCampoValor.Add "HorasMaquina", objProduto.lHorasMaquina, 0, "HorasMaquina"
''''''    colCampoValor.Add "PesoEspecifico", objProduto.dPesoEspecifico, 0, "PesoEspecifico"
''''''
''''''    Exit Sub
''''''
''''''Erro_Tela_Extrai:
''''''
''''''    Select Case Err
''''''
''''''        Case 64482
''''''
''''''        Case Else
''''''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143004)
''''''
''''''    End Select
''''''
''''''    Exit Sub
''''''
''''''End Sub
''''''
'''''''Preenche os campos da tela com os correspondentes do BD
''''''Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
''''''
''''''Dim lErro As Long
''''''Dim objProduto As New ClassProduto
''''''Dim iControleEstoque As Integer
''''''
''''''On Error GoTo Erro_Tela_Preenche
''''''
''''''    objProduto.sCodigo = colCampoValor.Item("Codigo").vValor
''''''
''''''    If objProduto.sCodigo <> 0 Then
''''''
''''''        'Carrega objProduto com os dados passados em colCampoValor
''''''        objProduto.iTipo = colCampoValor.Item("Tipo").vValor
''''''        objProduto.sDescricao = colCampoValor.Item("Descricao").vValor
''''''        objProduto.sNomeReduzido = colCampoValor.Item("NomeReduzido").vValor
''''''        objProduto.sModelo = colCampoValor.Item("Modelo").vValor
''''''        objProduto.iGerencial = colCampoValor.Item("Gerencial").vValor
''''''        objProduto.iNivel = colCampoValor.Item("Nivel").vValor
''''''        objProduto.sSubstituto1 = colCampoValor.Item("Substituto1").vValor
''''''        objProduto.sSubstituto2 = colCampoValor.Item("Substituto2").vValor
''''''        objProduto.iPrazoValidade = colCampoValor.Item("PrazoValidade").vValor
''''''        objProduto.sCodigoBarras = colCampoValor.Item("CodigoBarras").vValor
''''''        objProduto.iEtiquetasCodBarras = colCampoValor.Item("EtiquetasCodBarras").vValor
''''''        objProduto.dPesoLiq = colCampoValor.Item("PesoLiq").vValor
''''''        objProduto.dPesoBruto = colCampoValor.Item("PesoBruto").vValor
''''''        objProduto.dComprimento = colCampoValor.Item("Comprimento").vValor
''''''        objProduto.dEspessura = colCampoValor.Item("Espessura").vValor
''''''        objProduto.dLargura = colCampoValor.Item("Largura").vValor
''''''        objProduto.sCor = colCampoValor.Item("Cor").vValor
''''''        objProduto.sObsFisica = colCampoValor.Item("ObsFisica").vValor
''''''        objProduto.iClasseUM = colCampoValor.Item("ClasseUM").vValor
''''''        objProduto.sSiglaUMCompra = colCampoValor.Item("SiglaUMCompra").vValor
''''''        objProduto.sSiglaUMEstoque = colCampoValor.Item("SiglaUMEstoque").vValor
''''''        objProduto.sSiglaUMVenda = colCampoValor.Item("SiglaUMVenda").vValor
''''''        objProduto.iAtivo = colCampoValor.Item("Ativo").vValor
''''''        objProduto.iFaturamento = colCampoValor.Item("Faturamento").vValor
''''''        objProduto.iCompras = colCampoValor.Item("Compras").vValor
''''''        objProduto.iPCP = colCampoValor.Item("PCP").vValor
''''''        objProduto.iKitBasico = colCampoValor.Item("KitBasico").vValor
''''''        objProduto.iKitInt = colCampoValor.Item("KitInt").vValor
''''''        objProduto.dIPIAliquota = colCampoValor.Item("IPIAliquota").vValor
''''''        objProduto.sIPICodigo = colCampoValor.Item("IPICodigo").vValor
''''''        objProduto.sIPICodDIPI = colCampoValor.Item("IPICodDIPI").vValor
''''''        objProduto.iApropriacaoCusto = colCampoValor.Item("Apropriacao").vValor
''''''        objProduto.sContaContabil = colCampoValor.Item("ContaContabil").vValor
''''''        objProduto.iNatureza = colCampoValor.Item("Natureza").vValor
''''''
''''''        objProduto.sContaContabilProducao = colCampoValor.Item("ContaContabilProducao").vValor
''''''        objProduto.dPercentMaisQuantCotacaoAnterior = colCampoValor.Item("PercentMaisQuantCotacaoAnterior").vValor
''''''        objProduto.dPercentMenosQuantCotacaoAnterior = colCampoValor.Item("PercentMenosQuantCotacaoAnterior").vValor
''''''        objProduto.iConsideraQuantCotacaoAnterior = colCampoValor.Item("ConsideraQuantCotacaoAnterior").vValor
''''''        objProduto.iTemFaixaReceb = colCampoValor.Item("TemFaixaReceb").vValor
''''''        objProduto.dPercentMaisReceb = colCampoValor.Item("PercentMaisReceb").vValor
''''''        objProduto.dPercentMenosReceb = colCampoValor.Item("PercentMenosReceb").vValor
''''''        objProduto.iRecebForaFaixa = colCampoValor.Item("RecebForaFaixa").vValor
''''''        objProduto.iCreditoICMS = colCampoValor.Item("CreditoICMS").vValor
''''''        objProduto.iCreditoIPI = colCampoValor.Item("CreditoIPI").vValor
''''''        objProduto.dResiduo = colCampoValor.Item("Residuo").vValor
''''''        objProduto.iNatureza = colCampoValor.Item("Natureza").vValor
''''''        objProduto.dCustoReposicao = colCampoValor.Item("CustoReposicao").vValor
''''''        objProduto.iOrigemMercadoria = colCampoValor.Item("OrigemMercadoria").vValor
''''''        objProduto.iTempoProducao = colCampoValor.Item("TempoProducao").vValor
''''''        objProduto.iRastro = colCampoValor.Item("Rastro").vValor
''''''        objProduto.lHorasMaquina = colCampoValor.Item("HorasMaquina").vValor
''''''        objProduto.dPesoEspecifico = colCampoValor.Item("PesoEspecifico").vValor
''''''
''''''        lErro = Traz_Produto_Tela(objProduto)
''''''        If lErro <> SUCESSO Then Error 64483
''''''
''''''    End If
''''''
''''''    Exit Sub
''''''
''''''Erro_Tela_Preenche:
''''''
''''''    Select Case Err
''''''
''''''        Case 64483
''''''
''''''        Case Else
''''''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143005)
''''''
''''''    End Select
''''''
''''''    Exit Sub
''''''
''''''End Sub
''''???Esperando definir sistema de setas
'  Public Sub Form_Activate()

   ' Call TelaIndice_Preenche(Me)

'End Sub
''''???Esperando definir sistema de setas
'Public Sub Form_Deactivate()

 '   gi_ST_SetaIgnoraClick = 1

'End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoProduto = Nothing
    
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    '???Parent.HelpContextID = IDH_PRODUTO_DADOS_PRINCIPAIS
    Set Form_Load_Ocx = Me
    Caption = "Apuração De Custo Médio"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ApuracaoCustoMedio"
    
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

Private Sub Ano_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Ano_GotFocus()

    Call MaskEdBox_TrataGotFocus(Ano, iAlterado)
    
End Sub

Private Sub BotaoFechar_Click()
    
    Unload Me
    
End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 76481

    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    'Limpa a Tela
    Call Limpa_Tela_ApuracaoCustoMedio

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 76481

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143006)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_ApuracaoCustoMedio()

Dim iIndice As Integer

    'Chama Limpa_Tela
    Call Limpa_Tela(Me)

    'Limpa os Campos não limpos no limpa tela
    DescProduto.Caption = ""
    
    iAlterado = 0

End Sub

Private Sub LabelProduto_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto

    '??Confimar tela de browse
    Call Chama_Tela("ProdutoProduz_EstoqLista", colSelecao, objProduto, objEventoProduto)
    
    Exit Sub
    
End Sub

Private Sub Mes_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Mes_GotFocus()

    Call MaskEdBox_TrataGotFocus(Mes, iAlterado)
    
End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String
    
On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1
    
    If objProduto.iCompras <> PRODUTO_PRODUZIVEL Then gError 76482
    
    lErro = CF("Produto_Formata",Produto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 76483

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then gError 76484

    sProdutoMascarado = String(STRING_PRODUTO, 0)

    lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 76485

    Produto.PromptInclude = False
    Produto.Text = sProdutoMascarado
    Produto.PromptInclude = True
    
    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr
        
        Case 76482
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL", gErr, Produto.Text)
        
        Case 76483, 76484
            'Erros tratados nas rotinas chamadas
            
        Case 76485
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objProduto.sCodigo)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143007)

    End Select

    Exit Sub

End Sub

Private Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sProduto As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_Produto_Validate

    If Len(Trim(Produto.ClipText)) = 0 Then
        DescProduto.Caption = ""
        Exit Sub
    End If

    sProduto = Produto.Text

    'Critica o formato do Produto e se existe no BD
    lErro = CF("Produto_Critica",sProduto, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 25041 Then gError 76478

    'se o produto não estiver cadastrado ==> erro
    If lErro = 25041 Then gError 76479

    'se o produto não é produzido
    If objProduto.iCompras <> PRODUTO_PRODUZIVEL Then gError 76480
    
    DescProduto.Caption = objProduto.sDescricao

    Exit Sub

Erro_Produto_Validate:

    Cancel = True

    Select Case gErr

        Case 76478

        Case 76479
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 76480
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143008)

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

Public Sub Unload(objme As Object)
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

'**** fim do trecho a ser copiado *****

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
        
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Produto Then
            Call LabelProduto_Click
        End If
        
    End If

End Sub
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Height() As Long
    Height = UserControl.Height
End Property

Public Property Get Width() As Long
    Width = UserControl.Width
End Property

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property

