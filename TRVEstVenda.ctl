VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl TRVEstVenda 
   ClientHeight    =   7395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   7395
   ScaleMode       =   0  'User
   ScaleWidth      =   9510
   Begin VB.Frame Frame10 
      Caption         =   "Filtros"
      Height          =   1815
      Left            =   120
      TabIndex        =   76
      Top             =   60
      Width           =   7875
      Begin VB.ComboBox Marca 
         Height          =   315
         ItemData        =   "TRVEstVenda.ctx":0000
         Left            =   5340
         List            =   "TRVEstVenda.ctx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1080
         Width           =   1560
      End
      Begin VB.CommandButton BotaoConsultar 
         Caption         =   "Consultar"
         Height          =   675
         Left            =   6930
         Picture         =   "TRVEstVenda.ctx":0033
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Consultar"
         Top             =   1080
         Width           =   825
      End
      Begin VB.Frame Frame12 
         Caption         =   "Clientes"
         Height          =   855
         Left            =   4710
         TabIndex        =   83
         Top             =   180
         Width           =   3105
         Begin MSMask.MaskEdBox ClienteDe 
            Height          =   300
            Left            =   630
            TabIndex        =   8
            Top             =   165
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ClienteAte 
            Height          =   300
            Left            =   630
            TabIndex        =   9
            Top             =   480
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin VB.Label LabelClienteDe 
            AutoSize        =   -1  'True
            Caption         =   "De:"
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
            Left            =   195
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   85
            Top             =   210
            Width           =   315
         End
         Begin VB.Label LabelClienteAte 
            AutoSize        =   -1  'True
            Caption         =   "Até:"
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
            Left            =   210
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   84
            Top             =   540
            Width           =   360
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Período de emissão das Faturas"
         Height          =   705
         Left            =   150
         TabIndex        =   80
         Top             =   1050
         Width           =   4485
         Begin MSComCtl2.UpDown UpDownFatDe 
            Height          =   300
            Left            =   1725
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   285
            Width           =   225
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataFatDe 
            Height          =   300
            Left            =   600
            TabIndex        =   4
            Top             =   285
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataFatAte 
            Height          =   300
            Left            =   2475
            TabIndex        =   6
            Top             =   270
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownFatAte 
            Height          =   300
            Left            =   3645
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   270
            Width           =   225
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "De:"
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
            Index           =   3
            Left            =   210
            TabIndex        =   82
            Top             =   315
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Até:"
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
            Index           =   4
            Left            =   2070
            TabIndex        =   81
            Top             =   315
            Width           =   360
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Período de emissão dos vouchers"
         Height          =   855
         Left            =   150
         TabIndex        =   77
         Top             =   180
         Width           =   4485
         Begin MSComCtl2.UpDown UpDownEmiDe 
            Height          =   300
            Left            =   1710
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   300
            Width           =   225
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEmiDe 
            Height          =   300
            Left            =   585
            TabIndex        =   0
            Top             =   315
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataEmiAte 
            Height          =   300
            Left            =   2460
            TabIndex        =   2
            Top             =   300
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownEmiAte 
            Height          =   300
            Left            =   3630
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   300
            Width           =   225
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Até:"
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
            Index           =   0
            Left            =   2055
            TabIndex        =   79
            Top             =   345
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "De:"
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
            Index           =   1
            Left            =   195
            TabIndex        =   78
            Top             =   345
            Width           =   315
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Marca:"
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
         Left            =   4695
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   87
         Top             =   1125
         Width           =   600
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Vouchers emitidos e faturados no período"
      Height          =   1110
      Left            =   105
      TabIndex        =   69
      Top             =   6195
      Width           =   9165
      Begin VB.Frame Frame11 
         BorderStyle     =   0  'None
         Caption         =   "Todos"
         Height          =   870
         Left            =   135
         TabIndex        =   70
         Top             =   180
         Width           =   8880
         Begin VB.TextBox QtdeEmiPaxFat 
            Height          =   315
            Left            =   4470
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   510
            Width           =   1140
         End
         Begin VB.TextBox QtdeEmiVouFat 
            Height          =   315
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   510
            Width           =   1140
         End
         Begin VB.TextBox ValorFatVouFat 
            Height          =   315
            Left            =   7545
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   150
            Width           =   1140
         End
         Begin VB.TextBox ValorBrutoVouUSSFat 
            Height          =   315
            Left            =   4470
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox ValorBrutoVouRSFat 
            Height          =   315
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   180
            Width           =   1140
         End
         Begin VB.CommandButton BotaoExibir 
            Caption         =   "Exibir"
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
            Index           =   6
            Left            =   7545
            TabIndex        =   41
            Top             =   510
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor Bruto R$:"
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
            Index           =   31
            Left            =   315
            TabIndex        =   75
            Top             =   210
            Width           =   1320
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor Bruto US$:"
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
            Index           =   30
            Left            =   3015
            TabIndex        =   74
            Top             =   195
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor Faturável R$:"
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
            Index           =   29
            Left            =   5835
            TabIndex        =   73
            Top             =   195
            Width           =   1650
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Qtde Vouchers:"
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
            Index           =   28
            Left            =   315
            TabIndex        =   72
            Top             =   540
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   27
            Left            =   3585
            TabIndex        =   71
            Top             =   540
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Faturas emitidas no período"
      Height          =   2235
      Left            =   105
      TabIndex        =   58
      Top             =   3930
      Width           =   9165
      Begin VB.Frame Frame6 
         Caption         =   "Canceladas"
         Height          =   945
         Left            =   150
         TabIndex        =   64
         Top             =   1155
         Width           =   8880
         Begin VB.TextBox QtdeFatCPCanc 
            Height          =   315
            Left            =   4395
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   525
            Width           =   1140
         End
         Begin VB.TextBox ValorFatCPCanc 
            Height          =   315
            Left            =   1665
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   540
            Width           =   1140
         End
         Begin VB.TextBox QtdeFatCRCanc 
            Height          =   315
            Left            =   4395
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox ValorFatCRCanc 
            Height          =   315
            Left            =   1665
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   195
            Width           =   1140
         End
         Begin VB.CommandButton BotaoExibir 
            Caption         =   "Exibir"
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
            Index           =   5
            Left            =   5595
            TabIndex        =   35
            Top             =   540
            Width           =   1140
         End
         Begin VB.CommandButton BotaoExibir 
            Caption         =   "Exibir"
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
            Index           =   4
            Left            =   5595
            TabIndex        =   32
            Top             =   180
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Qtde Faturas CR:"
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
            Index           =   17
            Left            =   2910
            TabIndex        =   68
            Top             =   225
            Width           =   1485
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor a Receber:"
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
            Index           =   16
            Left            =   150
            TabIndex        =   67
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Qtde Faturas CP:"
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
            Index           =   15
            Left            =   2910
            TabIndex        =   66
            Top             =   585
            Width           =   1470
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor a Pagar:"
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
            Index           =   14
            Left            =   375
            TabIndex        =   65
            Top             =   600
            Width           =   1230
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Todas"
         Height          =   945
         Left            =   135
         TabIndex        =   59
         Top             =   195
         Width           =   8880
         Begin VB.TextBox QtdeFatCP 
            Height          =   315
            Left            =   4410
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   540
            Width           =   1140
         End
         Begin VB.TextBox ValorFatCP 
            Height          =   315
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   555
            Width           =   1140
         End
         Begin VB.TextBox QtdeFatCR 
            Height          =   315
            Left            =   4410
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   195
            Width           =   1140
         End
         Begin VB.TextBox ValorFatCR 
            Height          =   315
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   210
            Width           =   1140
         End
         Begin VB.CommandButton BotaoExibir 
            Caption         =   "Exibir"
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
            Index           =   3
            Left            =   5595
            TabIndex        =   29
            Top             =   540
            Width           =   1140
         End
         Begin VB.CommandButton BotaoExibir 
            Caption         =   "Exibir"
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
            Index           =   2
            Left            =   5610
            TabIndex        =   26
            Top             =   180
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor a Pagar:"
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
            Index           =   20
            Left            =   375
            TabIndex        =   63
            Top             =   600
            Width           =   1230
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Qtde Faturas CP:"
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
            Index           =   19
            Left            =   2910
            TabIndex        =   62
            Top             =   585
            Width           =   1470
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor a Receber:"
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
            Index           =   23
            Left            =   150
            TabIndex        =   61
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Qtde Faturas CR:"
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
            Index           =   22
            Left            =   2910
            TabIndex        =   60
            Top             =   225
            Width           =   1485
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Vouchers emitidos no período"
      Height          =   1995
      Left            =   105
      TabIndex        =   45
      Top             =   1920
      Width           =   9165
      Begin VB.Frame Frame4 
         Caption         =   "Cancelados"
         Height          =   870
         Left            =   135
         TabIndex        =   52
         Top             =   1050
         Width           =   8880
         Begin VB.TextBox QtdeEmiPaxCanc 
            Height          =   315
            Left            =   4500
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   495
            Width           =   1140
         End
         Begin VB.TextBox QtdeEmiVouCanc 
            Height          =   315
            Left            =   1695
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   495
            Width           =   1140
         End
         Begin VB.TextBox ValorFatVouCanc 
            Height          =   315
            Left            =   7575
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   150
            Width           =   1140
         End
         Begin VB.TextBox ValorBrutoVouUSSCanc 
            Height          =   315
            Left            =   4500
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   165
            Width           =   1140
         End
         Begin VB.TextBox ValorBrutoVouRSCanc 
            Height          =   315
            Left            =   1695
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   165
            Width           =   1140
         End
         Begin VB.CommandButton BotaoExibir 
            Caption         =   "Exibir"
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
            Index           =   1
            Left            =   7560
            TabIndex        =   23
            Top             =   495
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor Bruto R$:"
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
            Index           =   13
            Left            =   315
            TabIndex        =   57
            Top             =   210
            Width           =   1320
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor Bruto US$:"
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
            Index           =   12
            Left            =   3015
            TabIndex        =   56
            Top             =   195
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor Faturável R$:"
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
            Index           =   11
            Left            =   5835
            TabIndex        =   55
            Top             =   195
            Width           =   1650
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Qtde Vouchers:"
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
            Index           =   10
            Left            =   315
            TabIndex        =   54
            Top             =   540
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   9
            Left            =   3585
            TabIndex        =   53
            Top             =   540
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Todos"
         Height          =   870
         Left            =   135
         TabIndex        =   46
         Top             =   195
         Width           =   8880
         Begin VB.TextBox QtdeEmiPax 
            Height          =   315
            Left            =   4485
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   510
            Width           =   1140
         End
         Begin VB.TextBox QtdeEmiVou 
            Height          =   315
            Left            =   1695
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   510
            Width           =   1140
         End
         Begin VB.TextBox ValorFatVou 
            Height          =   315
            Left            =   7560
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   150
            Width           =   1140
         End
         Begin VB.TextBox ValorBrutoVouUSS 
            Height          =   315
            Left            =   4485
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox ValorBrutoVouRS 
            Height          =   315
            Left            =   1695
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   180
            Width           =   1140
         End
         Begin VB.CommandButton BotaoExibir 
            Caption         =   "Exibir"
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
            Index           =   0
            Left            =   7560
            TabIndex        =   17
            Top             =   495
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   8
            Left            =   3585
            TabIndex        =   51
            Top             =   540
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Qtde Vouchers:"
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
            Index           =   7
            Left            =   315
            TabIndex        =   50
            Top             =   540
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor Faturável R$:"
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
            Index           =   6
            Left            =   5835
            TabIndex        =   49
            Top             =   195
            Width           =   1650
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor Bruto US$:"
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
            Index           =   5
            Left            =   3015
            TabIndex        =   48
            Top             =   195
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor Bruto R$:"
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
            Left            =   315
            TabIndex        =   47
            Top             =   210
            Width           =   1320
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   8265
      ScaleHeight     =   495
      ScaleWidth      =   975
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   150
      Width           =   1035
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   30
         Picture         =   "TRVEstVenda.ctx":0B01
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   495
         Picture         =   "TRVEstVenda.ctx":1033
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Label Empresa 
      Caption         =   "TVA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   8265
      TabIndex        =   86
      Top             =   795
      Width           =   1065
   End
End
Attribute VB_Name = "TRVEstVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Dim iAlterado As Integer

Dim giClienteInicial As Integer

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1

'Property Variables:
Dim m_Caption As String
Event Unload()

Public Sub Padrao_Tela()

Dim iMes As Integer
Dim iAno As Integer
Dim dtData As Date
Dim dtData1 As Date

On Error GoTo Erro_Padrao_Tela

    iMes = Month(gdtDataAtual)
    iAno = Year(gdtDataAtual)

    dtData = CDate("01/" & iMes & "/" & iAno)
    dtData1 = DateAdd("d", -1, dtData)
    dtData = DateAdd("m", -1, dtData)

    DataFatDe.PromptInclude = False
    DataFatDe.Text = Format(dtData, "dd/mm/yy")
    DataFatDe.PromptInclude = True

    DataFatAte.PromptInclude = False
    DataFatAte.Text = Format(dtData1, "dd/mm/yy")
    DataFatAte.PromptInclude = True
    
    DataEmiDe.PromptInclude = False
    DataEmiDe.Text = Format(dtData, "dd/mm/yy")
    DataEmiDe.PromptInclude = True
    
    DataEmiAte.PromptInclude = False
    DataEmiAte.Text = Format(dtData1, "dd/mm/yy")
    DataEmiAte.PromptInclude = True
    
    Exit Sub

Erro_Padrao_Tela:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197260)

    End Select

    Exit Sub
    
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCliente = New AdmEvento

    Call Padrao_Tela
    
    Empresa.Caption = gsEmpresaTRV
    
    Call Combo_Seleciona_ItemData(Marca, 0)

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197260)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Private Sub BotaoExibir_Click(Index As Integer)

Dim colSelecao As New Collection
Dim sFiltro As String
Dim sNomeBrowse As String

    Select Case Index
    
        Case 0
            sFiltro = "DataEmiVou BETWEEN ? AND ? AND (Cliente >=? OR ?=0) AND (Cliente <=? OR ?=0)"
            sNomeBrowse = "TRVVouEstVendaLista"
            colSelecao.Add StrParaDate(DataEmiDe.Text)
            colSelecao.Add StrParaDate(DataEmiAte.Text)
        Case 1
            sFiltro = "DataEmiVou BETWEEN ? AND ? AND (Cliente >=? OR ?=0) AND (Cliente <=? OR ?=0) AND StatusCod = 7 AND NumFat = 0"
            sNomeBrowse = "TRVVouEstVendaLista"
            colSelecao.Add StrParaDate(DataEmiDe.Text)
            colSelecao.Add StrParaDate(DataEmiAte.Text)
        Case 2
            sFiltro = "DataEmissao BETWEEN ? AND ? AND (Cliente >=? OR ?=0) AND (Cliente <=? OR ?=0) AND TipoDocDestino = 3"
            sNomeBrowse = "TRVFaturasLista"
            colSelecao.Add StrParaDate(DataFatDe.Text)
            colSelecao.Add StrParaDate(DataFatAte.Text)
        Case 3
            sFiltro = "DataEmissao BETWEEN ? AND ? AND (Cliente >=? OR ?=0) AND (Cliente <=? OR ?=0) AND TipoDocDestino In (4,5)"
            sNomeBrowse = "TRVFaturasLista"
            colSelecao.Add StrParaDate(DataFatDe.Text)
            colSelecao.Add StrParaDate(DataFatAte.Text)
        
        Case 4
            sFiltro = "DataEmissao BETWEEN ? AND ? AND (Cliente >=? OR ?=0) AND (Cliente <=? OR ?=0) AND TipoDocDestino = 3 AND Status = 5"
            sNomeBrowse = "TRVFaturasLista"
            colSelecao.Add StrParaDate(DataFatDe.Text)
            colSelecao.Add StrParaDate(DataFatAte.Text)
        
        Case 5
            sFiltro = "DataEmissao BETWEEN ? AND ? AND (Cliente >=? OR ?=0) AND (Cliente <=? OR ?=0) AND TipoDocDestino In (4,5) AND Status = 5"
            sNomeBrowse = "TRVFaturasLista"
            colSelecao.Add StrParaDate(DataFatDe.Text)
            colSelecao.Add StrParaDate(DataFatAte.Text)
        
        Case 6
            sFiltro = "DataEmiVou BETWEEN ? AND ? AND DataEmiFat BETWEEN ? AND ? AND (Cliente >=? OR ?=0) AND (Cliente <=? OR ?=0) "
            sNomeBrowse = "TRVVouEstVendaLista"
            colSelecao.Add StrParaDate(DataEmiDe.Text)
            colSelecao.Add StrParaDate(DataEmiAte.Text)
            colSelecao.Add StrParaDate(DataFatDe.Text)
            colSelecao.Add StrParaDate(DataFatAte.Text)
            
    End Select

    colSelecao.Add LCodigo_Extrai(ClienteDe.Text)
    colSelecao.Add LCodigo_Extrai(ClienteDe.Text)
    colSelecao.Add LCodigo_Extrai(ClienteAte.Text)
    colSelecao.Add LCodigo_Extrai(ClienteAte.Text)

    Call Chama_Tela(sNomeBrowse, colSelecao, Nothing, Nothing, sFiltro)

End Sub

Private Sub LabelClienteAte_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As Collection

    giClienteInicial = 0
    
    If Len(Trim(ClienteAte.Text)) > 0 Then
        'Preenche com o cliente da tela
        objcliente.lCodigo = LCodigo_Extrai(ClienteAte.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoCliente)

End Sub

Private Sub LabelClienteDe_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As Collection

    giClienteInicial = 1

    If Len(Trim(ClienteDe.Text)) > 0 Then
        'Preenche com o cliente da tela
        objcliente.lCodigo = LCodigo_Extrai(ClienteDe.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoCliente)

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente

    Set objcliente = obj1
    
    'Preenche campo Cliente
    If giClienteInicial = 1 Then
        ClienteDe.Text = CStr(objcliente.lCodigo)
        Call ClienteDe_Validate(bSGECancelDummy)
    Else
        ClienteAte.Text = CStr(objcliente.lCodigo)
        Call ClienteAte_Validate(bSGECancelDummy)
    End If

    Me.Show

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoCliente = Nothing

    'Fecha o Comando de Setas
    Call ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_Activate()
   'Call TelaIndice_Preenche(Me)
End Sub

Public Sub Form_Deactivate()
    'gi_ST_SetaIgnoraClick = 1
End Sub

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Estatística de Vendas"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TRVEstVenda"

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

Private Sub BotaoConsultar_Click()

Dim lErro As Long
Dim objEstVenda As New ClassTRVEstVenda
Dim iMarca As Integer

On Error GoTo Erro_BotaoConsultar_Click

    If StrParaDate(DataEmiDe.Text) = DATA_NULA Then gError 200863
    If StrParaDate(DataEmiAte.Text) = DATA_NULA Then gError 200864
    If StrParaDate(DataFatDe.Text) = DATA_NULA Then gError 200865
    If StrParaDate(DataFatAte.Text) = DATA_NULA Then gError 200866

    If StrParaDate(DataEmiDe.Text) > StrParaDate(DataEmiAte.Text) Then gError 200867
    If StrParaDate(DataFatDe.Text) > StrParaDate(DataFatAte.Text) Then gError 200868
    
    objEstVenda.dtDataEmiAte = StrParaDate(DataEmiAte.Text)
    objEstVenda.dtDataEmiDe = StrParaDate(DataEmiDe.Text)
    objEstVenda.dtDataFatAte = StrParaDate(DataFatAte.Text)
    objEstVenda.dtDataFatDe = StrParaDate(DataFatDe.Text)
    
    iMarca = Marca.ItemData(Marca.ListIndex)
    
    If iMarca = TRV_EMPRESA_AMBOS Then
        objEstVenda.lClienteDe = LCodigo_Extrai(ClienteDe.Text)
        objEstVenda.lClienteAte = LCodigo_Extrai(ClienteAte.Text)
    ElseIf iMarca = TRV_EMPRESA_MY Then
        objEstVenda.lClienteDe = LCodigo_Extrai(ClienteDe.Text)
        objEstVenda.lClienteAte = LCodigo_Extrai(ClienteAte.Text)
        
        If objEstVenda.lClienteDe < TRV_CLIENTE_FAIXA_MY_DE Then objEstVenda.lClienteDe = TRV_CLIENTE_FAIXA_MY_DE
        If (objEstVenda.lClienteAte < TRV_CLIENTE_FAIXA_MY_DE) And objEstVenda.lClienteAte <> 0 Then objEstVenda.lClienteAte = -1
    Else
        objEstVenda.lClienteDe = LCodigo_Extrai(ClienteDe.Text)
        objEstVenda.lClienteAte = LCodigo_Extrai(ClienteAte.Text)
    
        If objEstVenda.lClienteDe >= TRV_CLIENTE_FAIXA_MY_DE Then objEstVenda.lClienteDe = 9999999
        If (objEstVenda.lClienteAte >= TRV_CLIENTE_FAIXA_MY_DE) Or objEstVenda.lClienteAte = 0 Then objEstVenda.lClienteAte = TRV_CLIENTE_FAIXA_MY_DE - 1
    End If
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = CF("TRVEstVenda_Le", objEstVenda)
    If lErro <> SUCESSO Then gError 200869
    
    ValorBrutoVouRS.Text = Format(objEstVenda.dValorBrutoVouRS, "STANDARD")
    ValorBrutoVouRSCanc.Text = Format(objEstVenda.dValorBrutoVouRSCanc, "STANDARD")
    ValorBrutoVouRSFat.Text = Format(objEstVenda.dValorBrutoVouRSFat, "STANDARD")
    ValorBrutoVouUSS.Text = Format(objEstVenda.dValorBrutoVouUSS, "STANDARD")
    ValorBrutoVouUSSCanc.Text = Format(objEstVenda.dValorBrutoVouUSSCanc, "STANDARD")
    ValorBrutoVouUSSFat.Text = Format(objEstVenda.dValorBrutoVouUSSFat, "STANDARD")
    ValorFatVou.Text = Format(objEstVenda.dValorFatVou, "STANDARD")
    ValorFatVouCanc.Text = Format(objEstVenda.dValorFatVouCanc, "STANDARD")
    ValorFatVouFat.Text = Format(objEstVenda.dValorFatVouFat, "STANDARD")
    QtdeEmiVou.Text = Format(objEstVenda.dQtdeEmiVou, "##0")
    QtdeEmiVouCanc.Text = Format(objEstVenda.dQtdeEmiVouCanc, "##0")
    QtdeEmiVouFat.Text = Format(objEstVenda.dQtdeEmiVouFat, "##0")
    QtdeEmiPax.Text = Format(objEstVenda.dQtdeEmiPax, "##0")
    QtdeEmiPaxCanc.Text = Format(objEstVenda.dQtdeEmiPaxCanc, "##0")
    QtdeEmiPaxFat.Text = Format(objEstVenda.dQtdeEmiPaxFat, "##0")
    ValorFatCR.Text = Format(objEstVenda.dValorFatCR, "STANDARD")
    ValorFatCP.Text = Format(objEstVenda.dValorFatCP, "STANDARD")
    QtdeFatCR.Text = Format(objEstVenda.dQtdeFatCR, "##0")
    QtdeFatCP.Text = Format(objEstVenda.dQtdeFatCP, "##0")
    ValorFatCRCanc.Text = Format(objEstVenda.dValorFatCRCanc, "STANDARD")
    ValorFatCPCanc.Text = Format(objEstVenda.dValorFatCPCanc, "STANDARD")
    QtdeFatCRCanc.Text = Format(objEstVenda.dQtdeFatCRCanc, "##0")
    QtdeFatCPCanc.Text = Format(objEstVenda.dQtdeFatCPCanc, "##0")
        
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoConsultar_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 200863
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INIC_NAO_PREENCHIDA", gErr)
        
        Case 200864
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_FIM_NAO_PREENCHIDA", gErr)
        
        Case 200865
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INIC_NAO_PREENCHIDA", gErr)
        
        Case 200866
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_FIM_NAO_PREENCHIDA", gErr)
        
        Case 200867
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
        
        Case 200868
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
        
        Case 200869
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200870)

    End Select

    Exit Sub

End Sub

Private Sub DataEmiDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataEmiDe_GotFocus()
     Call MaskEdBox_TrataGotFocus(DataEmiDe, iAlterado)
End Sub

Private Sub DataEmiDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmiDe_Validate

    'Verifica se a Data de Emi foi digitada
    If Len(Trim(DataEmiDe.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataEmiDe.Text)
    If lErro <> SUCESSO Then gError 197261

    Exit Sub

Erro_DataEmiDe_Validate:

    Cancel = True

    Select Case gErr

        Case 197261

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197262)

    End Select

    Exit Sub

End Sub

Private Sub DataEmiAte_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEmiAte_GotFocus()

     Call MaskEdBox_TrataGotFocus(DataEmiAte, iAlterado)

End Sub

Private Sub DataEmiAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmiAte_Validate

    'Verifica se a Data de Emi foi digitada
    If Len(Trim(DataEmiAte.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataEmiAte.Text)
    If lErro <> SUCESSO Then gError 197263

    Exit Sub

Erro_DataEmiAte_Validate:

    Cancel = True

    Select Case gErr

        Case 197263

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197264)

    End Select

    Exit Sub

End Sub

Private Sub DataFatDe_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataFatDe_GotFocus()

     Call MaskEdBox_TrataGotFocus(DataFatDe, iAlterado)

End Sub

Private Sub DataFatDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataFatDe_Validate

    'Verifica se a Data de Fat foi digitada
    If Len(Trim(DataFatDe.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataFatDe.Text)
    If lErro <> SUCESSO Then gError 197265

    Exit Sub

Erro_DataFatDe_Validate:

    Cancel = True

    Select Case gErr

        Case 197265

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197266)

    End Select

    Exit Sub

End Sub

Private Sub DataFatAte_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataFatAte_GotFocus()
     Call MaskEdBox_TrataGotFocus(DataFatAte, iAlterado)
End Sub

Private Sub DataFatAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataFatAte_Validate

    'Verifica se a Data de Fat foi digitada
    If Len(Trim(DataFatAte.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataFatAte.Text)
    If lErro <> SUCESSO Then gError 197267

    Exit Sub

Erro_DataFatAte_Validate:

    Cancel = True

    Select Case gErr

        Case 197267

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197268)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmiDe_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmiDe_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataEmiDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 197269

    Exit Sub

Erro_UpDownEmiDe_DownClick:

    Select Case gErr

        Case 197269

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197270)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmiDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmiDe_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEmiDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 197271

    Exit Sub

Erro_UpDownEmiDe_UpClick:

    Select Case gErr

        Case 197271

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197272)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmiAte_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmiAte_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataEmiAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 197277

    Exit Sub

Erro_UpDownEmiAte_DownClick:

    Select Case gErr

        Case 197277

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197278)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmiAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmiAte_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEmiAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 197279

    Exit Sub

Erro_UpDownEmiAte_UpClick:

    Select Case gErr

        Case 197279

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197280)

    End Select

    Exit Sub

End Sub

Private Sub UpDownFatDe_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownFatDe_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataFatDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 197273

    Exit Sub

Erro_UpDownFatDe_DownClick:

    Select Case gErr

        Case 197273

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197274)

    End Select

    Exit Sub

End Sub

Private Sub UpDownFatDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownFatDe_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataFatDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 197275

    Exit Sub

Erro_UpDownFatDe_UpClick:

    Select Case gErr

        Case 197275

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197276)

    End Select

    Exit Sub

End Sub

Private Sub UpDownFatAte_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownFatAte_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataFatAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 197281

    Exit Sub

Erro_UpDownFatAte_DownClick:

    Select Case gErr

        Case 197281

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197282)

    End Select

    Exit Sub

End Sub

Private Sub UpDownFatAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownFatAte_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataFatAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 197283

    Exit Sub

Erro_UpDownFatAte_UpClick:

    Select Case gErr

        Case 197283

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197284)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub


Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    Call Limpa_Tela_TRVEstVenda

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 197299

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197300)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_TRVEstVenda()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_TRVEstVenda

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    Call Limpa_Tela(Me)
    
    Call Combo_Seleciona_ItemData(Marca, 0)
          
'    ValorBrutoVouRS.Text = ""
'    ValorBrutoVouRSCanc.Text = ""
'    ValorBrutoVouRSFat.Text = ""
'    ValorBrutoVouUSS.Text = ""
'    ValorBrutoVouUSSCanc.Text = ""
'    ValorBrutoVouUSSFat.Text = ""
'    ValorFatVou.Text = ""
'    ValorFatVouCanc.Text = ""
'    ValorFatVouFat.Text = ""
'    QtdeEmiVou.Text = ""
'    QtdeEmiVouCanc.Text = ""
'    QtdeEmiVouFat.Text = ""
'    QtdeEmiPax.Text = ""
'    QtdeEmiPaxCanc.Text = ""
'    QtdeEmiPaxFat.Text = ""
'    ValorFatCR.Text = ""
'    ValorFatCP.Text = ""
'    QtdeFatCR.Text = ""
'    QtdeFatCP.Text = ""
'    ValorFatCRCanc.Text = ""
'    ValorFatCPCanc.Text = ""
'    QtdeFatCRCanc.Text = ""
'    QtdeFatCPCanc.Text = ""
    
    Call Padrao_Tela
    
    iAlterado = 0
 
    Exit Sub

Erro_Limpa_Tela_TRVEstVenda:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197301)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    'Call Tela_QueryUnload(Me, iAlterado, UnloadMode, Cancel, iTelaCorrenteAtiva)
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is ClienteDe Then
            Call LabelClienteDe_Click
        ElseIf Me.ActiveControl Is ClienteAte Then
            Call LabelClienteAte_Click
        End If
          
    End If

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

'***** fim do trecho a ser copiado ******

Private Sub ClienteDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objcliente As New ClassCliente

On Error GoTo Erro_ClienteDe_Validate

    If Len(Trim(ClienteDe.Text)) > 0 Then
   
        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteDe, objcliente, 0)
        If lErro <> SUCESSO Then Error 37793

    End If
    
    giClienteInicial = 1
    
    Exit Sub

Erro_ClienteDe_Validate:

    Cancel = True


    Select Case Err

        Case 37793
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO_2", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168897)

    End Select

End Sub

Private Sub ClienteAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objcliente As New ClassCliente

On Error GoTo Erro_ClienteAte_Validate

    If Len(Trim(ClienteAte.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteAte, objcliente, 0)
        If lErro <> SUCESSO Then Error 37794

    End If
    
    giClienteInicial = 0
 
    Exit Sub

Erro_ClienteAte_Validate:

    Cancel = True


    Select Case Err

        Case 37794
             lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", Err, objcliente.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168898)

    End Select

End Sub

