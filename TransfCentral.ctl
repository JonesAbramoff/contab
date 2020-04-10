VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl TransfCentral 
   ClientHeight    =   6060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9060
   KeyPreview      =   -1  'True
   ScaleHeight     =   6060
   ScaleWidth      =   9060
   Begin VB.CommandButton BotaoProxNum 
      Height          =   300
      Left            =   2370
      Picture         =   "TransfCentral.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Numeração Automática"
      Top             =   345
      Width           =   300
   End
   Begin VB.Frame Frame1 
      Caption         =   "De"
      Height          =   5175
      Left            =   105
      TabIndex        =   54
      Top             =   810
      Width           =   4380
      Begin VB.Frame FrameChqDe 
         BorderStyle     =   0  'None
         Height          =   4380
         Left            =   105
         TabIndex        =   55
         Top             =   645
         Visible         =   0   'False
         Width           =   4230
         Begin MSMask.MaskEdBox SeqChqDe 
            Height          =   300
            Left            =   1890
            TabIndex        =   4
            Top             =   82
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#########"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   7
            Left            =   1265
            TabIndex        =   80
            Top             =   4088
            Width           =   510
         End
         Begin VB.Label ValorChqDe 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1890
            TabIndex        =   14
            Top             =   4035
            Width           =   1260
         End
         Begin VB.Label CarneChqDe 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1890
            TabIndex        =   12
            Top             =   3242
            Width           =   1935
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Carnê:"
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
            Index           =   17
            Left            =   1205
            TabIndex        =   79
            Top             =   3295
            Width           =   570
         End
         Begin VB.Label DataBomParaChqDe 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1890
            TabIndex        =   13
            Top             =   3637
            Width           =   1170
         End
         Begin VB.Label CupomFiscalChqDe 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1875
            TabIndex        =   11
            Top             =   2850
            Width           =   1110
         End
         Begin VB.Label ECFChqDe 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1905
            TabIndex        =   10
            Top             =   2445
            Width           =   615
         End
         Begin VB.Label ClienteChqDe 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1890
            TabIndex        =   9
            Top             =   2057
            Width           =   1680
         End
         Begin VB.Label NumeroChqDe 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1890
            TabIndex        =   8
            Top             =   1662
            Width           =   1395
         End
         Begin VB.Label ContaChqDe 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1890
            TabIndex        =   7
            Top             =   1267
            Width           =   1395
         End
         Begin VB.Label AgenciaChqDe 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1890
            TabIndex        =   6
            Top             =   872
            Width           =   870
         End
         Begin VB.Label BancoChqDe 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1890
            TabIndex        =   5
            Top             =   477
            Width           =   870
         End
         Begin VB.Label LabelSeqChqDe 
            AutoSize        =   -1  'True
            Caption         =   "Sequencial:"
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
            Left            =   755
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   78
            Top             =   135
            Width           =   1020
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "ECF:"
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
            Index           =   20
            Left            =   1355
            TabIndex        =   77
            Top             =   2505
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
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
            Left            =   1160
            TabIndex        =   62
            Top             =   530
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Agência:"
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
            Index           =   3
            Left            =   1010
            TabIndex        =   61
            Top             =   925
            Width           =   765
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   5
            Left            =   1205
            TabIndex        =   60
            Top             =   1320
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   1055
            TabIndex        =   59
            Top             =   1715
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Bom Para:"
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
            Left            =   890
            TabIndex        =   58
            Top             =   3690
            Width           =   885
         End
         Begin VB.Label Label1 
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
            Height          =   195
            Index           =   4
            Left            =   1115
            TabIndex        =   57
            Top             =   2110
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cupom Fiscal (COO):"
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
            Index           =   14
            Left            =   5
            TabIndex        =   56
            Top             =   2900
            Width           =   1770
         End
      End
      Begin VB.Frame FrameOutDe 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   4380
         Left            =   105
         TabIndex        =   63
         Top             =   645
         Visible         =   0   'False
         Width           =   4230
         Begin VB.ComboBox ParcelamentoOutDe 
            Height          =   315
            ItemData        =   "TransfCentral.ctx":00EA
            Left            =   1860
            List            =   "TransfCentral.ctx":00EC
            TabIndex        =   97
            Top             =   555
            Width           =   2250
         End
         Begin VB.ComboBox AdmOutDe 
            Height          =   315
            ItemData        =   "TransfCentral.ctx":00EE
            Left            =   1860
            List            =   "TransfCentral.ctx":00F0
            TabIndex        =   19
            Top             =   135
            Width           =   2250
         End
         Begin MSMask.MaskEdBox ValorOutDe 
            Height          =   300
            Left            =   1860
            TabIndex        =   20
            Top             =   975
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
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
            Format          =   "Standard"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Parcelamento:"
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
            Index           =   23
            Left            =   540
            TabIndex        =   98
            Top             =   600
            Width           =   1230
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Meio Pagto:"
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
            Index           =   28
            Left            =   735
            TabIndex        =   65
            Top             =   180
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   31
            Left            =   1245
            TabIndex        =   64
            Top             =   1020
            Width           =   510
         End
      End
      Begin VB.Frame FrameTktDe 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   4380
         Left            =   105
         TabIndex        =   66
         Top             =   645
         Visible         =   0   'False
         Width           =   4230
         Begin VB.ComboBox ParcelamentoTktDe 
            Height          =   315
            ItemData        =   "TransfCentral.ctx":00F2
            Left            =   1875
            List            =   "TransfCentral.ctx":00F4
            TabIndex        =   95
            Top             =   570
            Width           =   2250
         End
         Begin VB.ComboBox AdmTktDe 
            Height          =   315
            ItemData        =   "TransfCentral.ctx":00F6
            Left            =   1890
            List            =   "TransfCentral.ctx":00F8
            TabIndex        =   21
            Top             =   135
            Width           =   2250
         End
         Begin MSMask.MaskEdBox ValorTktDe 
            Height          =   300
            Left            =   1860
            TabIndex        =   22
            Top             =   1020
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
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
            Format          =   "Standard"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Parcelamento:"
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
            Index           =   22
            Left            =   555
            TabIndex        =   96
            Top             =   615
            Width           =   1230
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   26
            Left            =   1290
            TabIndex        =   68
            Top             =   1065
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ticket:"
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
            Index           =   27
            Left            =   1170
            TabIndex        =   67
            Top             =   180
            Width           =   615
         End
      End
      Begin VB.Frame FrameCrtDe 
         BorderStyle     =   0  'None
         Caption         =   "FrameCrtDe"
         Height          =   4380
         Left            =   90
         TabIndex        =   69
         Top             =   645
         Visible         =   0   'False
         Width           =   4230
         Begin VB.OptionButton OptionManualDe 
            Caption         =   "&Manual"
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
            Left            =   1905
            TabIndex        =   90
            Top             =   1545
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton OptionPOSDe 
            Caption         =   "&POS"
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
            Left            =   3045
            TabIndex        =   89
            Top             =   1575
            Width           =   735
         End
         Begin VB.ComboBox AdmCrtDe 
            Height          =   315
            ItemData        =   "TransfCentral.ctx":00FA
            Left            =   1890
            List            =   "TransfCentral.ctx":00FC
            TabIndex        =   15
            Top             =   180
            Width           =   2340
         End
         Begin VB.ComboBox ParcelamentoCrtDe 
            Height          =   315
            ItemData        =   "TransfCentral.ctx":00FE
            Left            =   1860
            List            =   "TransfCentral.ctx":0100
            TabIndex        =   16
            Top             =   585
            Width           =   2340
         End
         Begin MSMask.MaskEdBox ValorCrtDe 
            Height          =   300
            Left            =   1860
            TabIndex        =   17
            Top             =   1050
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
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
            Format          =   "Standard"
            PromptChar      =   " "
         End
         Begin VB.Label LabelTerminalDe 
            Caption         =   "Terminal:"
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
            Left            =   990
            TabIndex        =   91
            Top             =   1530
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   18
            Left            =   1170
            TabIndex        =   72
            Top             =   195
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   21
            Left            =   1290
            TabIndex        =   71
            Top             =   1080
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Parcelamento:"
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
            Index           =   19
            Left            =   570
            TabIndex        =   70
            Top             =   630
            Width           =   1230
         End
      End
      Begin VB.Frame FrameDinDe 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   4380
         Left            =   105
         TabIndex        =   73
         Top             =   645
         Visible         =   0   'False
         Width           =   4230
         Begin MSMask.MaskEdBox ValorDinDe 
            Height          =   300
            Left            =   1875
            TabIndex        =   18
            Top             =   150
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
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
            Format          =   "Standard"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   16
            Left            =   1275
            TabIndex        =   74
            Top             =   195
            Width           =   510
         End
      End
      Begin VB.ComboBox TipoMeioPagtoDe 
         Height          =   315
         ItemData        =   "TransfCentral.ctx":0102
         Left            =   1980
         List            =   "TransfCentral.ctx":0104
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Tipo do meio de pagamento"
         Top             =   315
         Width           =   2220
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Meio Pagto:"
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
         Left            =   420
         TabIndex        =   75
         ToolTipText     =   "Tipo do meio de pagamento"
         Top             =   375
         Width           =   1470
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Para"
      Height          =   5175
      Left            =   4545
      TabIndex        =   40
      Top             =   810
      Width           =   4380
      Begin VB.Frame FrameChqPara 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   4260
         Left            =   105
         TabIndex        =   45
         Top             =   720
         Visible         =   0   'False
         Width           =   4215
         Begin MSMask.MaskEdBox ECFChqPara 
            Height          =   315
            Left            =   1845
            TabIndex        =   29
            Top             =   2273
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin VB.CommandButton BotaoLerChq 
            Caption         =   "Ler"
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
            Left            =   1860
            TabIndex        =   33
            Top             =   3435
            Width           =   1350
         End
         Begin MSMask.MaskEdBox AgenciaChqPara 
            Height          =   300
            Left            =   1850
            TabIndex        =   25
            Top             =   375
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   7
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ContaChqPara 
            Height          =   300
            Left            =   1850
            TabIndex        =   26
            Top             =   750
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataBomParaChqPara 
            Height          =   300
            Left            =   1845
            TabIndex        =   30
            Top             =   1897
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ClienteChqPara 
            Height          =   315
            Left            =   1850
            TabIndex        =   28
            ToolTipText     =   "CGC/CPF do Cliente"
            Top             =   1500
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   14
            Mask            =   "##############"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NumeroChqPara 
            Height          =   300
            Left            =   1850
            TabIndex        =   27
            Top             =   1125
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox BancoChqPara 
            Height          =   300
            Left            =   1845
            TabIndex        =   24
            Top             =   0
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   3
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CupomFiscalChqPara 
            Height          =   300
            Left            =   1850
            TabIndex        =   32
            Top             =   2646
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CarneChqPara 
            Height          =   300
            Left            =   1850
            TabIndex        =   31
            Top             =   3030
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   20
            Mask            =   "99999999999999999999"
            PromptChar      =   " "
         End
         Begin VB.Label LabelECFChqPara 
            AutoSize        =   -1  'True
            Caption         =   "ECF:"
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
            Left            =   1335
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   83
            Top             =   2333
            Width           =   420
         End
         Begin VB.Label LabelCarneChqPara 
            AutoSize        =   -1  'True
            Caption         =   "Carnê:"
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
            Left            =   1185
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   82
            Top             =   3083
            Width           =   570
         End
         Begin VB.Label LabelCupomFiscalChqPara 
            AutoSize        =   -1  'True
            Caption         =   "Cupom Fiscal (COO):"
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
            Left            =   -15
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   81
            Top             =   2700
            Width           =   1770
         End
         Begin VB.Label Label1 
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
            Height          =   195
            Index           =   8
            Left            =   1095
            TabIndex        =   51
            Top             =   1565
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Bom Para:"
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
            Index           =   9
            Left            =   870
            TabIndex        =   50
            Top             =   1950
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   10
            Left            =   1035
            TabIndex        =   49
            Top             =   1187
            Width           =   720
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   11
            Left            =   1185
            TabIndex        =   48
            Top             =   809
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Agência:"
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
            Index           =   12
            Left            =   990
            TabIndex        =   47
            Top             =   431
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
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
            Left            =   1140
            TabIndex        =   46
            Top             =   53
            Width           =   615
         End
      End
      Begin VB.Frame FrameCrtPara 
         BorderStyle     =   0  'None
         Caption         =   "FrameCrtDe"
         Height          =   4485
         Left            =   75
         TabIndex        =   84
         Top             =   645
         Visible         =   0   'False
         Width           =   4215
         Begin VB.OptionButton OptionPOSPara 
            Caption         =   "&POS"
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
            Left            =   3060
            TabIndex        =   93
            Top             =   1080
            Width           =   735
         End
         Begin VB.OptionButton OptionManualPara 
            Caption         =   "&Manual"
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
            Left            =   1890
            TabIndex        =   92
            Top             =   1065
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.ComboBox AdmCrtPara 
            Height          =   315
            ItemData        =   "TransfCentral.ctx":0106
            Left            =   1860
            List            =   "TransfCentral.ctx":0108
            TabIndex        =   86
            Top             =   165
            Width           =   2340
         End
         Begin VB.ComboBox ParcelamentoCrtPara 
            Height          =   315
            ItemData        =   "TransfCentral.ctx":010A
            Left            =   1860
            List            =   "TransfCentral.ctx":010C
            TabIndex        =   85
            Top             =   585
            Width           =   2340
         End
         Begin VB.Label LabelTerminalPara 
            Caption         =   "Terminal:"
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
            Left            =   990
            TabIndex        =   94
            Top             =   1065
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   34
            Left            =   1170
            TabIndex        =   88
            Top             =   195
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Parcelamento:"
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
            Left            =   570
            TabIndex        =   87
            Top             =   630
            Width           =   1230
         End
      End
      Begin VB.Frame FrameTktPara 
         BorderStyle     =   0  'None
         Caption         =   "FrameCrtDe"
         Height          =   4485
         Left            =   75
         TabIndex        =   43
         Top             =   630
         Visible         =   0   'False
         Width           =   4215
         Begin VB.ComboBox AdmTktPara 
            Height          =   315
            ItemData        =   "TransfCentral.ctx":010E
            Left            =   1860
            List            =   "TransfCentral.ctx":0110
            TabIndex        =   35
            Top             =   135
            Width           =   2220
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ticket:"
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
            Index           =   29
            Left            =   1170
            TabIndex        =   44
            Top             =   180
            Width           =   615
         End
      End
      Begin VB.Frame FrameOutPara 
         BorderStyle     =   0  'None
         Caption         =   "FrameCrtDe"
         Height          =   4485
         Left            =   105
         TabIndex        =   41
         Top             =   630
         Visible         =   0   'False
         Width           =   4215
         Begin VB.ComboBox AdmOutPara 
            Height          =   315
            ItemData        =   "TransfCentral.ctx":0112
            Left            =   1875
            List            =   "TransfCentral.ctx":0114
            TabIndex        =   34
            Top             =   135
            Width           =   2220
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Meio Pagto:"
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
            Index           =   33
            Left            =   735
            TabIndex        =   42
            Top             =   180
            Width           =   1035
         End
      End
      Begin VB.Frame FrameDinPara 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   4485
         Left            =   75
         TabIndex        =   52
         Top             =   630
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.ComboBox TipoMeioPagtoPara 
         Height          =   315
         ItemData        =   "TransfCentral.ctx":0116
         Left            =   1935
         List            =   "TransfCentral.ctx":0118
         Style           =   2  'Dropdown List
         TabIndex        =   23
         ToolTipText     =   "Tipo do meio de pagamento"
         Top             =   315
         Width           =   2220
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Meio Pagto:"
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
         Index           =   15
         Left            =   390
         TabIndex        =   53
         ToolTipText     =   "Tipo do meio de pagamento"
         Top             =   375
         Width           =   1470
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6765
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   150
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "TransfCentral.ctx":011A
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1110
         Picture         =   "TransfCentral.ctx":0298
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "TransfCentral.ctx":07CA
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   75
         Picture         =   "TransfCentral.ctx":0954
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   300
      Left            =   1290
      TabIndex        =   1
      Top             =   345
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   9
      Mask            =   "#########"
      PromptChar      =   " "
   End
   Begin VB.Label LabelCodigo 
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
      Left            =   525
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   76
      Top             =   375
      Width           =   660
   End
End
Attribute VB_Name = "TransfCentral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Eventos:
Event Unload()

Dim m_Caption As String

'Variáveis globais
Public iAlterado As Integer
Private WithEvents objEventoCheque As AdmEvento
Attribute objEventoCheque.VB_VarHelpID = -1
Private WithEvents objEventoTransf As AdmEvento
Attribute objEventoTransf.VB_VarHelpID = -1
Private WithEvents objEventoCupomFiscal As AdmEvento
Attribute objEventoCupomFiscal.VB_VarHelpID = -1

Dim giAdmCrtPara As Integer
Dim giAdmCrtDe As Integer
Dim giAdmTktDe As Integer
Dim giAdmOutDe As Integer
'constantes
Const TRANSFERENCIA_LOJA = 1

Private Sub Form_Load()
'Inicializacao da tela

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Carrega as combos de Tipos de Pagto (Din, Cheque, Cartão, etc)
    lErro = Carrega_TipoMeioPagto()
    If lErro <> SUCESSO Then gError 113567
    
    'carrega as combos de admmeiopagto que não são de cartão
    lErro = Carrega_AdmMeioPagto_Nao_Cartao()
    If lErro <> SUCESSO Then gError 113568
    
    Set objEventoTransf = New AdmEvento
    Set objEventoCheque = New AdmEvento
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 113567, 113568

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175410)

    End Select

    Exit Sub

End Sub

Private Function Carrega_TipoMeioPagto() As Long

Dim lErro As Long
Dim colTiposMeiosPagto As New Collection
Dim objTipoMeioPagto As ClassTMPLoja

On Error GoTo Erro_Carrega_TipoMeioPagto

    lErro = CF("TipoMeioPagto_Le_Todas", colTiposMeiosPagto)
    If lErro <> SUCESSO And lErro <> 104036 Then gError 113513
    
    'se não encontrar nenhum tipoMeioPagto-> erro
    If lErro = 104036 Then gError 113514
    
    'para cada elemento da coleção
    For Each objTipoMeioPagto In colTiposMeiosPagto
    
        'se for passível de transferência
        If objTipoMeioPagto.iTransferencia = TRANSFERENCIA_LOJA Then
            
            'preenche a combo "De"
            TipoMeioPagtoDe.AddItem (objTipoMeioPagto.iTipo & SEPARADOR & objTipoMeioPagto.sDescricao)
            TipoMeioPagtoDe.ItemData(TipoMeioPagtoDe.NewIndex) = objTipoMeioPagto.iTipo
            
            'preenche a combo "Para"
            TipoMeioPagtoPara.AddItem (objTipoMeioPagto.iTipo & SEPARADOR & objTipoMeioPagto.sDescricao)
            TipoMeioPagtoPara.ItemData(TipoMeioPagtoPara.NewIndex) = objTipoMeioPagto.iTipo
        
        End If
    
    Next
    
    Carrega_TipoMeioPagto = SUCESSO

    Exit Function
    
Erro_Carrega_TipoMeioPagto:

    Carrega_TipoMeioPagto = gErr

    Select Case gErr
    
        Case 113513
        
        Case 113514
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_VAZIA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175411)

    End Select

End Function

Private Function Carrega_AdmMeioPagto_Nao_Cartao() As Long

Dim colAdmMeioPagto As New Collection
Dim objAdmMeioPagto As ClassAdmMeioPagto
Dim lErro As Long

On Error GoTo Erro_Carrega_AdmMeioPagto_Nao_Cartao

    'lê todas as admMeiopagto
    lErro = CF("AdmMeioPagto_Le_Todas", colAdmMeioPagto)
    If lErro <> SUCESSO And lErro <> 104031 Then gError 113515
    
    If lErro = 104031 Then gError 113516
    
    For Each objAdmMeioPagto In colAdmMeioPagto
    
        'verifica o tipo de meio de pagaamento do admmeiopagto
        Select Case objAdmMeioPagto.iTipoMeioPagto
        
            'se for ticket
            Case TIPOMEIOPAGTOLOJA_VALE_TICKET
                AdmTktDe.AddItem (objAdmMeioPagto.iCodigo & SEPARADOR & objAdmMeioPagto.sNome)
                AdmTktDe.ItemData(AdmTktDe.NewIndex) = objAdmMeioPagto.iCodigo
                
                AdmTktPara.AddItem (objAdmMeioPagto.iCodigo & SEPARADOR & objAdmMeioPagto.sNome)
                AdmTktPara.ItemData(AdmTktPara.NewIndex) = objAdmMeioPagto.iCodigo
            
            'se for outros
            Case TIPOMEIOPAGTOLOJA_OUTROS
                AdmOutDe.AddItem (objAdmMeioPagto.iCodigo & SEPARADOR & objAdmMeioPagto.sNome)
                AdmOutDe.ItemData(AdmOutDe.NewIndex) = objAdmMeioPagto.iCodigo
                
                AdmOutPara.AddItem (objAdmMeioPagto.iCodigo & SEPARADOR & objAdmMeioPagto.sNome)
                AdmOutPara.ItemData(AdmOutPara.NewIndex) = objAdmMeioPagto.iCodigo
        
        End Select
    
    Next
    
    Carrega_AdmMeioPagto_Nao_Cartao = SUCESSO
    
    Exit Function
    
Erro_Carrega_AdmMeioPagto_Nao_Cartao:

    Carrega_AdmMeioPagto_Nao_Cartao = gErr
    
    Select Case gErr
    
        Case 113515
        
        Case 113516
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_VAZIA", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175412)
    
    End Select
    
    Exit Function

End Function

Public Function Trata_Parametros(Optional objTransfBrowser As ClassTransfBrowser) As Long

Dim lErro As Long
Dim objTransfCaixa As New ClassTransfCaixa

On Error GoTo Erro_Trata_Parametros

    'se objTransfCaixa está instanciado
    If Not objTransfBrowser Is Nothing Then
    
        Codigo.Text = CStr(objTransfBrowser.lTransferencia)
    
        objTransfCaixa.iFilialEmpresa = objTransfBrowser.iFilialEmpresa
        objTransfCaixa.lCodigo = objTransfBrowser.lTransferencia
    
        lErro = Traz_TransfCaixa_Tela(objTransfCaixa)
        If lErro <> SUCESSO And lErro <> 105277 Then gError 113569
    
    End If

    iAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:
    
    Trata_Parametros = gErr
    
    Select Case gErr
        
        Case 113569
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175413)
    
    End Select
    
    Exit Function

End Function

Private Function Traz_TransfCaixa_Tela(objTransfCaixa As ClassTransfCaixa) As Long
        
Dim lErro As Long

On Error GoTo Erro_Traz_TransfCaixa_Tela

    Call Limpa_Tela_TransfCentral
        
    lErro = CF("TransferenciaLoja_Le", objTransfCaixa)
    If lErro <> SUCESSO And lErro <> 105235 Then gError 105276
        
    'se nao encontrou a transferencia ==> erro
    If lErro = 105235 Then gError 105277
        
    'lê o Movto
    lErro = CF("MovimentosCaixa_Le_Transf", objTransfCaixa)
    If lErro <> SUCESSO And lErro <> 101994 And lErro <> 101997 Then gError 113570
    
    'se nao achou o movimento de origem
    If lErro = 101994 Then gError 113571

    'se nao achou o movimento de destino
    If lErro = 101997 Then gError 113572
    
    Codigo.Text = CStr(objTransfCaixa.lCodigo)
    
    'preenche o frame DE
    Call Preenche_Frame_De(objTransfCaixa.objMovCaixaDe)
    
    'preenche o frame PARA
    Call Preenche_Frame_Para(objTransfCaixa.objMovCaixaPara)
    
    Traz_TransfCaixa_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_TransfCaixa_Tela:
    
    Traz_TransfCaixa_Tela = gErr
    
    Select Case gErr
    
        Case 105276, 105277, 113570
        
        Case 113571, 113572
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSFERENCIALOJA_NAOCADASTRADA", gErr, objTransfCaixa.iFilialEmpresa, objTransfCaixa.lCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175414)
    
    End Select
    
    Exit Function

End Function
    
Private Function Preenche_Frame_De(objMovimentoCaixa As ClassMovimentoCaixa) As Long

Dim objTipoMovtoCaixa As New ClassTipoMovtoCaixa
Dim iIndice As Integer
Dim lErro As Long
Dim objChequePre As New ClassChequePre
Dim objCupomFiscal As New ClassCupomFiscal

On Error GoTo Erro_Preenche_Frame_De

    'preenche o codigo de um movtocaixa
    objTipoMovtoCaixa.iCodigo = objMovimentoCaixa.iTipo
    
    'verifica qual o tmpLoja desse movtoCaixa
    lErro = CF("TiposMovtoCaixa_Le_Codigo", objTipoMovtoCaixa)
    If lErro <> SUCESSO Then gError 113600
    
    'seleciona na combo "DE"
    For iIndice = 0 To TipoMeioPagtoDe.ListCount - 1
        
        If TipoMeioPagtoDe.ItemData(iIndice) = objTipoMovtoCaixa.iTMPLoja Then
            
            TipoMeioPagtoDe.ListIndex = iIndice
            Exit For
        
        End If
    
    Next
    
    'preenche o frame De
    Select Case objTipoMovtoCaixa.iTMPLoja
    
        Case TIPOMEIOPAGTOLOJA_CARTAO_CREDITO, TIPOMEIOPAGTOLOJA_CARTAO_DEBITO
            
            If objMovimentoCaixa.iAdmMeioPagto <> 0 Then
            
                'preenche a combo com o código do cartão
                AdmCrtDe.Text = objMovimentoCaixa.iAdmMeioPagto
                Call AdmCrtDe_Validate(bSGECancelDummy)
            
                If objMovimentoCaixa.iParcelamento <> 0 Then
            
                    'prenche o parcelamento
                    ParcelamentoCrtDe.Text = objMovimentoCaixa.iParcelamento
                    Call ParcelamentoCrtDe_Validate(bSGECancelDummy)
            
                End If
            
            End If
            
            If objTipoMovtoCaixa.iTMPLoja = TIPOMEIOPAGTOLOJA_CARTAO_CREDITO Then
            
                'seleciona se é manual ou pos
                If objMovimentoCaixa.iTipoCartao = TIPO_MANUAL Then
                    OptionManualDe.Value = True
                Else
                    OptionPOSDe.Value = True
                End If
            
            Else
            
                OptionPOSDe.Value = True
                
            End If
            
            'preenche o valor
            ValorCrtDe.Text = Format(objMovimentoCaixa.dValor, "STANDARD")
            
        Case TIPOMEIOPAGTOLOJA_CHEQUE
        
            'preenche os dados para buscar o cheque
            objChequePre.iFilialEmpresa = giFilialEmpresa
            objChequePre.iFilialEmpresaLoja = giFilialEmpresa
            
            If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
            
                objChequePre.lSequencialLoja = objMovimentoCaixa.lNumRefInterna
                
            Else
            
                objChequePre.lSequencialBack = objMovimentoCaixa.lNumRefInterna
                
            End If
            
            'tenta pegar o cupom ou o carne associado ao cheque
            lErro = Trata_Cheque(objChequePre)
            If lErro <> SUCESSO Then gError 105288
                        
            'preenche a tela com o cheque
            lErro = Traz_Cheque_Tela_De(objChequePre)
            If lErro <> SUCESSO Then gError 113613
            
        Case TIPOMEIOPAGTOLOJA_DINHEIRO
            
            'preenche o valor
            ValorDinDe.Text = Format(objMovimentoCaixa.dValor, "STANDARD")
        
        Case TIPOMEIOPAGTOLOJA_OUTROS
        
            If objMovimentoCaixa.iAdmMeioPagto <> 0 Then
        
                'preenche o admmeiopagto
                AdmOutDe.Text = objMovimentoCaixa.iAdmMeioPagto
                Call AdmOutDe_Validate(bSGECancelDummy)
            
                If objMovimentoCaixa.iParcelamento <> 0 Then
            
                    'prenche o parcelamento
                    ParcelamentoOutDe.Text = objMovimentoCaixa.iParcelamento
                    Call ParcelamentoOutDe_Validate(bSGECancelDummy)
            
                End If
                
            End If
            
            'preenche o valor
            ValorOutDe.Text = Format(objMovimentoCaixa.dValor, "STANDARD")
            
        Case TIPOMEIOPAGTOLOJA_VALE_TICKET
        
            If objMovimentoCaixa.iAdmMeioPagto <> 0 Then
        
                'preenche o admmeiopagto
                AdmTktDe.Text = objMovimentoCaixa.iAdmMeioPagto
                Call AdmTktDe_Validate(bSGECancelDummy)
            
            
                If objMovimentoCaixa.iParcelamento <> 0 Then
            
                    'prenche o parcelamento
                    ParcelamentoTktDe.Text = objMovimentoCaixa.iParcelamento
                    Call ParcelamentoTktDe_Validate(bSGECancelDummy)
            
                End If
                
            End If
            
            'preenche o valor
            ValorTktDe.Text = Format(objMovimentoCaixa.dValor, "STANDARD")
    
    End Select
    
    Preenche_Frame_De = SUCESSO
    
    Exit Function
    
Erro_Preenche_Frame_De:
    
    Preenche_Frame_De = gErr
    
    Select Case gErr
    
        Case 105288, 113600, 113613

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175415)
    
    End Select
    
    Exit Function

End Function
    
Private Function Preenche_Frame_Para(objMovimentoCaixa As ClassMovimentoCaixa) As Long

Dim objTipoMovtoCaixa As New ClassTipoMovtoCaixa
Dim iIndice As Integer
Dim lErro As Long
Dim objChequePre As New ClassChequePre
Dim objCupomFiscal As New ClassCupomFiscal

On Error GoTo Erro_Preenche_Frame_Para
    
    'preenche o codigo de um movto caixa
    objTipoMovtoCaixa.iCodigo = objMovimentoCaixa.iTipo
    
    'veriica qual o tmploja desse movtocaixa
    lErro = CF("TiposMovtoCaixa_Le_Codigo", objTipoMovtoCaixa)
    If lErro <> SUCESSO Then gError 113601
    
    'seleciona na combo "PARA"
    For iIndice = 0 To TipoMeioPagtoPara.ListCount - 1
        
        If TipoMeioPagtoPara.ItemData(iIndice) = objTipoMovtoCaixa.iTMPLoja Then
            
            TipoMeioPagtoPara.ListIndex = iIndice
            Exit For
        
        End If
    
    Next
    
    'preenche o frame Para
    Select Case objTipoMovtoCaixa.iTMPLoja
    
        Case TIPOMEIOPAGTOLOJA_CARTAO_CREDITO
            
            'preenche a combo com o código do cartão
            If objMovimentoCaixa.iAdmMeioPagto <> 0 Then
                
                AdmCrtPara.Text = objMovimentoCaixa.iAdmMeioPagto
                Call AdmCrtPara_Validate(bSGECancelDummy)
            
                'prenche o parcelamento
                If objMovimentoCaixa.iParcelamento <> 0 Then
                    
                    ParcelamentoCrtPara.Text = objMovimentoCaixa.iParcelamento
                    Call ParcelamentoCrtPara_Validate(bSGECancelDummy)
                
                End If
            
            End If
            
            'seleciona se é manual ou pos
            If objMovimentoCaixa.iTipoCartao = TIPO_MANUAL Then
                OptionManualPara.Value = True
            Else
                OptionPOSPara.Value = True
            End If
            
        Case TIPOMEIOPAGTOLOJA_CARTAO_DEBITO
            
            If objMovimentoCaixa.iAdmMeioPagto <> 0 Then
            
                'preenche a combo com o código do cartão
                AdmCrtPara.Text = objMovimentoCaixa.iAdmMeioPagto
                Call AdmCrtPara_Validate(bSGECancelDummy)
            
                        
                'prenche o parcelamento
                If objMovimentoCaixa.iParcelamento <> 0 Then
            
                    'prenche o parcelamento
                    ParcelamentoCrtPara.Text = objMovimentoCaixa.iParcelamento
                    Call ParcelamentoCrtPara_Validate(bSGECancelDummy)
        
                End If
        
            End If
        
        Case TIPOMEIOPAGTOLOJA_CHEQUE
        
            'preenche os dados para buscar o cheque
            objChequePre.iFilialEmpresaLoja = giFilialEmpresa
            objChequePre.iFilialEmpresa = giFilialEmpresa
            
            If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
                objChequePre.lSequencialLoja = objMovimentoCaixa.lNumRefInterna
            ElseIf giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL_BACKOFFICE Then
                objChequePre.lSequencialBack = objMovimentoCaixa.lNumRefInterna
            Else
                gError 105251
            End If
            
            'lê o cheque
            lErro = CF("Cheque_Le_Excluido", objChequePre)
            If lErro <> SUCESSO And lErro <> 105310 And lErro <> 105313 Then gError 105278
            
            'se não encontrar-> erro
            If lErro = 105310 Or lErro = 105313 Then gError 105279
            
            'se o movimento de caixa estiver associado a um cupom
            If objMovimentoCaixa.lCupomFiscal <> 0 Then
            
                objCupomFiscal.iFilialEmpresa = giFilialEmpresa
                objCupomFiscal.lNumIntDoc = objMovimentoCaixa.lCupomFiscal
            
                'tenta ler o cupom fiscal para trazer o ECF e o COO
                lErro = CF("CupomFiscal_Le_NumIntDoc", objCupomFiscal)
                If lErro <> SUCESSO And lErro <> 105268 Then gError 105280
                
                'se nao encontrou o cupom ==> erro
                If lErro = 105268 Then gError 105281
                
                objChequePre.lCupomFiscal = objCupomFiscal.lNumero
                objChequePre.iECF = objCupomFiscal.iECF
            
            Else
            
                'le o cheque se estiver vinculado a carne
                lErro = CF("ChequePre_Le_Carne", objChequePre)
                If lErro <> SUCESSO And lErro <> 105256 Then gError 105282
            
            End If
            
            'preenche a tela com o cheque
            lErro = Traz_Cheque_Tela_Para(objChequePre)
            If lErro <> SUCESSO Then gError 113620
            
        Case TIPOMEIOPAGTOLOJA_DINHEIRO
            
            'não faz nada
        
        Case TIPOMEIOPAGTOLOJA_OUTROS
        
            If objMovimentoCaixa.iAdmMeioPagto <> 0 Then
        
                'preenche o admmeiopagto
                AdmOutPara.Text = objMovimentoCaixa.iAdmMeioPagto
                Call AdmOutPara_Validate(bSGECancelDummy)
            
            End If
            
        Case TIPOMEIOPAGTOLOJA_VALE_TICKET
        
            If objMovimentoCaixa.iAdmMeioPagto <> 0 Then
        
                'preenche o admmeiopagto
                AdmTktPara.Text = objMovimentoCaixa.iAdmMeioPagto
                Call AdmTktPara_Validate(bSGECancelDummy)
            
            End If
            
    End Select
    
    Preenche_Frame_Para = SUCESSO
    
    Exit Function
    
Erro_Preenche_Frame_Para:
    
    Preenche_Frame_Para = gErr
    
    Select Case gErr
    
        Case 105251
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCALOPERACAO_INVALIDO", gErr, giLocalOperacao)
    
        Case 105257, 105278, 105280, 105282, 113601, 113618, 113620
        
        Case 105279
            Call Rotina_Erro(vbOKOnly, "ERRO_CHEQUEPRE_NAOENCONTRADO", gErr, objChequePre.iFilialEmpresa, objChequePre.lSequencial)
        
        Case 105281
            Call Rotina_Erro(vbOKOnly, "ERRO_CUPOM_FISCAL_NAO_CADASTRADO1", gErr, objCupomFiscal.lNumIntDoc)
        
        Case 105283
            Call Rotina_Erro(vbOKOnly, "ERRO_CHEQUE_NAO_CUPOM_CARNE", gErr, objCupomFiscal.lNumIntDoc)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175416)
        
    End Select
    
    Exit Function

End Function

Private Function Traz_Cheque_Tela_Para(objChequePre As ClassChequePre) As Long

Dim bCancel As Boolean

On Error GoTo Erro_Traz_Cheque_Tela_Para

    'se o cheque for especificado
    If objChequePre.iNaoEspecificado = CHEQUE_ESPECIFICADO Then
        
        'preenche o banco
        BancoChqPara.Text = objChequePre.iBanco
        
        'preenche a agencia
        AgenciaChqPara.Text = objChequePre.sAgencia
        
        'preenche a conta
        ContaChqPara.Text = objChequePre.sContaCorrente
        
        'preenche o número
        NumeroChqPara.Text = objChequePre.lNumero
        
        'preenche o campo com o cpf/cgc formatado
        ClienteChqPara.PromptInclude = False
        ClienteChqPara.Text = objChequePre.sCPFCGC
        ClienteChqPara.PromptInclude = True
        
        Call ClienteChqPara_Validate(bCancel)
        
    End If
    
    'preenche o ecf
    If objChequePre.iECF <> 0 Then ECFChqPara.Text = objChequePre.iECF
    
    'preenche a data de depósito
    DataBomParaChqPara.PromptInclude = False
    DataBomParaChqPara.Text = Format(objChequePre.dtDataDeposito, "dd/mm/yy")
    DataBomParaChqPara.PromptInclude = True
    
    'preenche o carnê
    CarneChqPara.PromptInclude = False
    CarneChqPara.Text = objChequePre.sCarne
    CarneChqPara.PromptInclude = True
    
    'preenche o cupom fiscal
    If objChequePre.lCupomFiscal <> 0 Then CupomFiscalChqPara.Text = objChequePre.lCupomFiscal
    
    Traz_Cheque_Tela_Para = SUCESSO
    
    Exit Function
    
Erro_Traz_Cheque_Tela_Para:
    
    Traz_Cheque_Tela_Para = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175417)
    
    End Select
    
    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objTransfCaixa As New ClassTransfCaixa
Dim vbMsgResp As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click
    
    'mouse-ampulheta
    GL_objMDIForm.MousePointer = vbHourglass
    
    'se o código não estiver preenchido-> erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 113715
    
    'peço confirmação ao usuário
    vbMsgResp = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_TRANSFCENTRAL", objTransfCaixa.objMovCaixaDe.lTransferencia)
    
    'se a resposta for sim
    If vbMsgResp = vbYes Then
    
        objTransfCaixa.lCodigo = StrParaLong(Codigo.Text)
        objTransfCaixa.iFilialEmpresa = giFilialEmpresa
    
        'chama a função de exclusão
        lErro = CF("TransfCentral_Exclui", objTransfCaixa)
        If lErro <> SUCESSO Then gError 113740
    
        'limpa a tela
        Call Limpa_Tela_TransfCentral
        
        iAlterado = 0
    
    End If
    
    'mouse-padrão
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub
    
Erro_BotaoExcluir_Click:
    
    Select Case gErr
    
        Case 113715
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
            
        Case 113741, 113740
        
        Case 113742, 113743
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSFCAIXA_NAO_CADASTRADO", gErr, objTransfCaixa.objMovCaixaDe.lTransferencia)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175418)
    
    End Select
    
    'mouse-padrão
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 113624
    
    Call Limpa_Tela_TransfCentral
    
    Exit Sub
    
Erro_BotaoGravar_Click:
    
    Select Case gErr
    
        Case 113624
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175419)
    
    End Select
    
    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objTransfCaixa As New ClassTransfCaixa
Dim objChequeDe As New ClassChequePre
Dim objChequePara As New ClassChequePre

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'se o código não estiver preenchido-> erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 113625
    
    'se o tipomeiopagtode nao estiver preenchido, erro
    If Len(Trim(TipoMeioPagtoDe.Text)) = 0 Then gError 126004

    'se o tipomeiopagtopara nao estiver preenchido, erro
    If Len(Trim(TipoMeioPagtoPara.Text)) = 0 Then gError 126005

    'verifica a integridade do frame "De"
    lErro = Trata_Frames_De()
    If lErro <> SUCESSO Then gError 113626
    
    'verifica a integridade do frame "Para"
    lErro = Trata_Frames_Para()
    If lErro <> SUCESSO Then gError 113627
    
    'move os dados para a memória
    lErro = Move_Tela_Memoria(objTransfCaixa, objChequeDe, objChequePara)
    If lErro <> SUCESSO Then gError 113628
    
    'pergunta se deseja alterar
    lErro = Trata_Alteracao(objTransfCaixa, objTransfCaixa.iFilialEmpresa, objTransfCaixa.lCodigo)
    If lErro <> SUCESSO Then gError 113629
    
    'chama a função de gravação
    lErro = CF("TransfCentral_Grava", objTransfCaixa, objChequeDe, objChequePara)
    If lErro <> SUCESSO Then gError 113630
    
    'mouse-padrão
    GL_objMDIForm.MousePointer = vbDefault
    
    iAlterado = 0
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:
    
    Gravar_Registro = gErr
    
    Select Case gErr
    
        Case 113625
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
            
        Case 113626 To 113630
        
        Case 126004
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTODE_NAO_PREENCHIDO", gErr)

        Case 126005
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTOPARA_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175420)
    
    End Select
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

End Function

Private Function Move_Tela_Memoria(objTransfCaixa As ClassTransfCaixa, objChequeDe As ClassChequePre, objChequePara As ClassChequePre) As Long

Dim objCupomFiscal As New ClassCupomFiscal
Dim objAdmMeioPagto As New ClassAdmMeioPagto
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objTransfCaixa.iFilialEmpresa = giFilialEmpresa
    objTransfCaixa.lCodigo = StrParaLong(Codigo.Text)

    'dados fixos independentes de tipomeiopagto
    objTransfCaixa.objMovCaixaDe.iFilialEmpresa = giFilialEmpresa
    objTransfCaixa.objMovCaixaDe.iCaixa = CODIGO_CAIXA_CENTRAL
    objTransfCaixa.objMovCaixaDe.dtDataMovimento = gdtDataAtual
    objTransfCaixa.objMovCaixaDe.dHora = CDbl(Time())
    objTransfCaixa.objMovCaixaDe.lTransferencia = StrParaLong(Codigo.Text)
    
    If TipoMeioPagtoDe.ListIndex <> -1 Then
        
        Select Case TipoMeioPagtoDe.ItemData(TipoMeioPagtoDe.ListIndex)
            
            Case TIPOMEIOPAGTOLOJA_CHEQUE
            
                objTransfCaixa.objMovCaixaDe.iTipo = MOVIMENTOCAIXA_SAIDA_TRANSF_CHEQUE
                objTransfCaixa.objMovCaixaDe.iAdmMeioPagto = MEIO_PAGAMENTO_CHEQUE
                objTransfCaixa.objMovCaixaDe.iParcelamento = PARCELAMENTO_AVISTA
                objTransfCaixa.objMovCaixaDe.dValor = StrParaDbl(ValorChqDe.Caption)
                '**** numero do cupom. Nao é o numintdoc que é gravado
                objTransfCaixa.objMovCaixaDe.lCupomFiscal = StrParaLong(CupomFiscalChqDe.Caption)
                objTransfCaixa.objMovCaixaDe.lNumero = StrParaLong(NumeroChqDe.Caption)
                objChequeDe.lCupomFiscal = StrParaLong(CupomFiscalChqDe.Caption)
                objChequeDe.iECF = StrParaInt(ECFChqDe.Caption)
                objChequeDe.iFilialEmpresa = giFilialEmpresa
                objChequeDe.iFilialEmpresaLoja = giFilialEmpresa
                objTransfCaixa.objMovCaixaDe.lNumRefInterna = StrParaLong(SeqChqDe.Text)
                
                If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
                    objChequeDe.iFilialEmpresaLoja = objTransfCaixa.objMovCaixaDe.iFilialEmpresa
                    objChequeDe.lSequencialLoja = StrParaLong(SeqChqDe.Text)
                Else
                    objChequeDe.lSequencialBack = StrParaLong(SeqChqDe.Text)
                End If
            
            Case TIPOMEIOPAGTOLOJA_DINHEIRO
                
                objTransfCaixa.objMovCaixaDe.iTipo = MOVIMENTOCAIXA_SAIDA_TRANSF_DINHEIRO
                objTransfCaixa.objMovCaixaDe.iAdmMeioPagto = MEIO_PAGAMENTO_DINHEIRO
                objTransfCaixa.objMovCaixaDe.iParcelamento = PARCELAMENTO_AVISTA
                objTransfCaixa.objMovCaixaDe.dValor = StrParaDbl(ValorDinDe.Text)
            
            Case TIPOMEIOPAGTOLOJA_CARTAO_CREDITO
                
                objTransfCaixa.objMovCaixaDe.iTipo = MOVIMENTOCAIXA_SAIDA_TRANSF_CARTAO_CREDITO
                objTransfCaixa.objMovCaixaDe.dValor = StrParaDbl(ValorCrtDe.Text)
                
                'se for cartão especificado
                If Len(Trim(AdmCrtDe.Text)) <> 0 Then
                    
                    objTransfCaixa.objMovCaixaDe.iAdmMeioPagto = AdmCrtDe.ItemData(AdmCrtDe.ListIndex)
                    objTransfCaixa.objMovCaixaDe.iParcelamento = ParcelamentoCrtDe.ItemData(ParcelamentoCrtDe.ListIndex)
                    
                End If
                
                If OptionManualDe.Value = True Then
                    objTransfCaixa.objMovCaixaDe.iTipoCartao = TIPO_MANUAL
                Else
                    objTransfCaixa.objMovCaixaDe.iTipoCartao = TIPO_POS
                End If
                
            Case TIPOMEIOPAGTOLOJA_CARTAO_DEBITO
                
                objTransfCaixa.objMovCaixaDe.iTipo = MOVIMENTOCAIXA_SAIDA_TRANSF_CARTAO_DEBITO
                objTransfCaixa.objMovCaixaDe.dValor = StrParaDbl(ValorCrtDe.Text)
                
                'se for cartão especificado
                If Len(Trim(AdmCrtDe.Text)) <> 0 Then
                    
                    objTransfCaixa.objMovCaixaDe.iAdmMeioPagto = AdmCrtDe.ItemData(AdmCrtDe.ListIndex)
                    objTransfCaixa.objMovCaixaDe.iParcelamento = ParcelamentoCrtDe.ItemData(ParcelamentoCrtDe.ListIndex)
                    
                End If
            
                objTransfCaixa.objMovCaixaDe.iTipoCartao = TIPO_POS
            
            Case TIPOMEIOPAGTOLOJA_OUTROS
                
                objTransfCaixa.objMovCaixaDe.iTipo = MOVIMENTOCAIXA_SAIDA_TRANSF_OUTROS
                objTransfCaixa.objMovCaixaDe.dValor = StrParaDbl(ValorOutDe.Text)
                
                'se a adm foi especificada
                If Len(Trim(AdmOutDe.Text)) <> 0 Then
                    objTransfCaixa.objMovCaixaDe.iAdmMeioPagto = AdmOutDe.ItemData(AdmOutDe.ListIndex)
                    objTransfCaixa.objMovCaixaDe.iParcelamento = ParcelamentoOutDe.ItemData(ParcelamentoOutDe.ListIndex)
                End If
            
            Case TIPOMEIOPAGTOLOJA_VALE_TICKET
                
                objTransfCaixa.objMovCaixaDe.iTipo = MOVIMENTOCAIXA_SAIDA_TRANSF_VALETICKET
                objTransfCaixa.objMovCaixaDe.dValor = StrParaDbl(ValorTktDe.Text)
                
                'se a adm foi especificada
                If Len(Trim(AdmTktDe.Text)) <> 0 Then
                    objTransfCaixa.objMovCaixaDe.iAdmMeioPagto = AdmTktDe.ItemData(AdmTktDe.ListIndex)
                    objTransfCaixa.objMovCaixaDe.iParcelamento = ParcelamentoTktDe.ItemData(ParcelamentoTktDe.ListIndex)
                End If
                
        End Select
        
    End If

    'dados que podem(e devem) ser copiados do Movto De
    objTransfCaixa.objMovCaixaPara.iFilialEmpresa = objTransfCaixa.objMovCaixaDe.iFilialEmpresa
    objTransfCaixa.objMovCaixaPara.iCaixa = objTransfCaixa.objMovCaixaDe.iCaixa
    objTransfCaixa.objMovCaixaPara.dtDataMovimento = objTransfCaixa.objMovCaixaDe.dtDataMovimento
    objTransfCaixa.objMovCaixaPara.dValor = objTransfCaixa.objMovCaixaDe.dValor
    objTransfCaixa.objMovCaixaPara.dHora = objTransfCaixa.objMovCaixaDe.dHora
    objTransfCaixa.objMovCaixaPara.lTransferencia = objTransfCaixa.objMovCaixaDe.lTransferencia
    
    If TipoMeioPagtoPara.ListIndex <> -1 Then
        
        Select Case TipoMeioPagtoPara.ItemData(TipoMeioPagtoPara.ListIndex)
            
            Case TIPOMEIOPAGTOLOJA_CHEQUE
                
                objChequePara.dtDataDeposito = StrParaDate(DataBomParaChqPara.Text)
                objChequePara.dValor = objTransfCaixa.objMovCaixaDe.dValor
                objChequePara.iBanco = StrParaInt(BancoChqPara.Text)
                objChequePara.iECF = StrParaInt(ECFChqPara.Text)
                objChequePara.iFilialEmpresa = giFilialEmpresa
                objChequePara.iFilialEmpresaLoja = giFilialEmpresa
                
                If objChequePara.iBanco <> 0 Then
                    objChequePara.iNaoEspecificado = CHEQUE_ESPECIFICADO
                Else
                    objChequePara.iNaoEspecificado = CHEQUE_NAO_ESPECIFICADO
                End If
                
                objChequePara.iStatus = STATUS_ATIVO
                '*** numero do cupom. nao o numintdoc que é gravado
                objChequePara.lCupomFiscal = StrParaLong(CupomFiscalChqPara.Text)
                objChequePara.lNumero = StrParaLong(NumeroChqPara.Text)
                objChequePara.sAgencia = Trim(AgenciaChqPara.Text)
                objChequePara.sCarne = Trim(CarneChqPara.Text)
                objChequePara.sContaCorrente = Trim(ContaChqPara.Text)
                objChequePara.sCPFCGC = Trim(ClienteChqPara.Text)
                
                objTransfCaixa.objMovCaixaPara.iTipo = MOVIMENTOCAIXA_ENTRADA_TRANSF_CHEQUE
                objTransfCaixa.objMovCaixaPara.iAdmMeioPagto = MEIO_PAGAMENTO_CHEQUE
                objTransfCaixa.objMovCaixaPara.iParcelamento = PARCELAMENTO_AVISTA
                objTransfCaixa.objMovCaixaPara.lNumero = objChequePara.lNumero
                
                
            Case TIPOMEIOPAGTOLOJA_DINHEIRO
                
                objTransfCaixa.objMovCaixaPara.iTipo = MOVIMENTOCAIXA_ENTRADA_TRANSF_DINHEIRO
                objTransfCaixa.objMovCaixaPara.iAdmMeioPagto = MEIO_PAGAMENTO_DINHEIRO
                objTransfCaixa.objMovCaixaPara.iParcelamento = PARCELAMENTO_AVISTA
            
            Case TIPOMEIOPAGTOLOJA_CARTAO_CREDITO
                
                objTransfCaixa.objMovCaixaPara.iTipo = MOVIMENTOCAIXA_ENTRADA_TRANSF_CARTAO_CREDITO
                
                'se for cartão especificado
                If Len(Trim(AdmCrtPara.Text)) <> 0 Then
                    
                    objTransfCaixa.objMovCaixaPara.iAdmMeioPagto = AdmCrtPara.ItemData(AdmCrtPara.ListIndex)
                    objTransfCaixa.objMovCaixaPara.iParcelamento = ParcelamentoCrtPara.ItemData(ParcelamentoCrtPara.ListIndex)
                    
                End If
                
                'verifica se é pos ou manual
                If OptionPOSPara.Value = True Then
                    objTransfCaixa.objMovCaixaPara.iTipoCartao = TIPO_POS
                Else
                    objTransfCaixa.objMovCaixaPara.iTipoCartao = TIPO_MANUAL
                End If
            
            Case TIPOMEIOPAGTOLOJA_CARTAO_DEBITO
                
                objTransfCaixa.objMovCaixaPara.iTipo = MOVIMENTOCAIXA_ENTRADA_TRANSF_CARTAO_DEBITO
                
                'se for cartão especificado
                If Len(Trim(AdmCrtPara.Text)) <> 0 Then
                    
                    objTransfCaixa.objMovCaixaPara.iAdmMeioPagto = AdmCrtPara.ItemData(AdmCrtPara.ListIndex)
                    objTransfCaixa.objMovCaixaPara.iParcelamento = ParcelamentoCrtPara.ItemData(ParcelamentoCrtPara.ListIndex)
                    
                End If
            
            Case TIPOMEIOPAGTOLOJA_OUTROS
                
                objTransfCaixa.objMovCaixaPara.iTipo = MOVIMENTOCAIXA_ENTRADA_TRANSF_OUTROS
                
                objTransfCaixa.objMovCaixaPara.iAdmMeioPagto = AdmOutPara.ItemData(AdmOutPara.ListIndex)
            
                objAdmMeioPagto.iCodigo = objTransfCaixa.objMovCaixaPara.iAdmMeioPagto
                objAdmMeioPagto.iFilialEmpresa = objTransfCaixa.objMovCaixaPara.iFilialEmpresa
        
                'lê os parcelamentos da adm
                lErro = CF("AdmMeioPagtoCondPagto_Le", objAdmMeioPagto)
                If lErro <> SUCESSO And lErro <> 104086 Then gError 126066
                
                For Each objAdmMeioPagtoCondPagto In objAdmMeioPagto.colCondPagtoLoja
                    If objAdmMeioPagtoCondPagto.iAtivo = ADMMEIOPAGTOCONDPAGTO_ATIVO Then
                        objTransfCaixa.objMovCaixaPara.iParcelamento = objAdmMeioPagtoCondPagto.iParcelamento
                        Exit For
                    End If
                Next
        
                'se nao encontrou parcelamento ativo para a administradora em questao ==> erro
                If objTransfCaixa.objMovCaixaPara.iParcelamento = 0 Then gError 126067
            
            Case TIPOMEIOPAGTOLOJA_VALE_TICKET
                
                objTransfCaixa.objMovCaixaPara.iTipo = MOVIMENTOCAIXA_ENTRADA_TRANSF_VALETICKET
                
                objTransfCaixa.objMovCaixaPara.iAdmMeioPagto = AdmTktPara.ItemData(AdmTktPara.ListIndex)
                objAdmMeioPagto.iCodigo = objTransfCaixa.objMovCaixaPara.iAdmMeioPagto
                objAdmMeioPagto.iFilialEmpresa = objTransfCaixa.objMovCaixaPara.iFilialEmpresa
        
                'lê os parcelamentos da adm
                lErro = CF("AdmMeioPagtoCondPagto_Le", objAdmMeioPagto)
                If lErro <> SUCESSO And lErro <> 104086 Then gError 126065
                
                For Each objAdmMeioPagtoCondPagto In objAdmMeioPagto.colCondPagtoLoja
                    If objAdmMeioPagtoCondPagto.iAtivo = ADMMEIOPAGTOCONDPAGTO_ATIVO Then
                        objTransfCaixa.objMovCaixaPara.iParcelamento = objAdmMeioPagtoCondPagto.iParcelamento
                        Exit For
                    End If
                Next
        
                'se nao encontrou parcelamento ativo para a administradora em questao ==> erro
                If objTransfCaixa.objMovCaixaPara.iParcelamento = 0 Then gError 126064
        
        End Select
        
    End If
    
    Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Tela_Memoria:
    
    Move_Tela_Memoria = gErr
    
    Select Case gErr
    
        Case 126064, 126067
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_SEM_PARCELAMENTO_ATIVO", gErr, objAdmMeioPagto.iCodigo)
    
        Case 126065, 126066
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175421)
    
    End Select
    
    Exit Function

End Function

Private Function Trata_Frames_De() As Long

Dim lErro As Long
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto
Dim objTMPLojaFilial As New ClassTMPLojaFilial
Dim dValor As Double

On Error GoTo Erro_Trata_Frames_De

    If TipoMeioPagtoDe.ListIndex = -1 Then gError 113639
    
    Select Case Codigo_Extrai(TipoMeioPagtoDe.Text)
    
        Case TIPOMEIOPAGTOLOJA_DINHEIRO
            
            'se o valor não estiver preenchido-> erro
            If Len(Trim(ValorDinDe.Text)) = 0 Then gError 113635
            
        Case TIPOMEIOPAGTOLOJA_CHEQUE
        
            'se o seqüencial não estiver preenchido-> erro
            If Len(Trim(SeqChqDe.Text)) = 0 Then gError 113638
            
        Case TIPOMEIOPAGTOLOJA_CARTAO_CREDITO, TIPOMEIOPAGTOLOJA_CARTAO_DEBITO
        
            'se a admmeiopagto estiver preenchida, mas o parcelamento não-> erro
            If Len(Trim(AdmCrtDe.Text)) <> 0 And _
               Len(Trim(ParcelamentoCrtDe.Text)) = 0 Then gError 113631
            
            'se o valor não estiver preenchido-> erro
            If Len(Trim(ValorCrtDe.Text)) = 0 Then gError 113632
            
            
        Case TIPOMEIOPAGTOLOJA_VALE_TICKET
        
            'se a admmeiopagto estiver preenchida, mas o parcelamento não-> erro
            If Len(Trim(AdmTktDe.Text)) <> 0 And _
               Len(Trim(ParcelamentoTktDe.Text)) = 0 Then gError 126059
        
            'se o valor não estiver preenchido-> erro
            If Len(Trim(ValorTktDe.Text)) = 0 Then gError 113636
            
        Case TIPOMEIOPAGTOLOJA_OUTROS
        
            'se a admmeiopagto estiver preenchida, mas o parcelamento não-> erro
            If Len(Trim(AdmOutDe.Text)) <> 0 And _
               Len(Trim(ParcelamentoOutDe.Text)) = 0 Then gError 126068
        
            'se o valor não estiver Preenchido-> erro
            If Len(Trim(ValorOutDe.Text)) = 0 Then gError 113637
            
    End Select

    Trata_Frames_De = SUCESSO
    
    Exit Function
    
Erro_Trata_Frames_De:
    
    Trata_Frames_De = gErr
    
    Select Case gErr
    
        Case 113631, 113633, 126059, 126068
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAMENTO_NAO_SELECIONADO1", gErr)
            
        Case 113632, 113634 To 113637
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_PREENCHIDO1", gErr)
            
        Case 113638
            Call Rotina_Erro(vbOKOnly, "ERRO_CHEQUEDE_SEQUENCIAL_NAO_INFORMADO", gErr)
            
        Case 113639
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175422)
    
    End Select
    
    Exit Function

End Function

Private Function Trata_Frames_Para() As Long

On Error GoTo Erro_Trata_Frames_Para

    If TipoMeioPagtoPara.ListIndex = -1 Then gError 113640
    
    Select Case TipoMeioPagtoPara.ItemData(TipoMeioPagtoPara.ListIndex)
    
        Case TIPOMEIOPAGTOLOJA_CARTAO_CREDITO
            
            'se a admmeiopagto estiver preenchida, mas o parcelamento não-> erro
            If Len(Trim(AdmCrtPara.Text)) <> 0 And _
               Len(Trim(ParcelamentoCrtPara.Text)) = 0 Then gError 113641
            
        Case TIPOMEIOPAGTOLOJA_CARTAO_DEBITO
        
            'se a admmeiopagto estiver preenchida, mas o parcelamento não-> erro
            If Len(Trim(AdmCrtPara.Text)) <> 0 And _
               Len(Trim(ParcelamentoCrtPara.Text)) = 0 Then gError 113642
            
        Case TIPOMEIOPAGTOLOJA_DINHEIRO, TIPOMEIOPAGTOLOJA_VALE_TICKET, TIPOMEIOPAGTOLOJA_OUTROS
            'não faz nada
            
        Case TIPOMEIOPAGTOLOJA_CHEQUE
            
            'se dentre Banco, agencia, número, conta e cheque, um deles ou mais estiverem preenchidos
            '-> TODOS os restantes também devê-lo-ão estar
            If (Len(Trim(BancoChqPara.Text)) <> 0 Or _
                Len(Trim(AgenciaChqPara.Text)) <> 0 Or _
                Len(Trim(NumeroChqPara.Text)) <> 0 Or _
                Len(Trim(ContaChqPara.Text)) <> 0 Or _
                Len(Trim(ClienteChqPara.Text)) <> 0) And _
               (Len(Trim(BancoChqPara.Text)) = 0 Or _
                Len(Trim(AgenciaChqPara.Text)) = 0 Or _
                Len(Trim(NumeroChqPara.Text)) = 0 Or _
                Len(Trim(ContaChqPara.Text)) = 0 Or _
                Len(Trim(ClienteChqPara.Text)) = 0) Then gError 113643
                
            'se a data de depósito não estiver preenchida-> erro
            If Len(Trim(DataBomParaChqPara.ClipText)) = 0 Then gError 113644
            
            'se o tipomeiopagtoDe não for cheque
            If TipoMeioPagtoDe.ItemData(TipoMeioPagtoDe.ListIndex) <> TIPOMEIOPAGTOLOJA_CHEQUE Then
                
'                'ou o carnê ou o Cupom fiscal devem estar preenchidos
'                If Len(Trim(CarneChqPara.Text)) = 0 And _
'                   Len(Trim(CupomFiscalChqPara.Text)) = 0 Then gError 113645
                   
                'o carne e o cupom NÃO podem estar preenchidos simultaneamente
                If Len(Trim(CarneChqPara.Text)) <> 0 And _
                   Len(Trim(CupomFiscalChqPara.Text)) <> 0 Then gError 113646
                   
                If (Len(Trim(CupomFiscalChqPara.Text)) <> 0 And Len(Trim(ECFChqPara.Text)) = 0) Or (Len(Trim(CupomFiscalChqPara.Text)) = 0 And Len(Trim(ECFChqPara.Text)) <> 0) Then gError 113759
                
            
            End If
    
    End Select

    Trata_Frames_Para = SUCESSO
    
    Exit Function
    
Erro_Trata_Frames_Para:
    
    Trata_Frames_Para = gErr
    
    Select Case gErr
    
        Case 113641, 113642
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAMENTO_NAO_SELECIONADO1", gErr)
            
        Case 113645
            Call Rotina_Erro(vbOKOnly, "ERRO_CARNE_E_CUPOM_NAO_PREENCHIDOS", gErr)
        
        Case 113646
            Call Rotina_Erro(vbOKOnly, "ERRO_CARNE_E_CUPOM_PREENCHIDOS", gErr)
            
        Case 113643
            Call Rotina_Erro(vbOKOnly, "ERRO_GRUPOCHQPARA_NAO_PREENCHIDO", gErr)
            
        Case 113644
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_CREDITO_NAO_PREENCHIDA", gErr)
            
        Case 113759
            Call Rotina_Erro(vbOKOnly, "ERRO_GRUPO_CUPOMFISCAL_E_ECF_INCOMPLETO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175423)
    
    End Select
    
    Exit Function

End Function

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    lErro = TransfCentral_Codigo_Automatico(lCodigo)
    If lErro <> SUCESSO Then gError 113556
    
    Codigo.Text = lCodigo
    
    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175424)
    
    End Select

End Sub
    
Private Function TransfCentral_Codigo_Automatico(lCodigo As Long) As Long

Dim lErro As Long

On Error GoTo Erro_TransfCentral_Codigo_Automatico

    lErro = CF("Config_ObterAutomatico", "LojaConfig", "NUM_PROX_TRANSFCAIXACENTRAL", "MovimentosCaixa", "Transferencia", lCodigo)
    If lErro <> SUCESSO Then gError 113555
    
    TransfCentral_Codigo_Automatico = SUCESSO

    Exit Function
    
Erro_TransfCentral_Codigo_Automatico:
    
    TransfCentral_Codigo_Automatico = gErr
    
    Select Case gErr
    
        Case 113555
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175425)
    
    End Select

End Function

Private Sub ClienteChqPara_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ClienteChqPara_Validate

    'Se CGC/CPF não foi preenchido -- Exit Sub
    If Len(Trim(ClienteChqPara.Text)) = 0 Then Exit Sub
    
    Select Case Len(Trim(ClienteChqPara.Text))
    
        Case STRING_CPF 'CPF
        
            'Critica Cpf
            lErro = Cpf_Critica(ClienteChqPara.Text)
            If lErro <> SUCESSO Then gError 101975
            
            'Formata e coloca na Tela
            ClienteChqPara.Format = "000\.000\.000-00; ; ; "
            ClienteChqPara.Text = ClienteChqPara.Text
        
        Case STRING_CGC 'CGC
        
            'Critica CGC
            lErro = Cgc_Critica(ClienteChqPara.Text)
            If lErro <> SUCESSO Then gError 101976
            
            'Formata e Coloca na Tela
            ClienteChqPara.Format = "00\.000\.000\/0000-00; ; ; "
            ClienteChqPara.Text = ClienteChqPara.Text
        
        Case Else
        
            gError 101977
    
    End Select
    
    Exit Sub

Erro_ClienteChqPara_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 101975, 101976
        
        Case 101977
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC_CPF", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175426)
        
    End Select
    
    Exit Sub
    
End Sub

Private Sub DataBomParaChqPara_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataBomParaChqPara_Validate

    'se campo estiver preenchido
    If Len(Trim(DataBomParaChqPara.ClipText)) > 0 Then
    
        'critica o conteudo
        lErro = Data_Critica(DataBomParaChqPara.Text)
        If lErro <> SUCESSO Then gError 101980
    
    End If
    
    Exit Sub

Erro_DataBomParaChqPara_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 101980
        
        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error, 175427)
    
    End Select
    
    Exit Sub

End Sub

Private Sub LabelCupomFiscalChqPara_Click()

Dim objCupomFiscal As New ClassCupomFiscal
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCupomFiscalChqPara_Click

    'se o ECF estiver preenchido
    If Len(Trim(ECFChqPara.Text)) > 0 Then

        'move o ECF para o obj
        objCupomFiscal.iECF = StrParaInt(ECFChqPara.Text)

    End If

    'se o COO estiver preenchido
    If Len(Trim(CupomFiscalChqPara.Text)) > 0 Then

        'move o COO para o obj
        objCupomFiscal.lNumero = StrParaLong(CupomFiscalChqPara.Text)

    End If

    'Chama o Browser '
    Call Chama_Tela("CupomFiscalLista", colSelecao, objCupomFiscal, objEventoCupomFiscal)

    Exit Sub

Erro_LabelCupomFiscalChqPara_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175428)

    End Select

    Exit Sub

End Sub

Private Sub LabelECFChqPara_Click()

Dim objCupomFiscal As New ClassCupomFiscal
Dim colSelecao As New Collection

On Error GoTo Erro_LabelECFChqPara_Click

    'se o ECF estiver preenchido
    If Len(Trim(ECFChqPara.Text)) > 0 Then

        'move o ECF para o obj
        objCupomFiscal.iECF = StrParaInt(ECFChqPara.Text)

    End If

    'se o COO estiver preenchido
    If Len(Trim(CupomFiscalChqPara.Text)) > 0 Then

        'move o COO para o obj
        objCupomFiscal.lNumero = StrParaLong(CupomFiscalChqPara.Text)

    End If

    'Chama o Browser '
    Call Chama_Tela("CupomFiscalLista", colSelecao, objCupomFiscal, objEventoCupomFiscal)

    Exit Sub

Erro_LabelECFChqPara_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175429)

    End Select

    Exit Sub

End Sub

Private Sub objEventoTransf_evSelecao(obj1 As Object)

Dim objTransfCaixa As New ClassTransfCaixa
Dim objTransfBrowser As ClassTransfBrowser
Dim lErro As Long

On Error GoTo Erro_objEventoTransf_evSelecao

    Set objTransfBrowser = obj1
    
    objTransfCaixa.iFilialEmpresa = objTransfBrowser.iFilialEmpresa
    objTransfCaixa.lCodigo = objTransfBrowser.lTransferencia

    'Move os dados para a tela
    lErro = Traz_TransfCaixa_Tela(objTransfCaixa)
    If lErro <> SUCESSO And lErro <> 105277 Then gError 111512

    If lErro = 105277 Then gError 126007

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoTransf_evSelecao:

    Select Case gErr

        Case 111512

        Case 126007
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSFERENCIALOJA_NAOCADASTRADA", gErr, objTransfCaixa.iFilialEmpresa, objTransfCaixa.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175430)

    End Select

    Exit Sub

End Sub

Private Sub ParcelamentoCrtDe_Validate(Cancel As Boolean)

Dim iCodigo As Integer
Dim lErro As Long
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_ParcelamentoCrtDe_Validate

    If Len(Trim(ParcelamentoCrtDe.Text)) <> 0 Then
        
        lErro = Combo_Seleciona(ParcelamentoCrtDe, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 113603
        
        'se não encontrar pelo código
        If lErro = 6730 Then
        
            objAdmMeioPagtoCondPagto.iAdmMeioPagto = AdmCrtDe.ItemData(AdmCrtDe.ListIndex)
            objAdmMeioPagtoCondPagto.iParcelamento = iCodigo
            objAdmMeioPagtoCondPagto.iFilialEmpresa = giFilialEmpresa
            
            'buso no BD se existe um Admmeiopagtocondpagto com o código lido
            lErro = CF("AdmMeioPagtoCondPagto_Le_Parcelamento", objAdmMeioPagtoCondPagto)
            If lErro <> SUCESSO And lErro <> 107297 Then gError 113604
            
            'se não encontrar-> erro
            If lErro = 107297 Then gError 113605
            
            'se encontrar, preenche
            ParcelamentoCrtDe.Text = objAdmMeioPagtoCondPagto.iParcelamento & SEPARADOR & objAdmMeioPagtoCondPagto.sNomeParcelamento
        
        End If
        
        'se não encontrar pela string->erro
        If lErro = 6731 Then gError 113606
    
    End If

    Exit Sub
    
Erro_ParcelamentoCrtDe_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 113603, 113604
        
        Case 113605
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTOCONDPAGTO_INEXISTENTE", gErr, iCodigo)
            
        Case 113606
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTOCONDPAGTO_INEXISTENTE2", gErr, ParcelamentoCrtDe)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175431)
    
    End Select
    
    Exit Sub

End Sub

Private Sub ParcelamentoCrtPara_Validate(Cancel As Boolean)

Dim iCodigo As Integer
Dim lErro As Long
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_ParcelamentoCrtPara_Validate

    If Len(Trim(ParcelamentoCrtPara.Text)) <> 0 Then
        
        lErro = Combo_Seleciona(ParcelamentoCrtPara, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 113614
        
        'se não encontrar pelo código
        If lErro = 6730 Then
        
            objAdmMeioPagtoCondPagto.iAdmMeioPagto = AdmCrtPara.ItemData(AdmCrtPara.ListIndex)
            objAdmMeioPagtoCondPagto.iParcelamento = iCodigo
            objAdmMeioPagtoCondPagto.iFilialEmpresa = giFilialEmpresa
            
            'buso no BD se existe um Admmeiopagtocondpagto com o código lido
            lErro = CF("AdmMeioPagtoCondPagto_Le_Parcelamento", objAdmMeioPagtoCondPagto)
            If lErro <> SUCESSO And lErro <> 107297 Then gError 113615
            
            'se não encontrar-> erro
            If lErro = 107297 Then gError 113616
            
            'se encontrar, preenche
            ParcelamentoCrtPara.Text = objAdmMeioPagtoCondPagto.iParcelamento & SEPARADOR & objAdmMeioPagtoCondPagto.sNomeParcelamento
        
        End If
        
        'se não encontrar pela string->erro
        If lErro = 6731 Then gError 113617
    
    End If

    Exit Sub
    
Erro_ParcelamentoCrtPara_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 113614, 113615
        
        Case 113616
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTOCONDPAGTO_INEXISTENTE", gErr, iCodigo)
            
        Case 113617
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTOCONDPAGTO_INEXISTENTE2", gErr, ParcelamentoCrtPara)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175432)
    
    End Select
    
    Exit Sub

End Sub

Private Sub ParcelamentoTktDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ParcelamentoTktDe_Validate(Cancel As Boolean)

Dim iCodigo As Integer
Dim lErro As Long
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_ParcelamentoTktDe_Validate

    If Len(Trim(ParcelamentoTktDe.Text)) <> 0 Then
        
        lErro = Combo_Seleciona(ParcelamentoTktDe, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 126060
        
        'se não encontrar pelo código
        If lErro = 6730 Then
        
            objAdmMeioPagtoCondPagto.iAdmMeioPagto = AdmTktDe.ItemData(AdmTktDe.ListIndex)
            objAdmMeioPagtoCondPagto.iParcelamento = iCodigo
            objAdmMeioPagtoCondPagto.iFilialEmpresa = giFilialEmpresa
            
            'buso no BD se existe um Admmeiopagtocondpagto com o código lido
            lErro = CF("AdmMeioPagtoCondPagto_Le_Parcelamento", objAdmMeioPagtoCondPagto)
            If lErro <> SUCESSO And lErro <> 107297 Then gError 126061
            
            'se não encontrar-> erro
            If lErro = 107297 Then gError 126062
            
            'se encontrar, preenche
            ParcelamentoTktDe.Text = objAdmMeioPagtoCondPagto.iParcelamento & SEPARADOR & objAdmMeioPagtoCondPagto.sNomeParcelamento
        
        End If
        
        'se não encontrar pela string->erro
        If lErro = 6731 Then gError 126063
    
    End If

    Exit Sub
    
Erro_ParcelamentoTktDe_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 126060, 126061
        
        Case 126062
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTOCONDPAGTO_INEXISTENTE", gErr, iCodigo)
            
        Case 126063
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTOCONDPAGTO_INEXISTENTE2", gErr, ParcelamentoTktDe.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175433)
    
    End Select
    
    Exit Sub

End Sub

Private Sub ParcelamentoOutDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ParcelamentoOutDe_Validate(Cancel As Boolean)

Dim iCodigo As Integer
Dim lErro As Long
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_ParcelamentoOutDe_Validate

    If Len(Trim(ParcelamentoOutDe.Text)) <> 0 Then
        
        lErro = Combo_Seleciona(ParcelamentoOutDe, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 126069
        
        'se não encontrar pelo código
        If lErro = 6730 Then
        
            objAdmMeioPagtoCondPagto.iAdmMeioPagto = AdmOutDe.ItemData(AdmOutDe.ListIndex)
            objAdmMeioPagtoCondPagto.iParcelamento = iCodigo
            objAdmMeioPagtoCondPagto.iFilialEmpresa = giFilialEmpresa
            
            'buso no BD se existe um Admmeiopagtocondpagto com o código lido
            lErro = CF("AdmMeioPagtoCondPagto_Le_Parcelamento", objAdmMeioPagtoCondPagto)
            If lErro <> SUCESSO And lErro <> 107297 Then gError 126070
            
            'se não encontrar-> erro
            If lErro = 107297 Then gError 126071
            
            'se encontrar, preenche
            ParcelamentoOutDe.Text = objAdmMeioPagtoCondPagto.iParcelamento & SEPARADOR & objAdmMeioPagtoCondPagto.sNomeParcelamento
        
        End If
        
        'se não encontrar pela string->erro
        If lErro = 6731 Then gError 126072
    
    End If

    Exit Sub
    
Erro_ParcelamentoOutDe_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 126069, 126070
        
        Case 126071
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTOCONDPAGTO_INEXISTENTE", gErr, iCodigo)
            
        Case 126072
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTOCONDPAGTO_INEXISTENTE2", gErr, ParcelamentoOutDe.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175434)
    
    End Select
    
    Exit Sub

End Sub

Private Sub SeqChqDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCheque As New ClassChequePre

On Error GoTo Erro_SeqChqDe_Validate

    'se o código estiver preenchido
    If Len(Trim(SeqChqDe.Text)) <> 0 Then
    
        'permite somente a entrada de valor positivo
        lErro = Valor_Positivo_Critica(SeqChqDe.Text)
        If lErro <> SUCESSO Then gError 113623
        
        If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then

            objCheque.lSequencialLoja = StrParaLong(SeqChqDe.Text)
            objCheque.iFilialEmpresaLoja = giFilialEmpresa

        Else

            objCheque.lSequencialBack = StrParaLong(SeqChqDe.Text)

        End If
    
    
        'tenta pegar o cupom ou o carne associado ao cheque
        lErro = Trata_Cheque(objCheque)
        If lErro <> SUCESSO Then gError 105289
        
        'preenche a tela
        lErro = Traz_Cheque_Tela_De(objCheque)
        If lErro <> SUCESSO Then gError 113763
    
    Else
    
        Call Limpa_ChequeDe
    
    End If

    Exit Sub
    
Erro_SeqChqDe_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 105289, 113623, 113761, 113763
        
        Case 113762
            Call Rotina_Erro(vbOKOnly, "ERRO_CHEQUEPRE_NAO_CADASTRADO", gErr, objCheque.lSequencialLoja)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175435)
    
    End Select
    
    Exit Sub

End Sub

Private Sub TipoMeioPagtoDe_Click()

    iAlterado = REGISTRO_ALTERADO

    'chama o tratamento de frames
    If TipoMeioPagtoDe.ListIndex <> -1 Then Call TipoMeioPagto_TrataFramesDe

End Sub

Private Sub TipoMeioPagto_TrataFramesDe()

    'invisibiliza todos os frames
    Call Invisibiliza_FramesDe

    'seleciona o item para exibir o frame correlato
    Select Case Codigo_Extrai(TipoMeioPagtoDe.Text)

        Case TIPOMEIOPAGTOLOJA_DINHEIRO
            FrameDinDe.Visible = True
            FrameDinDe.Enabled = True

        Case TIPOMEIOPAGTOLOJA_CHEQUE
            FrameChqDe.Visible = True
            FrameChqDe.Enabled = True

        Case TIPOMEIOPAGTOLOJA_CARTAO_CREDITO
            ParcelamentoCrtDe.Clear
            Call Carrega_AdmMeioPagto_Cartao(AdmCrtDe, Codigo_Extrai(TipoMeioPagtoDe.Text))
            FrameCrtDe.Visible = True
            FrameCrtDe.Enabled = True
            OptionManualDe.Enabled = True

        Case TIPOMEIOPAGTOLOJA_CARTAO_DEBITO
            ParcelamentoCrtDe.Clear
            Call Carrega_AdmMeioPagto_Cartao(AdmCrtDe, Codigo_Extrai(TipoMeioPagtoDe.Text))
            FrameCrtDe.Visible = True
            FrameCrtDe.Enabled = True
            OptionManualDe.Enabled = False
            OptionPOSDe.Value = True
        
        Case TIPOMEIOPAGTOLOJA_VALE_TICKET
            FrameTktDe.Visible = True
            FrameTktDe.Enabled = True

        Case TIPOMEIOPAGTOLOJA_OUTROS
            FrameOutDe.Visible = True
            FrameOutDe.Enabled = True

    End Select

End Sub

Private Sub Invisibiliza_FramesDe()

    FrameDinDe.Visible = False
    FrameTktDe.Visible = False
    FrameOutDe.Visible = False
    FrameCrtDe.Visible = False
    FrameChqDe.Visible = False

    FrameDinDe.Enabled = False
    FrameTktDe.Enabled = False
    FrameOutDe.Enabled = False
    FrameCrtDe.Enabled = False
    FrameChqDe.Enabled = False

End Sub

Private Sub Carrega_AdmMeioPagto_Cartao(objComboBox As ComboBox, iTipoMeioPagto As Integer)

Dim lErro As Long
Dim colAdmMeioPagto As New Collection
Dim objAdmMeioPagto As ClassAdmMeioPagto

On Error GoTo Erro_Carrega_AdmMeioPagto_Cartao

    objComboBox.Clear
    
    If objComboBox.Name = "AdmCrtDe" Then
        giAdmCrtDe = 0
    ElseIf objComboBox.Name = "AdmCrtPara" Then
        giAdmCrtPara = 0
    End If
    
        'lê todas as admmeiopagto
        lErro = CF("AdmMeioPagto_Le_Todas", colAdmMeioPagto)
        If lErro <> SUCESSO And lErro <> 104031 Then gError 113518
        
        If lErro = 104031 Then gError 113519
        
        'varro a coleção de admmeiopagto lida
        For Each objAdmMeioPagto In colAdmMeioPagto
        
            'se for do tipo passado por parâmetro
            If objAdmMeioPagto.iTipoMeioPagto = iTipoMeioPagto Then
                
                'adiciona à combo
                objComboBox.AddItem (objAdmMeioPagto.iCodigo & SEPARADOR & objAdmMeioPagto.sNome)
                objComboBox.ItemData(objComboBox.NewIndex) = objAdmMeioPagto.iCodigo
            
            End If
        
        Next
    
    Exit Sub
    
Erro_Carrega_AdmMeioPagto_Cartao:
    
    Select Case gErr
    
        Case 113518
        
        Case 113519
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_VAZIA", gErr)
    
    End Select
    
    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim objTransfBrowser As New ClassTransfBrowser
Dim colSelecao As New Collection
Dim sSelecao As String

On Error GoTo Erro_LabelCodigo_Click

    'se o codigo estiver preenchido
    If Len(Trim(Codigo.ClipText)) > 0 Then

        'move o codigo para o obj
        objTransfBrowser.lTransferencia = StrParaLong(Codigo.Text)

    End If

    'Chama o Browser '
    Call Chama_Tela("TransfCentralLista", colSelecao, objTransfBrowser, objEventoTransf, sSelecao)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175436)

    End Select

    Exit Sub

End Sub

Private Sub LabelSeqChqDe_Click()

Dim objCheque As New ClassChequePre
Dim colSelecao As New Collection
Dim sSelecao As String

On Error GoTo Erro_LabelSeqChqDe_Click

    'Verifica se o Código não é Nulo se não armazena no objeto
    If Len(Trim(SeqChqDe.Text)) > 0 Then

        If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then

            objCheque.lSequencialLoja = StrParaLong(SeqChqDe.Text)
            objCheque.iFilialEmpresaLoja = giFilialEmpresa

        Else

            objCheque.lSequencialBack = StrParaLong(SeqChqDe.Text)

        End If

    End If

    'Chama o Browser ChequeCarneCupom_Lista
    Call Chama_Tela("ChequePreLojaLista", colSelecao, objCheque, objEventoCheque, sSelecao)

    Exit Sub

Erro_LabelSeqChqDe_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175437)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()
    Call TelaIndice_Preenche(Me)
End Sub

Public Sub Form_Deactivate()
    gi_ST_SetaIgnoraClick = 1
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub AgenciaChqPara_GotFocus()
    Call MaskEdBox_TrataGotFocus(SeqChqDe, iAlterado)
End Sub

Private Sub BancoChqPara_GotFocus()
    Call MaskEdBox_TrataGotFocus(BancoChqPara, iAlterado)
End Sub

Private Sub CarneChqPara_GotFocus()
    Call MaskEdBox_TrataGotFocus(CarneChqPara, iAlterado)
End Sub

Private Sub ClienteChqPara_GotFocus()
    Call MaskEdBox_TrataGotFocus(ClienteChqPara, iAlterado)
End Sub

Private Sub Codigo_GotFocus()
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'se cod estiver preenchido
    If Len(Trim(Codigo.ClipText)) > 0 Then

        'critica o conteudo
        lErro = Long_Critica(Codigo.Text)
        If lErro <> SUCESSO Then gError 113561

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 113561

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175438)

    End Select

    Exit Sub

End Sub

Private Sub ContaChqPara_GotFocus()
    Call MaskEdBox_TrataGotFocus(SeqChqDe, iAlterado)
End Sub

Private Sub CupomFiscalChqPara_GotFocus()
    Call MaskEdBox_TrataGotFocus(CupomFiscalChqPara, iAlterado)
End Sub

Private Sub DataBomParaChqPara_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataBomParaChqPara, iAlterado)
End Sub

Private Sub ECFChqPara_GotFocus()
    Call MaskEdBox_TrataGotFocus(ECFChqPara, iAlterado)
End Sub

Private Sub NumeroChqPara_GotFocus()
    Call MaskEdBox_TrataGotFocus(NumeroChqPara, iAlterado)
End Sub

Private Sub objEventoCheque_evSelecao(obj1 As Object)

Dim objChequePre As ClassChequePre
Dim lErro As Long

On Error GoTo Erro_objEventoCheque_evSelecao

    Set objChequePre = obj1

    'tenta pegar o cupom ou o carne associado ao cheque
    lErro = Trata_Cheque(objChequePre)
    If lErro <> SUCESSO Then gError 105287

    'Move os dados para a tela
    lErro = Traz_Cheque_Tela_De(objChequePre)
    If lErro <> SUCESSO And lErro <> 104342 Then gError 111513

    'Cheque não Encontrado no Banco de Dados
    If lErro = 104342 Then gError 104321

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoCheque_evSelecao:

    Select Case gErr

        Case 104321
            Call Rotina_Erro(vbOKOnly, "ERRO_CHEQUEPRE_INEXISTENTE", gErr, objChequePre.lNumIntCheque)

        Case 105287, 111513

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175439)

    End Select

    Exit Sub

End Sub

Private Function Traz_Cheque_Tela_De(objCheque As ClassChequePre) As Long
'traz os dados do cheque para a tela
'objCheque eh parametro de input

Dim sCPFCGC As String
Dim bCancel As Boolean

On Error GoTo Erro_Traz_Cheque_Tela_De

    Call Limpa_ChequeDe

    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then

        SeqChqDe.Text = objCheque.lSequencialLoja
        
    Else
    
        SeqChqDe.Text = objCheque.lSequencialBack
    
    End If

    'Se o Banco estiver preenchido
    If objCheque.iBanco > 0 Then BancoChqDe.Caption = CStr(objCheque.iBanco)

    'poe a agencia na tela
    AgenciaChqDe.Caption = objCheque.sAgencia

    'poe a conta na tela
    ContaChqDe.Caption = objCheque.sContaCorrente

    'Se o numero estiver preenchido
    If objCheque.lNumero > 0 Then NumeroChqDe.Caption = objCheque.lNumero

    sCPFCGC = ClienteChqPara.FormattedText

    'preenche o campo com o cpf/cgc formatado
    ClienteChqPara.PromptInclude = False
    ClienteChqPara.Text = objCheque.sCPFCGC
    ClienteChqPara.PromptInclude = True

    Call ClienteChqPara_Validate(bCancel)

    'poe o cgc/cpf do cliente na tela
    ClienteChqDe.Caption = ClienteChqPara.FormattedText
    
    ClienteChqPara.PromptInclude = False
    ClienteChqPara.Text = sCPFCGC
    ClienteChqPara.PromptInclude = True
    
    Call ClienteChqPara_Validate(bCancel)
    
    'se o ecf estiver preenchido
    If objCheque.iECF > 0 Then ECFChqDe.Caption = CStr(objCheque.iECF)
    
    'se o cupom fiscal estiver preenchido
    If objCheque.lCupomFiscal > 0 Then CupomFiscalChqDe.Caption = CStr(objCheque.lCupomFiscal)

    'preenche o carne
    CarneChqDe.Caption = objCheque.sCarne
    
    'poe a data do deposito na tela
    DataBomParaChqDe.Caption = Format(objCheque.dtDataDeposito, "dd/mm/yyyy")

    'coloca, FINALMENTE, o valor na tela.........
    ValorChqDe.Caption = Format(objCheque.dValor, "Standard")
    
    Traz_Cheque_Tela_De = SUCESSO

    Exit Function

Erro_Traz_Cheque_Tela_De:

    Traz_Cheque_Tela_De = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175440)

    End Select

    Exit Function

End Function

Private Sub objEventoCupomFiscal_evSelecao(obj1 As Object)

Dim objCupomFiscal As ClassCupomFiscal
Dim lErro As Long

On Error GoTo Erro_objEventoCupomFiscal_evSelecao

    Set objCupomFiscal = obj1

    ECFChqPara.Text = CStr(objCupomFiscal.iECF)
    
    CupomFiscalChqPara.Text = CStr(objCupomFiscal.lNumero)

    Me.Show

    Exit Sub

Erro_objEventoCupomFiscal_evSelecao:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175441)

    End Select

    Exit Sub

End Sub


Private Sub SeqChqDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(SeqChqDe, iAlterado)
End Sub

Private Sub ValorCrtDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ValorCrtDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(ValorCrtDe, iAlterado)
End Sub

Private Sub ValorCrtDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorCrtDe_Validate

    'se campo estiver preenchido
    If Len(Trim(ValorCrtDe.ClipText)) > 0 Then

        'critica o conteudo
        lErro = Valor_Positivo_Critica(ValorCrtDe.Text)
        If lErro <> SUCESSO Then gError 105335

    End If

    Exit Sub

Erro_ValorCrtDe_Validate:

    Cancel = True

    Select Case gErr

        Case 105335

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175442)

    End Select

    Exit Sub

End Sub

Private Sub ValorDinDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ValorDinDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(ValorDinDe, iAlterado)
End Sub

Private Sub ValorDinDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorDinDe_Validate

    'se campo estiver preenchido
    If Len(Trim(ValorDinDe.ClipText)) > 0 Then

        'critica o conteudo
        lErro = Valor_Positivo_Critica(ValorDinDe.Text)
        If lErro <> SUCESSO Then gError 113558

    End If

    Exit Sub

Erro_ValorDinDe_Validate:

    Cancel = True

    Select Case gErr

        Case 113558

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175443)

    End Select

    Exit Sub

End Sub

Private Sub ValorOutDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ValorOutDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(ValorOutDe, iAlterado)
End Sub

Private Sub ValorOutDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorOutDe_Validate

    'se campo estiver preenchido
    If Len(Trim(ValorOutDe.ClipText)) > 0 Then

        'critica o conteudo
        lErro = Valor_Positivo_Critica(ValorOutDe.Text)
        If lErro <> SUCESSO Then gError 113559

    End If

    Exit Sub

Erro_ValorOutDe_Validate:

    Cancel = True

    Select Case gErr

        Case 113559

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175444)

    End Select

    Exit Sub

End Sub

Private Sub ValorTktDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ValorTktDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(ValorTktDe, iAlterado)
End Sub

Private Sub ValorTktDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorTktDe_Validate

    'se campo estiver preenchido
    If Len(Trim(ValorTktDe.ClipText)) > 0 Then

        'critica o conteudo
        lErro = Valor_Positivo_Critica(ValorTktDe.Text)
        If lErro <> SUCESSO Then gError 113560

    End If

    Exit Sub

Erro_ValorTktDe_Validate:

    Cancel = True

    Select Case gErr

        Case 113560

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175445)

    End Select

    Exit Sub

End Sub

Private Sub AdmTktDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AdmTktDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objAdmMeioPagto As New ClassAdmMeioPagto

On Error GoTo Erro_AdmTktDe_Validate

    'se a adm de ticket estiver preenchida
    If Len(Trim(AdmTktDe.Text)) > 0 Then
    
        'tenta selecionar
        lErro = Combo_Seleciona(AdmTktDe, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 113520
        
        'preenche os atributos de busca da admmeiopagto
        objAdmMeioPagto.iCodigo = iCodigo
        objAdmMeioPagto.iFilialEmpresa = giFilialEmpresa
            
        'se retornou código
        If lErro = 6730 Then
        
            'tenta ler
            lErro = CF("AdmMeioPagto_Le", objAdmMeioPagto)
            If lErro <> SUCESSO And lErro <> 104017 Then gError 113521
            
            'se não encontrar-> erro
            If lErro = 104017 Then gError 113522
            
            If objAdmMeioPagto.iTipoMeioPagto <> TIPOMEIOPAGTOLOJA_VALE_TICKET Then gError 113536
            
            AdmTktDe.Text = objAdmMeioPagto.iCodigo & SEPARADOR & objAdmMeioPagto.sNome
        
        'se não retornou código-> erro
        ElseIf lErro = 6731 Then gError 113523
        
        End If
    
        'lê os parcelamentos da adm
        lErro = CF("AdmMeioPagtoCondPagto_Le", objAdmMeioPagto)
        If lErro <> SUCESSO And lErro <> 104086 Then gError 126057
        
        'carrega a combo de parcelamentos
        If lErro = 104086 Then gError 126058
        
        If giAdmTktDe <> iCodigo Then
        
            giAdmTktDe = iCodigo
            Call Carrega_Parcelamento(ParcelamentoTktDe, objAdmMeioPagto.colCondPagtoLoja)
    
        End If
        
    Else
    
        ParcelamentoTktDe.Clear
        giAdmTktDe = 0
    
    End If
    
    Exit Sub
    
Erro_AdmTktDe_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 113520, 113521, 126057
        
        Case 113522, 113536
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_NAO_CADASTRADO", gErr, objAdmMeioPagto.iCodigo)
            
        Case 113523
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_NAO_CADASTRADO", gErr, AdmTktDe.Text)
            
        Case 126058
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_SEM_PARCELAMENTO", gErr, objAdmMeioPagto.iCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175446)
    
    End Select
    
    Exit Sub

End Sub

Private Sub AdmTktPara_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AdmTktPara_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objAdmMeioPagto As New ClassAdmMeioPagto

On Error GoTo Erro_AdmTktPara_Validate

    'se a adm de ticket estiver preenchida
    If Len(Trim(AdmTktPara.Text)) > 0 Then
    
        'tenta selecionar
        lErro = Combo_Seleciona(AdmTktPara, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 113524
        
        'se retornou código
        If lErro = 6730 Then
        
            'preenche os atributos de busca da admmeiopagto
            objAdmMeioPagto.iCodigo = iCodigo
            objAdmMeioPagto.iFilialEmpresa = giFilialEmpresa
            
            'tenta ler
            lErro = CF("AdmMeioPagto_Le", objAdmMeioPagto)
            If lErro <> SUCESSO And lErro <> 104017 Then gError 113525
            
            'se não encontrar-> erro
            If lErro = 104017 Then gError 113526
            
            If objAdmMeioPagto.iTipoMeioPagto <> TIPOMEIOPAGTOLOJA_VALE_TICKET Then gError 113537
        
            AdmTktPara.Text = objAdmMeioPagto.iCodigo & SEPARADOR & objAdmMeioPagto.sNome
        
        'se não retornou código-> erro
        ElseIf lErro = 6731 Then gError 113527
        
        End If
    
    End If

    Exit Sub
    
Erro_AdmTktPara_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 113524, 113525
        
        Case 113526, 113537
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_NAO_CADASTRADO", gErr, objAdmMeioPagto.iCodigo)
            
        Case 113527
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_NAO_CADASTRADO", gErr, AdmTktPara.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175447)
    
    End Select
    
    Exit Sub

End Sub

Private Sub AdmCrtDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Carrega_Parcelamento(objComboBox As ComboBox, colAdmMeioPagtoCondPagto As Collection)

Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
    
    objComboBox.Clear
    
    For Each objAdmMeioPagtoCondPagto In colAdmMeioPagtoCondPagto
    
        objComboBox.AddItem (objAdmMeioPagtoCondPagto.iParcelamento & SEPARADOR & objAdmMeioPagtoCondPagto.sNomeParcelamento)
        objComboBox.ItemData(objComboBox.NewIndex) = objAdmMeioPagtoCondPagto.iParcelamento
    
    Next

End Sub

Private Sub AdmCrtDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objAdmMeioPagto As New ClassAdmMeioPagto

On Error GoTo Erro_AdmCrtDe_Validate

    'se a adm de ticket estiver preenchida
    If Len(Trim(AdmCrtDe.Text)) > 0 Then
    
        'tenta selecionar
        lErro = Combo_Seleciona(AdmCrtDe, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 113540
        
        'preenche os atributos de busca da admmeiopagto
        objAdmMeioPagto.iCodigo = iCodigo
        objAdmMeioPagto.iFilialEmpresa = giFilialEmpresa
       
        'se retornou código
        If lErro = 6730 Then
        
            'tenta ler
            lErro = CF("AdmMeioPagto_Le", objAdmMeioPagto)
            If lErro <> SUCESSO And lErro <> 104017 Then gError 113541
            
            'se não encontrar-> erro
            If lErro = 104017 Then gError 113542
            
            'se o tipo não for cartão de crédito/débito (o que estiver selecionado)-> erro
            If objAdmMeioPagto.iTipoMeioPagto <> Codigo_Extrai(TipoMeioPagtoDe.Text) Then gError 113543
            
            AdmCrtDe.Text = objAdmMeioPagto.iCodigo & SEPARADOR & objAdmMeioPagto.sNome
        
        'se não retornou código-> erro
        ElseIf lErro = 6731 Then gError 113544
        
        End If
        
        'lê os parcelamentos da adm
        lErro = CF("AdmMeioPagtoCondPagto_Le", objAdmMeioPagto)
        If lErro <> SUCESSO And lErro <> 104086 Then gError 113550
        
        'carrega a combo de parcelamentos
        If lErro = 104086 Then gError 113551
        
        If giAdmCrtDe <> iCodigo Then
        
            giAdmCrtDe = iCodigo
            Call Carrega_Parcelamento(ParcelamentoCrtDe, objAdmMeioPagto.colCondPagtoLoja)
    
        End If
        
    Else
    
        ParcelamentoCrtDe.Clear
        giAdmCrtDe = 0
    
    End If

    Exit Sub
    
Erro_AdmCrtDe_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 113540, 113541, 113550
        
        Case 113542, 113543
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_NAO_CADASTRADO", gErr, objAdmMeioPagto.iCodigo)
            
        Case 113544
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_NAO_CADASTRADO", gErr, AdmCrtDe.Text)
            
        Case 113551
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_SEM_PARCELAMENTO", gErr, objAdmMeioPagto.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175448)
    
    End Select
    
    Exit Sub

End Sub

Private Sub AdmCrtPara_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AdmCrtPara_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objAdmMeioPagto As New ClassAdmMeioPagto

On Error GoTo Erro_AdmCrtPara_Validate

    'se a adm de ticket estiver preenchida
    If Len(Trim(AdmCrtPara.Text)) > 0 Then
    
        'tenta selecionar
        lErro = Combo_Seleciona(AdmCrtPara, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 113545
        
        'preenche os atributos de busca da admmeiopagto
        objAdmMeioPagto.iCodigo = iCodigo
        objAdmMeioPagto.iFilialEmpresa = giFilialEmpresa
        
        'se retornou código
        If lErro = 6730 Then
        
            'tenta ler
            lErro = CF("AdmMeioPagto_Le", objAdmMeioPagto)
            If lErro <> SUCESSO And lErro <> 104017 Then gError 113546
            
            'se não encontrar-> erro
            If lErro = 104017 Then gError 113547
            
            If objAdmMeioPagto.iTipoMeioPagto <> Codigo_Extrai(TipoMeioPagtoPara.Text) Then gError 113548
            
            AdmCrtPara.Text = objAdmMeioPagto.iCodigo & SEPARADOR & objAdmMeioPagto.sNome
        
        'se não retornou código-> erro
        ElseIf lErro = 6731 Then gError 113549
        
        End If
        
        'lê os parcelamentos da adm
        lErro = CF("AdmMeioPagtoCondPagto_Le", objAdmMeioPagto)
        If lErro <> SUCESSO And lErro <> 104086 Then gError 113552
        
        If lErro = 104086 Then gError 113553
        
        If giAdmCrtPara <> iCodigo Then
        
            giAdmCrtPara = iCodigo
        
            'carrega a combo de parcelamentos
            Call Carrega_Parcelamento(ParcelamentoCrtPara, objAdmMeioPagto.colCondPagtoLoja)

        End If

    Else
    
        ParcelamentoCrtPara.Clear
        giAdmCrtPara = 0

    
    End If

    Exit Sub
    
Erro_AdmCrtPara_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 113545, 113546, 113552
        
        Case 113547, 113548
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_NAO_CADASTRADO", gErr, objAdmMeioPagto.iCodigo)
            
        Case 113553
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_SEM_PARCELAMENTO", gErr, objAdmMeioPagto.iCodigo)
            
        Case 113549
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_NAO_CADASTRADO", gErr, AdmCrtPara.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175449)
    
    End Select
    
    Exit Sub

End Sub

Private Sub AdmOutDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AdmOutDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objAdmMeioPagto As New ClassAdmMeioPagto

On Error GoTo Erro_AdmOutDe_Validate

    'se a adm de ticket estiver preenchida
    If Len(Trim(AdmOutDe.Text)) > 0 Then
    
        'tenta selecionar
        lErro = Combo_Seleciona(AdmOutDe, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 113528
        
        'preenche os atributos de busca da admmeiopagto
        objAdmMeioPagto.iCodigo = iCodigo
        objAdmMeioPagto.iFilialEmpresa = giFilialEmpresa
        
        'se retornou código
        If lErro = 6730 Then
        
            'tenta ler
            lErro = CF("AdmMeioPagto_Le", objAdmMeioPagto)
            If lErro <> SUCESSO And lErro <> 104017 Then gError 113529
            
            'se não encontrar-> erro
            If lErro = 104017 Then gError 113530
            
            If objAdmMeioPagto.iTipoMeioPagto <> TIPOMEIOPAGTOLOJA_OUTROS Then gError 113538
        
            AdmOutDe.Text = objAdmMeioPagto.iCodigo & SEPARADOR & objAdmMeioPagto.sNome
        
        'se não retornou código-> erro
        ElseIf lErro = 6731 Then gError 113531
        
        End If
    
        'lê os parcelamentos da adm
        lErro = CF("AdmMeioPagtoCondPagto_Le", objAdmMeioPagto)
        If lErro <> SUCESSO And lErro <> 104086 Then gError 126073
        
        'carrega a combo de parcelamentos
        If lErro = 104086 Then gError 126074
        
        If giAdmOutDe <> iCodigo Then
        
            giAdmOutDe = iCodigo
            Call Carrega_Parcelamento(ParcelamentoOutDe, objAdmMeioPagto.colCondPagtoLoja)
    
        End If
        
    Else
    
        ParcelamentoOutDe.Clear
        giAdmOutDe = 0
    
    End If

    Exit Sub
    
Erro_AdmOutDe_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 113528, 113529, 126073
        
        Case 113530, 113538
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_NAO_CADASTRADO", gErr, objAdmMeioPagto.iCodigo)
            
        Case 113531
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_NAO_CADASTRADO", gErr, AdmOutDe.Text)
        
        Case 126074
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_SEM_PARCELAMENTO", gErr, objAdmMeioPagto.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175450)
    
    End Select
    
    Exit Sub

End Sub

Private Sub AdmOutPara_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AdmOutPara_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objAdmMeioPagto As New ClassAdmMeioPagto

On Error GoTo Erro_AdmOutPara_Validate

    'se a adm de ticket estiver preenchida
    If Len(Trim(AdmOutPara.Text)) > 0 Then
    
        'tenta selecionar
        lErro = Combo_Seleciona(AdmOutPara, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 113532
        
        'se retornou código
        If lErro = 6730 Then
        
            'preenche os atributos de busca da admmeiopagto
            objAdmMeioPagto.iCodigo = iCodigo
            objAdmMeioPagto.iFilialEmpresa = giFilialEmpresa
            
            'tenta ler
            lErro = CF("AdmMeioPagto_Le", objAdmMeioPagto)
            If lErro <> SUCESSO And lErro <> 104017 Then gError 113533
            
            'se não encontrar-> erro
            If lErro = 104017 Then gError 113534
                        
            If objAdmMeioPagto.iTipoMeioPagto <> TIPOMEIOPAGTOLOJA_OUTROS Then gError 113539
        
            AdmOutPara.Text = objAdmMeioPagto.iCodigo & SEPARADOR & objAdmMeioPagto.sNome
        
        'se não retornou código-> erro
        ElseIf lErro = 6731 Then gError 113535
        
        End If
    
    End If

    Exit Sub
    
Erro_AdmOutPara_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 113532, 113533
        
        Case 113534, 113539
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_NAO_CADASTRADO", gErr, objAdmMeioPagto.iCodigo)
            
        Case 113535
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_NAO_CADASTRADO", gErr, AdmOutPara.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175451)
    
    End Select
    
    Exit Sub

End Sub

Private Sub AgenciaChqPara_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub BancoChqPara_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub Form_Unload()
'liberar as variaveis globais
   
   Call ComandoSeta_Liberar(Me.Name)
   Set objEventoCheque = Nothing
   Set objEventoTransf = Nothing

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click

    'testa se houve alteração
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 113680

    'limpa os campos que se fizerem necessários
    Call Limpa_Tela_TransfCentral

    'fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then gError 103759

    iAlterado = 0

    Exit Sub
    
Erro_Botaolimpar_Click:
    
    Select Case gErr
    
        Case 113680
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175452)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Limpa_Tela_TransfCentral()

Dim objCombo As Object

    Call Limpa_Tela(Me)

    'limpa as combos
    For Each objCombo In Me.Controls

      'se for realmente combo
      If TypeName(objCombo) = "ComboBox" Then
         objCombo.ListIndex = -1
      End If
    
    Next
    
'    OptionManualDe.Value = True
    OptionManualPara.Value = True

    'coloca todos os frames invisiveis
    Call Invisibiliza_Frames
    
    Call Limpa_ChequeDe
    
End Sub

Private Sub Invisibiliza_Frames()

    Call Invisibiliza_FramesDe
    Call Invisibiliza_FramesPara

End Sub

Private Sub CarneChqPara_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ClienteChqPara_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Codigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ContaChqPara_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CupomFiscalChqPara_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataBomParaChqPara_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ECFChqPara_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NumeroChqPara_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

'Private Sub OptionManualDe_Click()
'    iAlterado = REGISTRO_ALTERADO
'End Sub

Private Sub OptionManualPara_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

'Private Sub OptionPOSDe_Click()
'    iAlterado = REGISTRO_ALTERADO
'End Sub

Private Sub OptionPOSPara_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ParcelamentoCrtDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ParcelamentoCrtPara_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub SeqChqDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub TipoMeioPagtoPara_Click()

    iAlterado = REGISTRO_ALTERADO

    'chama o tratamento de frames
    If TipoMeioPagtoPara.ListIndex <> -1 Then Call TipoMeioPagto_TrataFramesPara

End Sub

Private Sub TipoMeioPagto_TrataFramesPara()

    'invisibiliza todos os frames
    Call Invisibiliza_FramesPara

    'seleciona o item para exibir o frame correlato
    Select Case Codigo_Extrai(TipoMeioPagtoPara.Text)

        Case TIPOMEIOPAGTOLOJA_DINHEIRO
            FrameDinPara.Visible = True
            FrameDinPara.Enabled = True

        Case TIPOMEIOPAGTOLOJA_CHEQUE
            FrameChqPara.Visible = True
            FrameChqPara.Enabled = True

        Case TIPOMEIOPAGTOLOJA_CARTAO_CREDITO
            ParcelamentoCrtPara.Clear
            Call Carrega_AdmMeioPagto_Cartao(AdmCrtPara, Codigo_Extrai(TipoMeioPagtoPara.Text))
            FrameCrtPara.Visible = True
            FrameCrtPara.Enabled = True
            OptionManualPara.Enabled = True

        Case TIPOMEIOPAGTOLOJA_CARTAO_DEBITO
            ParcelamentoCrtPara.Clear
            Call Carrega_AdmMeioPagto_Cartao(AdmCrtPara, Codigo_Extrai(TipoMeioPagtoPara.Text))
            FrameCrtPara.Visible = True
            FrameCrtPara.Enabled = True
            OptionManualPara.Enabled = False
            OptionPOSPara.Value = True
        
        Case TIPOMEIOPAGTOLOJA_VALE_TICKET
            FrameTktPara.Visible = True
            FrameTktPara.Enabled = True

        Case TIPOMEIOPAGTOLOJA_OUTROS
            FrameOutPara.Visible = True
            FrameOutPara.Enabled = True

    End Select

End Sub

Private Sub Invisibiliza_FramesPara()

    FrameDinPara.Visible = False
    FrameTktPara.Visible = False
    FrameOutPara.Visible = False
    FrameCrtPara.Visible = False
    FrameChqPara.Visible = False


    FrameDinPara.Enabled = False
    FrameTktPara.Enabled = False
    FrameOutPara.Enabled = False
    FrameCrtPara.Enabled = False
    FrameChqPara.Enabled = False

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Transferência de Caixa"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TransfCentral"

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

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai o Caixa da tela

Dim lErro As Long
Dim objTransfCaixa As New ClassTransfCaixa
Dim objChequeDe As New ClassChequePre
Dim objChequePara As New ClassChequePre

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TransfCaixa"

    'le os dados da tela
    lErro = Move_Tela_Memoria(objTransfCaixa, objChequeDe, objChequePara)
    If lErro <> SUCESSO Then gError 111527

    'Preenche a coleção colCampoValor com a chave
    colCampoValor.Add "Codigo", objTransfCaixa.lCodigo, 0, "Codigo"
    colCampoValor.Add "FilialEmpresa", objTransfCaixa.iFilialEmpresa, 0, "FilialEmpresa"

    'Faz o filtro dos dados que serão exibidos
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 111527

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175453)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objTransfCaixa As New ClassTransfCaixa

On Error GoTo Erro_Tela_Preenche

    'Carrega objCaixa com os dados passados em colCampoValor
    objTransfCaixa.lCodigo = colCampoValor.Item("Codigo").vValor
    objTransfCaixa.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor

    lErro = Traz_TransfCaixa_Tela(objTransfCaixa)
    If lErro <> SUCESSO And lErro <> 105277 Then gError 111529

    If lErro <> SUCESSO Then gError 111530
    
    iAlterado = 0

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 111529

        Case 111530
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSFCAIXA_NAO_CADASTRADO", gErr, objTransfCaixa.objMovCaixaDe.lTransferencia)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175454)

    End Select

    Exit Sub

End Sub

Private Function Trata_Cheque(objChequePre As ClassChequePre) As Long
'tenta pegar o cupom ou o carne associado ao cheque

Dim objCupomFiscal As New ClassCupomFiscal
Dim lErro As Long

On Error GoTo Erro_Trata_Cheque


    'lê o cheque
    lErro = CF("Cheque_Le_Excluido", objChequePre)
    If lErro <> SUCESSO And lErro <> 105310 And lErro <> 105313 Then gError 105284
    
    'se não encontrar-> erro
    If lErro = 105310 Or lErro = 105313 Then gError 105285
    
    If objChequePre.lCupomFiscal <> 0 Then
        
        objCupomFiscal.iFilialEmpresa = giFilialEmpresa
        objCupomFiscal.lNumIntDoc = objChequePre.lCupomFiscal
    
        'tenta ler o cupom fiscal para trazer o ECF e o COO
        lErro = CF("CupomFiscal_Le_NumIntDoc", objCupomFiscal)
        If lErro <> SUCESSO And lErro <> 105268 Then gError 105286
        
        'se nao encontrou o cupom ==> erro
        If lErro = 105268 Then gError 105287
        
        objChequePre.lCupomFiscal = objCupomFiscal.lNumero
        objChequePre.iECF = objCupomFiscal.iECF
    
    Else
    
        'le o cheque se estiver vinculado a carne
        lErro = CF("ChequePre_Le_Carne", objChequePre)
        If lErro <> SUCESSO And lErro <> 105256 Then gError 105288
    
    End If

    Trata_Cheque = SUCESSO
    
    Exit Function
    
Erro_Trata_Cheque:
    
    Trata_Cheque = gErr
    
    Select Case gErr
    
        Case 105284, 105286, 105288
        
        Case 105285
            Call Rotina_Erro(vbOKOnly, "ERRO_CHEQUEPRE_NAOENCONTRADO", gErr, objChequePre.iFilialEmpresaLoja, objChequePre.lSequencial)
        
        Case 105287
            Call Rotina_Erro(vbOKOnly, "ERRO_CUPOM_FISCAL_NAO_CADASTRADO1", gErr, objCupomFiscal.lNumIntDoc)
        
        Case 105289
            Call Rotina_Erro(vbOKOnly, "ERRO_CHEQUE_NAO_CUPOM_CARNE", gErr, objCupomFiscal.lNumIntDoc)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175455)
    
    End Select

    Exit Function

End Function

Private Sub Limpa_ChequeDe()

    BancoChqDe.Caption = ""
    AgenciaChqDe.Caption = ""
    ContaChqDe.Caption = ""
    NumeroChqDe.Caption = ""
    ClienteChqDe.Caption = ""
    ECFChqDe.Caption = ""
    CupomFiscalChqDe.Caption = ""
    CarneChqDe.Caption = ""
    DataBomParaChqDe.Caption = ""
    ValorChqDe.Caption = ""
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
    
        If Me.ActiveControl Is Codigo Then Call LabelCodigo_Click
        
        If Me.ActiveControl Is SeqChqDe Then Call LabelSeqChqDe_Click
        
        If Me.ActiveControl Is ECFChqPara Then Call LabelECFChqPara_Click
        
        If Me.ActiveControl Is CupomFiscalChqPara Then Call LabelCupomFiscalChqPara_Click
        
    End If

End Sub

