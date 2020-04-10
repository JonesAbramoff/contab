VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl TransfCaixa 
   ClientHeight    =   6150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9060
   KeyPreview      =   -1  'True
   ScaleHeight     =   6150
   ScaleWidth      =   9060
   Begin VB.CommandButton Command1 
      Caption         =   "Desmembrar"
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
      Left            =   3285
      TabIndex        =   68
      Top             =   180
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6795
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "Transfcaixa.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   86
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   645
         Picture         =   "Transfcaixa.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1155
         Picture         =   "Transfcaixa.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1650
         Picture         =   "Transfcaixa.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   89
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Para"
      Height          =   5295
      Left            =   4575
      TabIndex        =   29
      Top             =   765
      Width           =   4380
      Begin VB.Frame FrameCrtPara 
         BorderStyle     =   0  'None
         Caption         =   "FrameCrtDe"
         Height          =   4470
         Left            =   120
         TabIndex        =   58
         Top             =   780
         Visible         =   0   'False
         Width           =   4215
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
            Left            =   1845
            TabIndex        =   61
            Top             =   930
            Value           =   -1  'True
            Width           =   975
         End
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
            Left            =   3015
            TabIndex        =   64
            Top             =   945
            Width           =   735
         End
         Begin VB.ComboBox AdmCrtPara 
            Height          =   315
            ItemData        =   "Transfcaixa.ctx":0994
            Left            =   1860
            List            =   "Transfcaixa.ctx":0996
            TabIndex        =   53
            Top             =   15
            Width           =   2250
         End
         Begin VB.ComboBox ParcelamentoCrtPara 
            Height          =   315
            ItemData        =   "Transfcaixa.ctx":0998
            Left            =   1860
            List            =   "Transfcaixa.ctx":099A
            TabIndex        =   56
            Top             =   465
            Width           =   2250
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
            Left            =   945
            TabIndex        =   69
            Top             =   930
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ca&rtão:"
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
            Index           =   24
            Left            =   1155
            TabIndex        =   51
            Top             =   75
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Parc&elamento:"
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
            TabIndex        =   54
            Top             =   495
            Width           =   1230
         End
      End
      Begin VB.Frame FrameChqPara 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   4470
         Left            =   120
         TabIndex        =   30
         Top             =   780
         Visible         =   0   'False
         Width           =   4215
         Begin MSMask.MaskEdBox CupomFiscalChqPara 
            Height          =   300
            Left            =   1875
            TabIndex        =   50
            Top             =   3000
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox AgenciaChqPara 
            Height          =   300
            Left            =   1875
            TabIndex        =   37
            Top             =   450
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
            Left            =   1860
            TabIndex        =   39
            Top             =   870
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
            Left            =   1860
            TabIndex        =   45
            Top             =   2175
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
            Left            =   1860
            TabIndex        =   43
            ToolTipText     =   "CGC/CPF do Cliente"
            Top             =   1725
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
            Left            =   1860
            TabIndex        =   41
            Top             =   1305
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
            Left            =   1875
            TabIndex        =   35
            Top             =   15
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   3
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ECFChqPara 
            Height          =   300
            Left            =   1875
            TabIndex        =   47
            Top             =   2595
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#########"
            PromptChar      =   " "
         End
         Begin VB.Label LabelECF 
            AutoSize        =   -1  'True
            Caption         =   "&ECF:"
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
            Left            =   1380
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   46
            Top             =   2655
            Width           =   420
         End
         Begin VB.Label LabelCupomFiscalChqPara 
            Caption         =   "Cupom &Fiscal (COO):"
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
            Left            =   15
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   48
            Top             =   3045
            Width           =   1770
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "&Banco:"
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
            Left            =   1185
            TabIndex        =   34
            Top             =   60
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "C&liente:"
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
            Left            =   1125
            TabIndex        =   42
            Top             =   1785
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Bom &Para:"
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
            Left            =   915
            TabIndex        =   44
            Top             =   2220
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "&Número:"
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
            Left            =   1065
            TabIndex        =   40
            Top             =   1350
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "&Conta:"
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
            Left            =   1215
            TabIndex        =   38
            Top             =   930
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Agê&ncia:"
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
            Left            =   1035
            TabIndex        =   36
            Top             =   495
            Width           =   765
         End
      End
      Begin VB.Frame FrameOutPara 
         BorderStyle     =   0  'None
         Caption         =   "FrameCrtDe"
         Height          =   4470
         Left            =   120
         TabIndex        =   67
         Top             =   780
         Visible         =   0   'False
         Width           =   4215
         Begin VB.ComboBox AdmOutPara 
            Height          =   315
            ItemData        =   "Transfcaixa.ctx":099C
            Left            =   1860
            List            =   "Transfcaixa.ctx":099E
            TabIndex        =   59
            Top             =   135
            Width           =   2220
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Meio Pa&gto:"
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
            TabIndex        =   57
            Top             =   180
            Width           =   1035
         End
      End
      Begin VB.Frame FrameTktPara 
         BorderStyle     =   0  'None
         Caption         =   "FrameCrtDe"
         Height          =   4470
         Left            =   120
         TabIndex        =   65
         Top             =   780
         Visible         =   0   'False
         Width           =   4215
         Begin VB.ComboBox AdmTktPara 
            Height          =   315
            ItemData        =   "Transfcaixa.ctx":09A0
            Left            =   1860
            List            =   "Transfcaixa.ctx":09A2
            TabIndex        =   62
            Top             =   135
            Width           =   2220
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tic&ket:"
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
            Left            =   1155
            TabIndex        =   60
            Top             =   180
            Width           =   615
         End
      End
      Begin VB.Frame FrameDinPara 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   4470
         Left            =   135
         TabIndex        =   52
         Top             =   795
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.ComboBox TipoMeioPagtoPara 
         Height          =   315
         ItemData        =   "Transfcaixa.ctx":09A4
         Left            =   1995
         List            =   "Transfcaixa.ctx":09A6
         Style           =   2  'Dropdown List
         TabIndex        =   33
         ToolTipText     =   "Tipo do meio de pagamento"
         Top             =   315
         Width           =   2220
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "T&ipo Meio Pagto:"
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
         Left            =   450
         TabIndex        =   32
         ToolTipText     =   "Tipo do meio de pagamento"
         Top             =   375
         Width           =   1485
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "De"
      Height          =   5295
      Left            =   135
      TabIndex        =   27
      Top             =   765
      Width           =   4380
      Begin VB.Frame FrameCrtDe 
         BorderStyle     =   0  'None
         Caption         =   "FrameCrtDe"
         Height          =   4620
         Left            =   90
         TabIndex        =   55
         Top             =   645
         Visible         =   0   'False
         Width           =   4230
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
            Left            =   3060
            TabIndex        =   16
            Top             =   1605
            Width           =   735
         End
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
            Left            =   1920
            TabIndex        =   15
            Top             =   1575
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.ComboBox ParcelamentoCrtDe 
            Height          =   315
            ItemData        =   "Transfcaixa.ctx":09A8
            Left            =   1875
            List            =   "Transfcaixa.ctx":09AA
            TabIndex        =   11
            Top             =   625
            Width           =   2340
         End
         Begin VB.ComboBox AdmCrtDe 
            Height          =   315
            ItemData        =   "Transfcaixa.ctx":09AC
            Left            =   1890
            List            =   "Transfcaixa.ctx":09AE
            TabIndex        =   9
            Top             =   165
            Width           =   2340
         End
         Begin MSMask.MaskEdBox ValorCrtDe 
            Height          =   300
            Left            =   1875
            TabIndex        =   13
            Top             =   1095
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
            Format          =   "#,##0.00"
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
            Left            =   1005
            TabIndex        =   14
            Top             =   1560
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Parc&elamento:"
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
            TabIndex        =   10
            Top             =   645
            Width           =   1230
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "&Valor:"
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
            TabIndex        =   12
            Top             =   1105
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ca&rtão:"
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
            TabIndex        =   8
            Top             =   195
            Width           =   630
         End
      End
      Begin VB.Frame FrameChqDe 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   4620
         Left            =   120
         TabIndex        =   28
         Top             =   645
         Visible         =   0   'False
         Width           =   4230
         Begin VB.CommandButton BotaoChequeDe 
            Caption         =   "(F9)  Cheques"
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
            Left            =   1335
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   3990
            Width           =   1605
         End
         Begin MSMask.MaskEdBox SeqChqDe 
            Height          =   300
            Left            =   1860
            TabIndex        =   6
            Top             =   120
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
            Left            =   0
            TabIndex        =   85
            Top             =   2955
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
            Index           =   4
            Left            =   1110
            TabIndex        =   84
            Top             =   2145
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
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   885
            TabIndex        =   83
            Top             =   2565
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
            Index           =   1
            Left            =   1050
            TabIndex        =   82
            Top             =   1740
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
            Index           =   5
            Left            =   1200
            TabIndex        =   81
            Top             =   1365
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
            Index           =   3
            Left            =   1005
            TabIndex        =   80
            Top             =   960
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
            Index           =   0
            Left            =   1155
            TabIndex        =   79
            Top             =   555
            Width           =   615
         End
         Begin VB.Label LabelSeqChqDe 
            AutoSize        =   -1  'True
            Caption         =   "&Sequencial:"
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
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   5
            Top             =   180
            Width           =   1020
         End
         Begin VB.Label BancoChqDe 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1860
            TabIndex        =   78
            Top             =   510
            Width           =   870
         End
         Begin VB.Label AgenciaChqDe 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1860
            TabIndex        =   77
            Top             =   915
            Width           =   870
         End
         Begin VB.Label ContaChqDe 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1860
            TabIndex        =   76
            Top             =   1305
            Width           =   1395
         End
         Begin VB.Label NumeroChqDe 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1860
            TabIndex        =   75
            Top             =   1695
            Width           =   1395
         End
         Begin VB.Label ClienteChqDe 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1860
            TabIndex        =   74
            Top             =   2100
            Width           =   1680
         End
         Begin VB.Label CupomFiscalChqDe 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1860
            TabIndex        =   73
            Top             =   2910
            Width           =   1185
         End
         Begin VB.Label DataBomParaChqDe 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1860
            TabIndex        =   72
            Top             =   2505
            Width           =   1170
         End
         Begin VB.Label ValorChqDe 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1860
            TabIndex        =   71
            Top             =   3315
            Width           =   1260
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
            Left            =   1260
            TabIndex        =   70
            Top             =   3360
            Width           =   510
         End
      End
      Begin VB.Frame FrameTktDe 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   4620
         Left            =   105
         TabIndex        =   63
         Top             =   660
         Visible         =   0   'False
         Width           =   4230
         Begin MSMask.MaskEdBox ValorTktDe 
            Height          =   300
            Left            =   1860
            TabIndex        =   26
            Top             =   570
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.ComboBox AdmTktDe 
            Height          =   315
            ItemData        =   "Transfcaixa.ctx":09B0
            Left            =   1875
            List            =   "Transfcaixa.ctx":09B2
            TabIndex        =   24
            Top             =   135
            Width           =   2250
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tic&ket:"
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
            TabIndex        =   23
            Top             =   180
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Val&or:"
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
            TabIndex        =   25
            Top             =   615
            Width           =   510
         End
      End
      Begin VB.Frame FrameOutDe 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   4620
         Left            =   105
         TabIndex        =   66
         Top             =   645
         Visible         =   0   'False
         Width           =   4230
         Begin VB.ComboBox AdmOutDe 
            Height          =   315
            ItemData        =   "Transfcaixa.ctx":09B4
            Left            =   1875
            List            =   "Transfcaixa.ctx":09B6
            TabIndex        =   20
            Top             =   120
            Width           =   2250
         End
         Begin MSMask.MaskEdBox ValorOutDe 
            Height          =   300
            Left            =   1860
            TabIndex        =   22
            Top             =   570
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Va&lor:"
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
            Left            =   1260
            TabIndex        =   21
            Top             =   600
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Meio Pa&gto:"
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
            TabIndex        =   19
            Top             =   180
            Width           =   1035
         End
      End
      Begin VB.Frame FrameDinDe 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   4620
         Left            =   105
         TabIndex        =   49
         Top             =   645
         Visible         =   0   'False
         Width           =   4230
         Begin MSMask.MaskEdBox ValorDinDe 
            Height          =   300
            Left            =   1875
            TabIndex        =   18
            Top             =   165
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "V&alor:"
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
            Index           =   17
            Left            =   1290
            TabIndex        =   17
            Top             =   180
            Width           =   510
         End
      End
      Begin VB.ComboBox TipoMeioPagtoDe 
         Height          =   315
         ItemData        =   "Transfcaixa.ctx":09B8
         Left            =   1980
         List            =   "Transfcaixa.ctx":09BA
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Tipo do meio de pagamento"
         Top             =   315
         Width           =   2220
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Tipo Meio Pagto:"
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
         Left            =   435
         TabIndex        =   3
         ToolTipText     =   "Tipo do meio de pagamento"
         Top             =   360
         Width           =   1470
      End
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   300
      Left            =   2400
      Picture         =   "Transfcaixa.ctx":09BC
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Numeração Automática"
      Top             =   300
      Width           =   300
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Top             =   315
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
      Caption         =   "Có&digo:"
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
      Left            =   570
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   0
      Top             =   345
      Width           =   660
   End
End
Attribute VB_Name = "TransfCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_Caption As String

Event Unload()


Const TRANSFERENCIA_LOJA = 1


Public iAlterado As Integer

Private Sub Form_Load()
'Inicializacao da tela

On Error GoTo Erro_Form_Load

    'Carrega as combos de admmeiopagto (credicard, visa, VR, TR, Ticket etc)
    Call Carrega_AdmMeioPagto_Nao_Cartao

    'carrega as combos de tipomeiopagto (Dinheiro, Cheque, Cartao, etc)
    Call Carrega_TipoMeioPagto

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175321)

    End Select

    Exit Sub

End Sub

Private Sub Carrega_AdmMeioPagto_Nao_Cartao()

Dim objAdmMeioPagto As ClassAdmMeioPagto

On Error GoTo Erro_Carrega_AdmMeioPagto_Nao_Cartao

    'varre a coleção global de AdmMeioPagto
    For Each objAdmMeioPagto In gcolAdmMeioPagto

        'verifica o tipomeiopagto atual
        Select Case objAdmMeioPagto.iTipoMeioPagto

            'se for ticket, adiciona à combo de valeticket(De e Para)
            Case TIPOMEIOPAGTOLOJA_VALE_TICKET, TIPOMEIOPAGTOLOJA_VALE_REFEICAO, TIPOMEIOPAGTOLOJA_VALE_PRESENTE, TIPOMEIOPAGTOLOJA_VALE_COMBUSTIVEL
                AdmTktDe.AddItem (objAdmMeioPagto.iCodigo & SEPARADOR & objAdmMeioPagto.sNome)
                AdmTktDe.ItemData(AdmTktDe.NewIndex) = objAdmMeioPagto.iCodigo

                AdmTktPara.AddItem (objAdmMeioPagto.iCodigo & SEPARADOR & objAdmMeioPagto.sNome)
                AdmTktPara.ItemData(AdmTktPara.NewIndex) = objAdmMeioPagto.iCodigo

            'se for outros, adiciona à combo de outros(De e Para)
            Case TIPOMEIOPAGTOLOJA_OUTROS
                AdmOutDe.AddItem (objAdmMeioPagto.iCodigo & SEPARADOR & objAdmMeioPagto.sNome)
                AdmOutDe.ItemData(AdmOutDe.NewIndex) = objAdmMeioPagto.iCodigo

                AdmOutPara.AddItem (objAdmMeioPagto.iCodigo & SEPARADOR & objAdmMeioPagto.sNome)
                AdmOutPara.ItemData(AdmOutPara.NewIndex) = objAdmMeioPagto.iCodigo

            End Select

    Next

    Exit Sub

Erro_Carrega_AdmMeioPagto_Nao_Cartao:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175322)

    End Select

    Exit Sub

End Sub

Private Sub Carrega_TipoMeioPagto()

Dim objTipoMeioPagto As ClassTMPLoja

On Error GoTo Erro_Carrega_TipoMeioPagto

    'para cada tipomeiopagto em colmeiopagto
    For Each objTipoMeioPagto In gcolTiposMeiosPagtos

        'se o tipo de for passivel de transferencia
        If objTipoMeioPagto.iTransferencia = TRANSFERENCIA_LOJA Then

            'insere os itens nas combos de tipomeiopagto de e para
            TipoMeioPagtoDe.AddItem (objTipoMeioPagto.iTipo & SEPARADOR & objTipoMeioPagto.sDescricao)
            TipoMeioPagtoPara.AddItem (objTipoMeioPagto.iTipo & SEPARADOR & objTipoMeioPagto.sDescricao)

            'carrega os itemdatas das combos tipomeiopagto de e para
            TipoMeioPagtoDe.ItemData(TipoMeioPagtoDe.NewIndex) = objTipoMeioPagto.iTipo
            TipoMeioPagtoPara.ItemData(TipoMeioPagtoPara.NewIndex) = objTipoMeioPagto.iTipo

        End If

    Next

    Exit Sub

Erro_Carrega_TipoMeioPagto:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175323)

    End Select

    Exit Sub

End Sub

Public Function Trata_Parametros(Optional objTransfCaixa As ClassTransfCaixa) As Long
'Trata os parametros passados para a tela
'objTransfCaixa eh parametro opcional de Input

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se o obj nao eh nothing
    If Not (objTransfCaixa Is Nothing) Then

        'Traz os dados pra tela...
        lErro = Traz_TransfCaixa_Tela(objTransfCaixa)
        If lErro <> SUCESSO Then gError 101796

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 101796

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175324)

    End Select

    iAlterado = 0

    Exit Function

End Function

Private Function Traz_TransfCaixa_Tela(objTransfCaixa As ClassTransfCaixa) As Long
'Traz os dados de transfcaixa e possivelmetente do cheque para a tela
'objTransfCaixa eh parametro de input

Dim objChequeDe As New ClassChequePre
Dim objChequePara As New ClassChequePre
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Traz_TransfCaixa_Tela

    'limpa a tela
    Call Limpa_Tela_TransfCaixa

    lErro = CF_ECF("TransfCaixa_Le", objTransfCaixa)
    If lErro <> SUCESSO And lErro <> 109404 Then gError 109483
    
    If lErro = 109404 Then gError 109484

    'preenche o código
    Codigo.Text = objTransfCaixa.objMovCaixaDe.lTransferencia

    'Seleciona o Conteudo apropriado na combo de tmp DE
    For iIndice = 0 To TipoMeioPagtoDe.ListCount - 1
    
        If CF_ECF("DeTipoMovtoParaTMP", objTransfCaixa.objMovCaixaDe.iTipo) = TipoMeioPagtoDe.ItemData(iIndice) Then
            TipoMeioPagtoDe.ListIndex = iIndice
            Exit For
        End If
    
    Next
    
    'Seleciona o Conteudo apropriado na combo de tmp PARA
    For iIndice = 0 To TipoMeioPagtoPara.ListCount - 1
    
        If CF_ECF("DeTipoMovtoParaTMP", objTransfCaixa.objMovCaixaPara.iTipo) = TipoMeioPagtoPara.ItemData(iIndice) Then
            TipoMeioPagtoPara.ListIndex = iIndice
            Exit For
        End If
    
    Next

    'seleciona o frame em questao do "DE" e traz para a tela
    Select Case Codigo_Extrai(TipoMeioPagtoDe.Text)

        Case TIPOMEIOPAGTOLOJA_DINHEIRO
            Call Preenche_Frame_DinDe(objTransfCaixa)

        Case TIPOMEIOPAGTOLOJA_CHEQUE

            'preenche o sequencial
            objChequeDe.lSequencialCaixa = objTransfCaixa.objMovCaixaDe.lNumRefInterna
            objChequeDe.lCupomFiscal = objTransfCaixa.objMovCaixaDe.lCupomFiscal

            'obtem os cheques (cheque de)
            lErro = CF_ECF("Cheque_Le_SequencialCaixa", objChequeDe)
            If lErro <> SUCESSO And lErro <> 109462 Then gError 109416
            If lErro = 109462 Then gError 109417

            Call Preenche_Frame_ChqDe(objChequeDe)

        Case TIPOMEIOPAGTOLOJA_CARTAO_CREDITO, TIPOMEIOPAGTOLOJA_CARTAO_DEBITO
            Call Preenche_Frame_CrtDe(objTransfCaixa)

        Case TIPOMEIOPAGTOLOJA_VALE_TICKET, TIPOMEIOPAGTOLOJA_VALE_REFEICAO, TIPOMEIOPAGTOLOJA_VALE_PRESENTE, TIPOMEIOPAGTOLOJA_VALE_COMBUSTIVEL
            Call Preenche_Frame_TktDe(objTransfCaixa)

        Case TIPOMEIOPAGTOLOJA_OUTROS
            Call Preenche_Frame_OutDe(objTransfCaixa)

    End Select

    'seleciona o frame em questao do "PARA" e traz para a tela
    Select Case Codigo_Extrai(TipoMeioPagtoPara.Text)

        Case TIPOMEIOPAGTOLOJA_DINHEIRO
            'não faz nada

        Case TIPOMEIOPAGTOLOJA_CHEQUE

            'preenche o sequencial
            objChequePara.lSequencialCaixa = objTransfCaixa.objMovCaixaPara.lNumRefInterna

            'obtem os cheques (cheque Para)
            lErro = CF_ECF("Cheque_Le_SequencialCaixa", objChequePara)
            If lErro <> SUCESSO And lErro <> 109462 Then gError 109418
            If lErro = 109462 Then gError 109419

            Call Preenche_Frame_ChqPara(objTransfCaixa, objChequePara)

        Case TIPOMEIOPAGTOLOJA_CARTAO_CREDITO, TIPOMEIOPAGTOLOJA_CARTAO_DEBITO
            Call Preenche_Frame_CrtPara(objTransfCaixa)

        Case TIPOMEIOPAGTOLOJA_VALE_TICKET, TIPOMEIOPAGTOLOJA_VALE_REFEICAO, TIPOMEIOPAGTOLOJA_VALE_PRESENTE, TIPOMEIOPAGTOLOJA_VALE_COMBUSTIVEL
            Call Preenche_Frame_TktPara(objTransfCaixa)

        Case TIPOMEIOPAGTOLOJA_OUTROS
            Call Preenche_Frame_OutPara(objTransfCaixa)

    End Select

    Traz_TransfCaixa_Tela = SUCESSO

    Exit Function

Erro_Traz_TransfCaixa_Tela:

    Traz_TransfCaixa_Tela = gErr

    Select Case gErr

        Case 109417
            Call Rotina_ErroECF(vbOKOnly, ERRO_CHEQUEPRE_NAO_ENCONTRADO2, gErr, objChequeDe.lSequencialCaixa)

        Case 109419
            Call Rotina_ErroECF(vbOKOnly, ERRO_CHEQUEPRE_NAO_ENCONTRADO2, gErr, objChequePara.lSequencialCaixa)

        Case 109416, 109418

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175325)

    End Select

    Exit Function

End Function


'Private Sub TipoMeioPagtoDe_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim iCodigo As Integer
'
'On Error GoTo Erro_TipoMeioPagtoDe_Validate
'
'    'Se foi preenchida a Combo TipoMeioPagtoDe
'    If Len(Trim(TipoMeioPagtoDe.Text)) > 0 Then
'
'        'Se o tipo atual for diferente do selecionado
'        If TipoMeioPagtoDe.ListIndex = -1 Then
'
'            'Verifica se existe o item na List da Combo. Se existir seleciona.
'            lErro = Combo_Seleciona(TipoMeioPagtoDe, iCodigo)
'            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 101819
'
'            'Nao existe o item com o CODIGO na List da ComboBox
'            If lErro = 6730 Then gError 101821
'
'            'Nao existe o item com a STRING na List da ComboBox
'            If lErro = 6731 Then gError 101820
'
'        End If
'
'        Call TipoMeioPagto_TrataFramesDe
'
'    End If
'
'    Exit Sub
'
'Erro_TipoMeioPagtoDe_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 101819
'
'        Case 101820
'            Call Rotina_ErroECF(vbOKOnly, ERRO_NOME_TIPOMEIOPAGTO_NAO_EXISTENTE, gErr, TipoMeioPagtoDe.Text)
'
'        Case 101821
'            Call Rotina_ErroECF(vbOKOnly, ERRO_CODIGO_TIPOMEIOPAGTO_NAO_EXISTENTE, gErr, TipoMeioPagtoDe.Text)
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175326)
'
'    End Select
'
'    Exit Sub
'
'End Sub

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
            Call Carrega_AdmMeioPagto_Cartao(AdmCrtDe, Codigo_Extrai(TipoMeioPagtoDe.Text))
            FrameCrtDe.Visible = True
            FrameCrtDe.Enabled = True
            OptionManualDe.Enabled = True

        Case TIPOMEIOPAGTOLOJA_CARTAO_DEBITO
            Call Carrega_AdmMeioPagto_Cartao(AdmCrtDe, Codigo_Extrai(TipoMeioPagtoDe.Text))
            FrameCrtDe.Visible = True
            FrameCrtDe.Enabled = True
            OptionManualDe.Enabled = False
            OptionPOSDe.Value = True
        
        Case TIPOMEIOPAGTOLOJA_VALE_TICKET, TIPOMEIOPAGTOLOJA_VALE_REFEICAO, TIPOMEIOPAGTOLOJA_VALE_PRESENTE, TIPOMEIOPAGTOLOJA_VALE_COMBUSTIVEL
            FrameTktDe.Visible = True
            FrameTktDe.Enabled = True

        Case TIPOMEIOPAGTOLOJA_OUTROS
            FrameOutDe.Visible = True
            FrameOutDe.Enabled = True

    End Select

End Sub
'Private Sub TipoMeioPagtoPara_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim iCodigo As Integer
'
'On Error GoTo Erro_TipoMeioPagtoPara_Validate
'
'    'Se foi preenchida a Combo TipoMeioPagtoDe
'    If Len(Trim(TipoMeioPagtoPara.Text)) > 0 Then
'
'        'Se o tipo atual for diferente do selecionado
'        If TipoMeioPagtoPara.ListIndex = -1 Then
'
'            'Verifica se existe o item na List da Combo. Se existir seleciona.
'            lErro = Combo_Seleciona(TipoMeioPagtoPara, iCodigo)
'            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 101832
'
'            'Nao existe o item com o CODIGO na List da ComboBox
'            If lErro = 6730 Then gError 101834
'
'            'Nao existe o item com a STRING na List da ComboBox
'            If lErro = 6731 Then gError 101833
'
'        'senao
'        End If
'
'        'coloca os frames invisiveis
'        Call TipoMeioPagto_TrataFramesPara
'
'    End If
'
'    Exit Sub
'
'Erro_TipoMeioPagtoPara_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 101832
'
'        Case 101833
'            Call Rotina_ErroECF(vbOKOnly, ERRO_NOME_TIPOMEIOPAGTO_NAO_EXISTENTE, gErr, TipoMeioPagtoPara.Text)
'
'        Case 101834
'            Call Rotina_ErroECF(vbOKOnly, ERRO_CODIGO_TIPOMEIOPAGTO_NAO_EXISTENTE, gErr, TipoMeioPagtoPara.Text)
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175327)
'
'    End Select
'
'    Exit Sub
'
'End Sub

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
            Call Carrega_AdmMeioPagto_Cartao(AdmCrtPara, Codigo_Extrai(TipoMeioPagtoPara.Text))
            FrameCrtPara.Visible = True
            FrameCrtPara.Enabled = True
            OptionManualPara.Enabled = True

        Case TIPOMEIOPAGTOLOJA_CARTAO_DEBITO
            Call Carrega_AdmMeioPagto_Cartao(AdmCrtPara, Codigo_Extrai(TipoMeioPagtoPara.Text))
            FrameCrtPara.Visible = True
            FrameCrtPara.Enabled = True
            OptionManualPara.Enabled = False
            OptionPOSPara.Value = True
        
        Case TIPOMEIOPAGTOLOJA_VALE_TICKET, TIPOMEIOPAGTOLOJA_VALE_REFEICAO, TIPOMEIOPAGTOLOJA_VALE_PRESENTE, TIPOMEIOPAGTOLOJA_VALE_COMBUSTIVEL
            FrameTktPara.Visible = True
            FrameTktPara.Enabled = True

        Case TIPOMEIOPAGTOLOJA_OUTROS
            FrameOutPara.Visible = True
            FrameOutPara.Enabled = True

    End Select

End Sub

Private Sub Preenche_Frame_DinDe(objTransfCaixa As ClassTransfCaixa)

    'coloca o valor na tela
    ValorDinDe.Text = Format(objTransfCaixa.objMovCaixaDe.dValor, "Standard")

End Sub

Private Sub Preenche_Frame_ChqDe(objCheque As ClassChequePre)
'preenche o frame chq de origem
'ObjCheque eh parametro de input

Dim sCPFCGC As String
Dim bCancel As Boolean

    'se banco estiver preenchido, coloca na tela
    If objCheque.iBanco > 0 Then

        BancoChqDe.Caption = CStr(objCheque.iBanco)

    End If

    'preenche a agencia
    AgenciaChqDe.Caption = objCheque.sAgencia

    'preenche a conta corrente
    ContaChqDe.Caption = objCheque.sContaCorrente

    'preenche o numero
    If objCheque.lNumero > 0 Then

        NumeroChqDe.Caption = CStr(objCheque.lNumero)

    End If

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

    'preenche o cupom fiscal
    If objCheque.lCupomFiscal > 0 Then

        CupomFiscalChqDe.Caption = CStr(objCheque.lCupomFiscal)

    End If

    'preenche a data bom para
    DataBomParaChqDe.Caption = Format(objCheque.dtDataDeposito, "dd/mm/yyyy")

    'preenche o valor
    ValorChqDe.Caption = Format(objCheque.dValor, "standard")
    
    SeqChqDe.Text = objCheque.lSequencialCaixa

End Sub

Private Sub Preenche_Frame_CrtDe(objTransfCaixa As ClassTransfCaixa)
    
    If objTransfCaixa.objMovCaixaDe.iAdmMeioPagto <> 0 Then
        
        AdmCrtDe.Text = objTransfCaixa.objMovCaixaDe.iAdmMeioPagto
        Call AdmCrtDe_Validate(bSGECancelDummy)
    
        ParcelamentoCrtDe.Text = objTransfCaixa.objMovCaixaDe.iParcelamento
        Call ParcelamentoCrtDe_Validate(bSGECancelDummy)
    
    End If

    If CF_ECF("DeTipoMovtoParaTMP", objTransfCaixa.objMovCaixaDe.iTipo) = TIPOMEIOPAGTOLOJA_CARTAO_CREDITO Then
        
        If objTransfCaixa.objMovCaixaDe.iTipoCartao = TIPO_MANUAL Then OptionManualDe.Value = True
        
        If objTransfCaixa.objMovCaixaDe.iTipoCartao = TIPO_POS Then OptionPOSDe.Value = True
    
    End If

    'preenche o valor
    ValorCrtDe.Text = Format(objTransfCaixa.objMovCaixaDe.dValor, "standard")

End Sub
Private Sub AdmCrtDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AdmCrtDe_Validate

    ParcelamentoCrtDe.Clear
    
    'Se foi preenchida a Combo
    If Len(Trim(AdmCrtDe.Text)) > 0 Then

        'Se o cartao atual for diferente do selecionado
        If AdmCrtDe.ListIndex = -1 Then

            'Verifica se existe o item na List da Combo. Se existir seleciona.
            lErro = Combo_Seleciona(AdmCrtDe, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 101845

            'Nao existe o item com o CODIGO na List da ComboBox
            If lErro = 6730 Then gError 101847

            'Nao existe o item com a STRING na List da ComboBox
            If lErro = 6731 Then gError 101846

        End If

        Call Carrega_Parcelamento(ParcelamentoCrtDe, Codigo_Extrai(AdmCrtDe.Text))

    End If

    Exit Sub

Erro_AdmCrtDe_Validate:

    Cancel = True

    Select Case gErr

        Case 101845

        Case 101846
            Call Rotina_ErroECF(vbOKOnly, ERRO_NOME_CARTAO_NAO_EXISTENTE, gErr, AdmCrtDe.Text)

        Case 101847
            Call Rotina_ErroECF(vbOKOnly, ERRO_CODIGO_CARTAO_NAO_EXISTENTE, gErr, AdmCrtDe.Text)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175328)

    End Select

    Exit Sub

End Sub

Private Sub Carrega_Parcelamento(objComboParcelamento As ComboBox, iAdmMeioPagto As Integer)

Dim objAdmMeioPagto As ClassAdmMeioPagto
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
Dim bAchou As Boolean

On Error GoTo Erro_Carrega_Parcelamento

    bAchou = False

    'busca na coleção global de admmeiopagto o admmeiopagto passado por parâmetro
    For Each objAdmMeioPagto In gcolAdmMeioPagto

        'se encontrou
        If objAdmMeioPagto.iCodigo = iAdmMeioPagto Then

            bAchou = True
            Exit For

        End If

    Next

    'se não encontrou-> erro
    If bAchou = False Then gError 109421

    'no admmeiopagto encontrado, varre sua coleção de admmeiopagtocondpagto e preenche a combo de parcelamento recebida por parâmetro
    For Each objAdmMeioPagtoCondPagto In objAdmMeioPagto.colCondPagtoLoja

        objComboParcelamento.AddItem (objAdmMeioPagtoCondPagto.iParcelamento & SEPARADOR & objAdmMeioPagtoCondPagto.sNomeParcelamento)
        objComboParcelamento.ItemData(objComboParcelamento.NewIndex) = objAdmMeioPagtoCondPagto.iParcelamento

    Next

    Exit Sub

Erro_Carrega_Parcelamento:

    Select Case gErr

        Case 109421
            Call Rotina_ErroECF(vbOKOnly, ERRO_ADMMEIOPAGTO_NAO_CADASTRADO, gErr, iAdmMeioPagto)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175329)

    End Select

    Exit Sub

End Sub


Private Sub Preenche_Frame_TktDe(objTransfCaixa As ClassTransfCaixa)

    If objTransfCaixa.objMovCaixaDe.iAdmMeioPagto <> 0 Then
        
        AdmTktDe.Text = objTransfCaixa.objMovCaixaDe.iAdmMeioPagto
        Call AdmTktDe_Validate(bSGECancelDummy)
    
    End If
    
    'coloca o valor na tela
    ValorTktDe.Text = Format(objTransfCaixa.objMovCaixaDe.dValor, "standard")

End Sub
Private Sub AdmTktDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_AdmTktDe_Validate
    
    'Se foi preenchida a Combo
    If Len(Trim(AdmTktDe.Text)) > 0 Then

        'Se o cartao atual for diferente do selecionado
        If AdmTktDe.ListIndex = -1 Then

            'Verifica se existe o item na List da Combo. Se existir seleciona.
            lErro = Combo_Seleciona(AdmTktDe, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 101860

            'Nao existe o item com o CODIGO na List da ComboBox
            If lErro = 6730 Then gError 101862

            'Nao existe o item com a STRING na List da ComboBox
            If lErro = 6731 Then gError 101861

        End If

    End If

    Exit Sub

Erro_AdmTktDe_Validate:

    Cancel = True

    Select Case gErr

        Case 101860

        Case 101861
            Call Rotina_ErroECF(vbOKOnly, ERRO_NOME_VALETICKET_NAO_EXISTENTE, gErr, AdmTktDe.Text)

        Case 101862
            Call Rotina_ErroECF(vbOKOnly, ERRO_CODIGO_VALETICKET_NAO_EXISTENTE, gErr, AdmTktDe.Text)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175330)

    End Select

    Exit Sub

End Sub

Private Sub Preenche_Frame_OutDe(objTransfCaixa As ClassTransfCaixa)
    
    If objTransfCaixa.objMovCaixaDe.iAdmMeioPagto <> 0 Then
        
        AdmOutDe.Text = objTransfCaixa.objMovCaixaDe.iAdmMeioPagto
        Call AdmOutDe_Validate(bSGECancelDummy)
    
    End If

    'coloca o valor na tela
    ValorOutDe.Text = Format(objTransfCaixa.objMovCaixaDe.dValor, "standard")
    
End Sub

Private Sub AdmOutDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_AdmOutDe_Validate

    'Se foi preenchida a Combo
    If Len(Trim(AdmOutDe.Text)) > 0 Then

        'Se o cartao atual for diferente do selecionado
        If AdmOutDe.ListIndex = -1 Then

            'Verifica se existe o item na List da Combo. Se existir seleciona.
            lErro = Combo_Seleciona(AdmOutDe, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 109422

            'Nao existe o item com o CODIGO na List da ComboBox
            If lErro = 6730 Then gError 109423

            'Nao existe o item com a STRING na List da ComboBox
            If lErro = 6731 Then gError 109424

        End If

    End If

    Exit Sub

Erro_AdmOutDe_Validate:

    Cancel = True

    Select Case gErr

        Case 109422

        Case 109424
            Call Rotina_ErroECF(vbOKOnly, ERRO_NOME_OUTROS_NAO_EXISTENTE, gErr, AdmOutDe.Text)

        Case 109423
            Call Rotina_ErroECF(vbOKOnly, ERRO_CODIGO_OUTROS_NAO_EXISTENTE, gErr, AdmOutDe.Text)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175331)

    End Select

    Exit Sub

End Sub

Private Sub Preenche_Frame_ChqPara(objTransfCaixa As ClassTransfCaixa, objCheque As ClassChequePre)

    'se banco estiver preenchido, coloca na tela
    If objCheque.iBanco > 0 Then

        BancoChqPara.Text = CStr(objCheque.iBanco)

    End If

    'preenche a agencia
    AgenciaChqPara.Text = (objCheque.sAgencia)


    'preenche a conta corrente
    ContaChqPara.Text = objCheque.sContaCorrente

    'preenche o numero
    NumeroChqPara.Text = CStr(objCheque.lNumero)

'    'preenche o carne
'    CarnePara.Text = objCheque.sCarne

    'se o cupom fiscal estiver preenchido
    If objCheque.lCupomFiscal > 0 Then

        CupomFiscalChqPara.Text = CStr(objCheque.lCupomFiscal)
        ECFChqPara.Text = CStr(objCheque.iECF)

    End If

    'preenche o cpf/cgc do cliente
    ClienteChqPara.PromptInclude = False
    ClienteChqPara.Text = objCheque.sCPFCGC
    ClienteChqPara.PromptInclude = True
    
    'coloca a mascara
    Call ClienteChqPara_Validate(bSGECancelDummy)

    'preenche a data bom para
    DataBomParaChqPara.Text = Format(objCheque.dtDataDeposito, "dd/mm/yy")

End Sub

Private Sub ClienteChqPara_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ClienteChqPara_Validate

    'Se CGC/CPF não foi preenchido -- Exit Sub
    If Len(Trim(ClienteChqPara.ClipText)) = 0 Then Exit Sub

    Select Case Len(Trim(ClienteChqPara.ClipText))

        Case STRING_CPF 'CPF

            'Critica Cpf
            lErro = Cpf_Critica(ClienteChqPara.Text)
            If lErro <> SUCESSO Then gError 101842

            'Formata e coloca na Tela
            ClienteChqPara.Format = "000\.000\.000-00; ; ; "
            ClienteChqPara.Text = ClienteChqPara.Text

        Case STRING_CGC 'CGC

            'Critica CGC
            lErro = Cgc_Critica(ClienteChqPara.Text)
            If lErro <> SUCESSO Then gError 101843

            'Formata e Coloca na Tela
            ClienteChqPara.Format = "00\.000\.000\/0000-00; ; ; "
            ClienteChqPara.Text = ClienteChqPara.Text

        Case Else

            gError 101844

    End Select

    Exit Sub

Erro_ClienteChqPara_Validate:

    Cancel = True

    Select Case gErr

        Case 101842, 101843

        Case 101844
            Call Rotina_ErroECF(vbOKOnly, ERRO_TAMANHO_CGC_CPF, gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175332)

    End Select

    Exit Sub

End Sub

Private Sub Preenche_Frame_CrtPara(objTransfCaixa As ClassTransfCaixa)

    If objTransfCaixa.objMovCaixaPara.iAdmMeioPagto <> 0 Then
        
        AdmCrtPara.Text = objTransfCaixa.objMovCaixaPara.iAdmMeioPagto
        Call AdmCrtPara_Validate(bSGECancelDummy)
    
        ParcelamentoCrtPara.Text = objTransfCaixa.objMovCaixaPara.iParcelamento
        Call ParcelamentoCrtPara_Validate(bSGECancelDummy)
    
    End If
    
    If CF_ECF("DeTipoMovtoParaTMP", objTransfCaixa.objMovCaixaPara.iTipo) = TIPOMEIOPAGTOLOJA_CARTAO_CREDITO Then
        
        If objTransfCaixa.objMovCaixaPara.iTipoCartao = TIPO_MANUAL Then OptionManualPara.Value = True
        
        If objTransfCaixa.objMovCaixaPara.iTipoCartao = TIPO_POS Then OptionPOSPara.Value = True
    
    End If

End Sub

Private Sub Preenche_Frame_TktPara(objTransfCaixa As ClassTransfCaixa)

    If objTransfCaixa.objMovCaixaPara.iAdmMeioPagto <> 0 Then
    
        AdmTktPara.Text = objTransfCaixa.objMovCaixaPara.iAdmMeioPagto
        Call AdmTktPara_Validate(bSGECancelDummy)
        
    End If

End Sub

Private Sub Preenche_Frame_OutPara(objTransfCaixa As ClassTransfCaixa)

    If objTransfCaixa.objMovCaixaPara.iAdmMeioPagto <> 0 Then
    
        AdmOutPara.Text = objTransfCaixa.objMovCaixaPara.iAdmMeioPagto
        Call AdmOutPara_Validate(bSGECancelDummy)
    
    End If

End Sub

Private Sub AdmCrtPara_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AdmCrtPara_Validate

    ParcelamentoCrtPara.Clear
    
    'Se foi preenchida a Combo
    If Len(Trim(AdmCrtPara.Text)) > 0 Then

        'Se o cartao atual for diferente do selecionado
        If AdmCrtPara.ListIndex = -1 Then

            'Verifica se existe o item na List da Combo. Se existir seleciona.
            lErro = Combo_Seleciona(AdmCrtPara, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 101849

            'Nao existe o item com o CODIGO na List da ComboBox
            If lErro = 6730 Then gError 101851

            'Nao existe o item com a STRING na List da ComboBox
            If lErro = 6731 Then gError 101850


        End If

        Call Carrega_Parcelamento(ParcelamentoCrtPara, Codigo_Extrai(AdmCrtPara.Text))

    End If

    Exit Sub

Erro_AdmCrtPara_Validate:

    Cancel = True

    Select Case gErr

        Case 101849

        Case 101850
            Call Rotina_ErroECF(vbOKOnly, ERRO_NOME_CARTAO_NAO_EXISTENTE, gErr, AdmCrtPara.Text)

        Case 101851
            Call Rotina_ErroECF(vbOKOnly, ERRO_CODIGO_CARTAO_NAO_EXISTENTE, gErr, AdmCrtPara.Text)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175333)

    End Select

    Exit Sub

End Sub

Private Sub ECFChqPara_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ECFChqPara_GotFocus()
        Call MaskEdBox_TrataGotFocus(ECFChqPara, iAlterado)
End Sub

Private Sub LabelECF_Click()

Dim objCupom As New ClassCupomFiscal
Dim lErro As Long

    'chama o browser cupomFiscalLista passando objcupom
    Call Chama_TelaECF_Modal("CupomFiscalLista", objCupom)

    'se objCupom tiver preenchido, chama pra tela..
    If giRetornoTela = vbOK Then
    
        ECFChqPara.Text = CStr(objCupom.iECF)
        CupomFiscalChqPara.Text = CStr(objCupom.lNumero)

    End If

End Sub

Private Sub LabelCupomFiscalChqPara_Click()

Dim objCupom As New ClassCupomFiscal
Dim lErro As Long

On Error GoTo Erro_LabelCupomFiscalChqPara_Click

    'chama o browser cupomFiscalLista passando objcupom
    Call Chama_TelaECF_Modal("CupomFiscalLista", objCupom)

    'se objCupom tiver preenchido, chama pra tela..
    If giRetornoTela = vbOK Then
    
        ECFChqPara.Text = CStr(objCupom.iECF)
        CupomFiscalChqPara.Text = CStr(objCupom.lNumero)

    End If

    Exit Sub

Erro_LabelCupomFiscalChqPara_Click:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175334)

    End Select

    Exit Sub

End Sub

Private Sub LabelSeqChqDe_Click()
    Call BotaoChequeDe_Click
End Sub

Private Sub ParcelamentoCrtPara_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_ParcelamentoCrtPara_Validate

    'Se foi preenchida a Combo
    If Len(Trim(ParcelamentoCrtPara.Text)) > 0 Then

        'Se o parcelamento atual for diferente do selecionado
        If ParcelamentoCrtPara.ListIndex = -1 Then

            'Verifica se existe o item na List da Combo. Se existir seleciona.
            lErro = Combo_Seleciona(ParcelamentoCrtPara, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 101856

            'Nao existe o item com o CODIGO na List da ComboBox
            If lErro = 6730 Then gError 101858

            'Nao existe o item com a STRING na List da ComboBox
            If lErro = 6731 Then gError 101857

        End If

    End If

    Exit Sub

Erro_ParcelamentoCrtPara_Validate:

    Cancel = True

    Select Case gErr

        Case 101856

        Case 101857
            Call Rotina_ErroECF(vbOKOnly, ERRO_NOME_PARCELAMENTO_NAO_EXISTENTE, gErr, ParcelamentoCrtPara.Text)

        Case 101858
            Call Rotina_ErroECF(vbOKOnly, ERRO_CODIGO_PARCELAMENTO_NAO_EXISTENTE, gErr, ParcelamentoCrtPara.Text)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175335)

    End Select

    Exit Sub

End Sub
Private Sub AdmTktPara_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_AdmTktPara_Validate

    'Se foi preenchida a Combo
    If Len(Trim(AdmTktPara.Text)) > 0 Then

        'Se o cartao atual for diferente do selecionado
        If AdmTktPara.ListIndex = -1 Then

            'Verifica se existe o item na List da Combo. Se existir seleciona.
            lErro = Combo_Seleciona(AdmTktPara, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 101961

            'Nao existe o item com o CODIGO na List da ComboBox
            If lErro = 6730 Then gError 101962

            'Nao existe o item com a STRING na List da ComboBox
            If lErro = 6731 Then gError 101963

        End If

    End If

    Exit Sub

Erro_AdmTktPara_Validate:

    Cancel = True

    Select Case gErr

        Case 101961

        Case 101962
            Call Rotina_ErroECF(vbOKOnly, ERRO_CODIGO_VALETICKET_NAO_EXISTENTE, gErr, AdmTktPara.Text)

        Case 101963
            Call Rotina_ErroECF(vbOKOnly, ERRO_NOME_VALETICKET_NAO_EXISTENTE, gErr, AdmTktPara.Text)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175336)

    End Select

    Exit Sub

End Sub

Private Sub AdmOutPara_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_AdmOutPara_Validate

    'Se foi preenchida a Combo
    If Len(Trim(AdmOutPara.Text)) > 0 Then

        'Se o cartao atual for diferente do selecionado
        If AdmOutPara.ListIndex = -1 Then

            'Verifica se existe o item na List da Combo. Se existir seleciona.
            lErro = Combo_Seleciona(AdmOutPara, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 109425

            'Nao existe o item com o CODIGO na List da ComboBox
            If lErro = 6730 Then gError 109426

            'Nao existe o item com a STRING na List da ComboBox
            If lErro = 6731 Then gError 109427

        End If

    End If

    Exit Sub

Erro_AdmOutPara_Validate:

    Cancel = True

    Select Case gErr

        Case 109425

        Case 109426
            Call Rotina_ErroECF(vbOKOnly, ERRO_CODIGO_OUTROS_NAO_EXISTENTE, gErr, AdmOutPara.Text)

        Case 109427
            Call Rotina_ErroECF(vbOKOnly, ERRO_NOME_OUTROS_NAO_EXISTENTE, gErr, AdmOutPara.Text)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175337)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 207996

    'chama a funcao q ira realizar efetivamente a gravacao
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 101871

    'limpa a tela
    Call Limpa_Tela_TransfCaixa

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 101871, 207996

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175338)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objTransfCaixa As New ClassTransfCaixa
Dim objChequeDe As New ClassChequePre
Dim objChequePara As New ClassChequePre

On Error GoTo Erro_Gravar_Registro

    'se o codigo nao estiver preenchido, erro
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 101880

    'se o tipomeiopagtode nao estiver preenchido, erro
    If Len(Trim(TipoMeioPagtoDe.Text)) = 0 Then gError 101881

    'se o tipomeiopagtopara nao estiver preenchido, erro
    If Len(Trim(TipoMeioPagtoPara.Text)) = 0 Then gError 101882

    'verifica se os campos obrigatorios do frame DE foram preenchidos
    lErro = Checa_Frames_De()
    If lErro <> SUCESSO Then gError 101887

    'verifica se os campos obrigatorios do frame PARA foram preenchidos
    lErro = Checa_Frames_Para()
    If lErro <> SUCESSO Then gError 101888

    'move os dados da tela para a memoria
    lErro = Move_Tela_Memoria(objTransfCaixa, objChequeDe, objChequePara)
    If lErro <> SUCESSO Then gError 109458

    'grava os dados
    lErro = TransfCaixa_Grava(objTransfCaixa, objChequeDe, objChequePara)
    If lErro <> SUCESSO Then gError 101889

    'limpa a tela
    Call Limpa_Tela_TransfCaixa

    iAlterado = 0

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 101880
            Call Rotina_ErroECF(vbOKOnly, ERRO_TRANSFCAIXA_NAO_PREENCHIDO, gErr)

        Case 101881
            Call Rotina_ErroECF(vbOKOnly, ERRO_TIPOMEIOPAGTODE_NAO_PREENCHIDO, gErr)

        Case 101882
            Call Rotina_ErroECF(vbOKOnly, ERRO_TIPOMEIOPAGTOPARA_NAO_PREENCHIDO, gErr)

        Case 101887 To 101890

        Case 109458

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175339)

    End Select

    Exit Function

End Function

Private Function Checa_Frames_De() As Long
'verifica se os valores obrigatorios do frame DE estao preenchidos

Dim lErro As Long

On Error GoTo Erro_Checa_Frames_De

    'dependendo do que está selecionado em tipomeiopagtode
    Select Case TipoMeioPagtoDe.ItemData(TipoMeioPagtoDe.ListIndex)

        Case TIPOMEIOPAGTOLOJA_DINHEIRO
            lErro = Checa_Frame_DinDe()
            If lErro <> SUCESSO Then gError 109428

        Case TIPOMEIOPAGTOLOJA_CHEQUE
            lErro = Checa_Frame_ChqDe()
            If lErro <> SUCESSO Then gError 109429

        Case TIPOMEIOPAGTOLOJA_CARTAO_CREDITO, TIPOMEIOPAGTOLOJA_CARTAO_DEBITO
            lErro = Checa_Frame_CrtDe()
            If lErro <> SUCESSO Then gError 109430

        Case TIPOMEIOPAGTOLOJA_VALE_TICKET, TIPOMEIOPAGTOLOJA_VALE_REFEICAO, TIPOMEIOPAGTOLOJA_VALE_PRESENTE, TIPOMEIOPAGTOLOJA_VALE_COMBUSTIVEL
            lErro = Checa_Frame_TktDe()
            If lErro <> SUCESSO Then gError 109431

        Case TIPOMEIOPAGTOLOJA_OUTROS
            lErro = Checa_Frame_OutDe()
            If lErro <> SUCESSO Then gError 109432

    End Select

    Checa_Frames_De = SUCESSO

    Exit Function

Erro_Checa_Frames_De:

    Checa_Frames_De = gErr

    Select Case gErr

        Case 109428 To 109432

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175340)

    End Select

    Exit Function

End Function

Private Function Checa_Frame_DinDe() As Long

On Error GoTo Erro_Checa_Frame_DinDe

    If Len(Trim(ValorDinDe.ClipText)) = 0 Then gError 101883

    Checa_Frame_DinDe = SUCESSO

    Exit Function

Erro_Checa_Frame_DinDe:

    Checa_Frame_DinDe = gErr

    Select Case gErr

        Case 101883
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_NAO_PREENCHIDO_ORIGEM, gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175341)

    End Select

    Exit Function

End Function

Private Function Checa_Frame_ChqDe() As Long

On Error GoTo Erro_Checa_Frame_ChqDe

    'se o seqüencial não estiver preenchido-> erro
    If Len(Trim(SeqChqDe.Text)) = 0 Then gError 126006

    Checa_Frame_ChqDe = SUCESSO

    Exit Function

Erro_Checa_Frame_ChqDe:

    Checa_Frame_ChqDe = gErr

    Select Case gErr

        Case 126006
            Call Rotina_Erro(vbOKOnly, "ERRO_CHEQUEDE_SEQUENCIAL_NAO_INFORMADO", gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175342)

    End Select

    Exit Function

End Function

Private Function Checa_Frame_CrtDe() As Long

On Error GoTo Erro_Checa_Frame_CrtDe

    '... se valor nao estiver preenchido
    If Len(Trim(ValorCrtDe.ClipText)) = 0 Then gError 101892

    '... se o cartão estiver preenchido e o parcelamento, não-> erro
    If Len(Trim(AdmCrtDe.Text)) > 0 And Len(Trim(ParcelamentoCrtDe.Text)) = 0 Then gError 109434

    Checa_Frame_CrtDe = SUCESSO

    Exit Function

Erro_Checa_Frame_CrtDe:

    Checa_Frame_CrtDe = gErr

    Select Case gErr

        Case 101892
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_NAO_PREENCHIDO_ORIGEM, gErr)

        Case 109434
            Call Rotina_ErroECF(vbOKOnly, ERRO_PARCELAMENTO_NAO_SELECIONADO, gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175343)

    End Select

    Exit Function

End Function

Private Function Checa_Frame_TktDe() As Long

On Error GoTo Erro_Checa_Frame_TktDe

    'se o valor não estiver preenchido
    If Len(Trim(ValorTktDe.ClipText)) = 0 Then gError 101894

    Checa_Frame_TktDe = SUCESSO

    Exit Function

Erro_Checa_Frame_TktDe:

    Checa_Frame_TktDe = gErr

    Select Case gErr

        Case 101894
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_NAO_PREENCHIDO_ORIGEM, gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175344)

    End Select

    Exit Function

End Function

Private Function Checa_Frame_OutDe() As Long

On Error GoTo Erro_Checa_Frame_OutDe

    '... se valor nao estiver preenchido
    If Len(Trim(ValorOutDe.ClipText)) = 0 Then gError 101895

    Checa_Frame_OutDe = SUCESSO

    Exit Function

Erro_Checa_Frame_OutDe:

    Checa_Frame_OutDe = gErr

    Select Case gErr

        Case 101895
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_NAO_PREENCHIDO1, gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175345)

    End Select

    Exit Function

End Function

Private Function Checa_Frames_Para() As Long
'verifica se os valores obrigatorios do frame DE estao preenchidos

Dim lErro As Long

On Error GoTo Erro_Checa_Frames_Para

    'dependendo do que está selecionado em tipomeiopagtode
    Select Case TipoMeioPagtoPara.ItemData(TipoMeioPagtoPara.ListIndex)

        Case TIPOMEIOPAGTOLOJA_DINHEIRO, TIPOMEIOPAGTOLOJA_VALE_TICKET, TIPOMEIOPAGTOLOJA_VALE_REFEICAO, TIPOMEIOPAGTOLOJA_VALE_PRESENTE, TIPOMEIOPAGTOLOJA_VALE_COMBUSTIVEL, TIPOMEIOPAGTOLOJA_OUTROS
            'não tem o que checar

        Case TIPOMEIOPAGTOLOJA_CHEQUE
            lErro = Checa_Frame_ChqPara()
            If lErro <> SUCESSO Then gError 109437

        Case TIPOMEIOPAGTOLOJA_CARTAO_CREDITO, TIPOMEIOPAGTOLOJA_CARTAO_DEBITO
            lErro = Checa_Frame_CrtPara()
            If lErro <> SUCESSO Then gError 109438

    End Select

    Checa_Frames_Para = SUCESSO

    Exit Function

Erro_Checa_Frames_Para:

    Checa_Frames_Para = gErr

    Select Case gErr

        Case 109436 To 109440

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175346)

    End Select

    Exit Function

End Function

Private Function Checa_Frame_ChqPara() As Long

Dim objVenda As ClassVenda
Dim bAchou As Boolean

On Error GoTo Erro_Checa_Frame_ChqPara

    '... se a data bom para nao estiver preenchida
    If Len(Trim(DataBomParaChqPara.ClipText)) = 0 Then gError 109442

    'se o tipo (de) não for um cheque...
    If CF_ECF("DeTipoMovtoParaTMP", Codigo_Extrai(TipoMeioPagtoDe.Text)) <> TIPOMEIOPAGTOLOJA_CHEQUE Then
    
'        ou cupom fiscal ou carne devem estar preenchidos
'        If Len(Trim(CarnePara.ClipText)) = 0 And Len(Trim(CupomFiscalChqPara.ClipText)) = 0 Then gError 109443
        
        'se o cupom fiscal não estiver preenchido-> erro
        If Len(Trim(CupomFiscalChqPara.ClipText)) = 0 Then gError 109478
        
        bAchou = False
        
        'verifica se o cupom informado existe na gcolvendas
        For Each objVenda In gcolVendas
            
            'se for um cupom fiscal
            If objVenda.iTipo = OPTION_CF Then
                
                'se encontrou um com o número que está na tela
                If objVenda.objCupomFiscal.lNumero = StrParaLong(CupomFiscalChqPara.Text) And objVenda.objCupomFiscal.iECF = StrParaInt(ECFChqPara.Text) Then
                    
                    bAchou = True
                    Exit For
                
                End If
            
            End If
        
        Next
        
        'se não achou
        If bAchou = False Then gError 109481
    
    End If

'    'só o cupom fiscal ou (exclusivo) carne devem estar estar preenchidos
'    If Len(Trim(CarnePara.ClipText)) > 0 And Len(Trim(CupomFiscalChqPara.ClipText)) > 0 Then gError 109444

    '... se o numero, conta,banco ou agencia estiverem preenchidos-> todos os complementares devem estar preenchidos
    If (Len(Trim(BancoChqPara.ClipText)) > 0 Or Len(Trim(AgenciaChqPara.ClipText)) > 0 Or Len(Trim(NumeroChqPara.ClipText)) > 0 Or Len(Trim(ContaChqPara.ClipText)) > 0) And _
       (Len(Trim(BancoChqPara.ClipText)) = 0 Or Len(Trim(AgenciaChqPara.ClipText)) = 0 Or Len(Trim(NumeroChqPara.ClipText)) = 0 Or Len(Trim(ContaChqPara.ClipText)) = 0) Then gError 109445
       

    Checa_Frame_ChqPara = SUCESSO

    Exit Function

Erro_Checa_Frame_ChqPara:

    Checa_Frame_ChqPara = gErr

    Select Case gErr

        Case 109442
            Call Rotina_ErroECF(vbOKOnly, ERRO_DATABOMPARA_NAO_PREENCHIDA, gErr)

'        Case 109443
'            Call Rotina_ErroECF(vbOKOnly, ERRO_CARNE_CUPOMFISCAL_NAO_PREENCHIDOS, gErr)

        Case 109478
            Call Rotina_ErroECF(vbOKOnly, ERRO_CUPOMFISCAL_NAO_PREENCHIDO, gErr)

'        Case 109444
'            Call Rotina_ErroECF(vbOKOnly, ERRO_CARNE_CUPOMFISCAL_PREENCHIDOS, gErr)

        Case 109445
            Call Rotina_ErroECF(vbOKOnly, ERRO_GRUPOCHQDE_NAO_PREENCHIDO, gErr)
            
        Case 109481
            Call Rotina_ErroECF(vbOKOnly, ERRO_CUPOMFISCAL_NAO_ENCONTRADO, gErr, CupomFiscalChqPara.Text, ECFChqPara.Text)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175347)

    End Select

    Exit Function

End Function

Private Function Checa_Frame_CrtPara() As Long

On Error GoTo Erro_Checa_Frame_CrtPara

    '... se o cartão estiver preenchido e o parcelamento, não -> erro
    If Len(Trim(AdmCrtPara.Text)) > 0 And Len(Trim(ParcelamentoCrtPara.Text)) = 0 Then gError 109447

    Checa_Frame_CrtPara = SUCESSO

    Exit Function

Erro_Checa_Frame_CrtPara:

    Checa_Frame_CrtPara = gErr

    Select Case gErr

        Case 109447
            Call Rotina_ErroECF(vbOKOnly, ERRO_PARCELAMENTO_NAO_SELECIONADO, gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175348)

    End Select

    Exit Function

End Function

Public Function Move_Tela_Memoria(objTransfCaixa As ClassTransfCaixa, objChequeDe As ClassChequePre, objChequePara As ClassChequePre) As Long
'objTransfCaixa eh parametro de input
'objCheque eh parametro de input

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    'mover o parcelamento default
    objTransfCaixa.objMovCaixaDe.iParcelamento = PARCELAMENTO_AVISTA
    objTransfCaixa.objMovCaixaPara.iParcelamento = PARCELAMENTO_AVISTA

    'move os dados do frame De
    lErro = Move_Tela_MemoriaDe(objTransfCaixa, objChequeDe)
    If lErro <> SUCESSO Then gError 109456

    'move os dados do frame Para
    lErro = Move_Tela_MemoriaPara(objTransfCaixa, objChequePara)
    If lErro <> SUCESSO Then gError 109457

    Call Move_Tela_Memoria1(objTransfCaixa)
    
    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 109456, 109457

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175349)

    End Select

End Function

Private Function Move_Tela_MemoriaDe(objTransfCaixa As ClassTransfCaixa, objChequeDe As ClassChequePre) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_MemoriaDe

    'seleciona o tmp DE
    Select Case Codigo_Extrai(TipoMeioPagtoDe.Text)

        Case TIPOMEIOPAGTOLOJA_DINHEIRO

            'mover o tipo de movimento
            objTransfCaixa.objMovCaixaDe.iTipo = MOVIMENTOCAIXA_SAIDA_TRANSF_DINHEIRO

            'mover o valor
            objTransfCaixa.objMovCaixaDe.dValor = StrParaDbl(ValorDinDe.Text)

            'mover o admmeiopagto
            objTransfCaixa.objMovCaixaDe.iAdmMeioPagto = MEIO_PAGAMENTO_DINHEIRO
            
        Case TIPOMEIOPAGTOLOJA_CHEQUE

            'mover o tipo de movimento
            objTransfCaixa.objMovCaixaDe.iTipo = MOVIMENTOCAIXA_SAIDA_TRANSF_CHEQUE

            'mover o valor
            objTransfCaixa.objMovCaixaDe.dValor = StrParaDbl(ValorChqDe.Caption)

            'mover o admmeiopagto
            objTransfCaixa.objMovCaixaDe.iAdmMeioPagto = MEIO_PAGAMENTO_CHEQUE

            'preenche os atributos chaves de cheque
            Call Move_Dados_ChequeDe(objChequeDe)

            'busca o cheque na coleção global
            lErro = CF_ECF("Cheque_Le_SequencialCaixa", objChequeDe)
            If lErro <> SUCESSO And lErro <> 109462 Then gError 109399

            'se não encontrar->Erro
            If lErro = 109462 Then gError 109400

            'se o cheque já tiver sangrado ==> erro
            If objChequeDe.lNumMovtoSangria <> 0 Then gError 117564

            'guarda em numrefInterna o sequencialcaixa do cheque
            objTransfCaixa.objMovCaixaDe.lNumRefInterna = objChequeDe.lSequencialCaixa

        Case TIPOMEIOPAGTOLOJA_CARTAO_CREDITO

            'mover o tipo de movimento
            objTransfCaixa.objMovCaixaDe.iTipo = MOVIMENTOCAIXA_SAIDA_TRANSF_CARTAO_CREDITO

            'mover o valor
            objTransfCaixa.objMovCaixaDe.dValor = StrParaDbl(ValorCrtDe.Text)
           
           'move o codigo
            If Len(Trim(AdmCrtDe.Text)) <> 0 Then objTransfCaixa.objMovCaixaDe.iAdmMeioPagto = Codigo_Extrai(AdmCrtDe.Text)

            'move o parcelamento
            If Len(Trim(ParcelamentoCrtDe.Text)) <> 0 Then objTransfCaixa.objMovCaixaDe.iParcelamento = Codigo_Extrai(ParcelamentoCrtDe.Text)
            
            'move o tipo de terminal
            If OptionManualDe.Value = True Then objTransfCaixa.objMovCaixaDe.iTipoCartao = TIPO_MANUAL
            
            'move o tipo de terminal
            If OptionPOSDe.Value = True Then objTransfCaixa.objMovCaixaDe.iTipoCartao = TIPO_POS
            

        Case TIPOMEIOPAGTOLOJA_CARTAO_DEBITO

            'mover o tipo de movimento
            objTransfCaixa.objMovCaixaDe.iTipo = MOVIMENTOCAIXA_SAIDA_TRANSF_CARTAO_DEBITO

            'mover o valor
            objTransfCaixa.objMovCaixaDe.dValor = StrParaDbl(ValorCrtDe.Text)

            'move o codigo
            If Len(Trim(AdmCrtDe.Text)) <> 0 Then objTransfCaixa.objMovCaixaDe.iAdmMeioPagto = Codigo_Extrai(AdmCrtDe.Text)

            'move o parcelamento
            If Len(Trim(ParcelamentoCrtDe.Text)) <> 0 Then objTransfCaixa.objMovCaixaDe.iParcelamento = Codigo_Extrai(ParcelamentoCrtDe.Text)

            'move o tipo de terminal
            objTransfCaixa.objMovCaixaDe.iTipoCartao = TIPO_POS
        
        Case TIPOMEIOPAGTOLOJA_VALE_TICKET, TIPOMEIOPAGTOLOJA_VALE_REFEICAO, TIPOMEIOPAGTOLOJA_VALE_PRESENTE, TIPOMEIOPAGTOLOJA_VALE_COMBUSTIVEL

            'mover o tipo de movimento
            objTransfCaixa.objMovCaixaDe.iTipo = MOVIMENTOCAIXA_SAIDA_TRANSF_VALETICKET

            'mover o valor
            objTransfCaixa.objMovCaixaDe.dValor = StrParaDbl(ValorTktDe.Text)

            'move o codigo
            If Len(Trim(AdmTktDe.Text)) <> 0 Then objTransfCaixa.objMovCaixaDe.iAdmMeioPagto = Codigo_Extrai(AdmTktDe.Text)

        Case TIPOMEIOPAGTOLOJA_OUTROS

            'mover o tipo de movimento
            objTransfCaixa.objMovCaixaDe.iTipo = MOVIMENTOCAIXA_SAIDA_TRANSF_OUTROS

            'mover o valor
            objTransfCaixa.objMovCaixaDe.dValor = StrParaDbl(ValorOutDe.Text)

            'move o codigo
            If Len(Trim(AdmOutDe.Text)) <> 0 Then objTransfCaixa.objMovCaixaDe.iAdmMeioPagto = Codigo_Extrai(AdmOutDe.Text)

    End Select

    Move_Tela_MemoriaDe = SUCESSO

    Exit Function

Erro_Move_Tela_MemoriaDe:

    Move_Tela_MemoriaDe = gErr

    Select Case gErr

        Case 109399

        Case 109400
            Call Rotina_ErroECF(vbOKOnly, ERRO_CHEQUEPRE_NAO_ENCONTRADO2, gErr, objChequeDe.lSequencialCaixa)

        Case 117564
            Call Rotina_ErroECF(vbOKOnly, ERRO_CHEQUE_SANGRADO, gErr, objChequeDe.lSequencialCaixa)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175350)

    End Select

    Exit Function

End Function

Private Sub Move_Dados_ChequeDe(objCheque As ClassChequePre)

    objCheque.lSequencialCaixa = StrParaLong(SeqChqDe.Text)

'   objCheque.dtDataDeposito = StrParaDate(DataBomParaChqDe.Text)
'   objCheque.iBanco = StrParaInt(BancoChqDe.Text)
'   objCheque.iFilialEmpresaLoja = giFilialEmpresa
'   objCheque.iFilialEmpresa = giFilialEmpresa
'   objCheque.sAgencia = AgenciaChqDe.Text
'   objCheque.sCPFCGC = ClienteChqDe.Text
'   objCheque.sContaCorrente = ContaChqDe.Text
'   objCheque.dValor = StrParaDbl(ValorChqDe.Text)
'   objCheque.lNumero = StrParaLong(NumeroChqDe.Text)
'   objCheque.lCupomFiscal = StrParaLong(CupomFiscalChqDe.Text)
'   objCheque.iECF = giCodECF

End Sub

Private Function Move_Tela_MemoriaPara(objTransfCaixa As ClassTransfCaixa, objChequePara As ClassChequePre) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_MemoriaPara

    'mover o valor
    objTransfCaixa.objMovCaixaPara.dValor = objTransfCaixa.objMovCaixaDe.dValor

    'seleciona o tmp PARA
    Select Case Codigo_Extrai(TipoMeioPagtoPara.Text)

        Case TIPOMEIOPAGTOLOJA_DINHEIRO

            'mover o tipo de movimento
            objTransfCaixa.objMovCaixaPara.iTipo = MOVIMENTOCAIXA_ENTRADA_TRANSF_DINHEIRO

            'mover o admmeiopagto
            objTransfCaixa.objMovCaixaPara.iAdmMeioPagto = MEIO_PAGAMENTO_DINHEIRO
            
        Case TIPOMEIOPAGTOLOJA_CHEQUE

            'mover o tipo de movimento
            objTransfCaixa.objMovCaixaPara.iTipo = MOVIMENTOCAIXA_ENTRADA_TRANSF_CHEQUE

            'mover o admmeiopagto
            objTransfCaixa.objMovCaixaPara.iAdmMeioPagto = MEIO_PAGAMENTO_CHEQUE
            
            lErro = Move_Dados_ChequePara(objTransfCaixa, objChequePara)
            
            'guarda em numrefInterna o sequencialcaixa do cheque
            objTransfCaixa.objMovCaixaPara.lNumRefInterna = objChequePara.lSequencialCaixa
            
            'guarda o numero do cupom fiscal
            objTransfCaixa.objMovCaixaPara.lCupomFiscal = objChequePara.lCupomFiscal
            
            
        Case TIPOMEIOPAGTOLOJA_CARTAO_CREDITO
            
            'mover o tipo de movimento
            objTransfCaixa.objMovCaixaPara.iTipo = MOVIMENTOCAIXA_ENTRADA_TRANSF_CARTAO_CREDITO

            'move o codigo
            If Len(Trim(AdmCrtPara.Text)) <> 0 Then objTransfCaixa.objMovCaixaPara.iAdmMeioPagto = Codigo_Extrai(AdmCrtPara.Text)

            'move o parcelamento
            If Len(Trim(ParcelamentoCrtPara.Text)) <> 0 Then objTransfCaixa.objMovCaixaPara.iParcelamento = Codigo_Extrai(ParcelamentoCrtPara.Text)

            'move o tipo de terminal
            If OptionManualPara.Value = True Then objTransfCaixa.objMovCaixaPara.iTipoCartao = TIPO_MANUAL
            
            'move o tipo de terminal
            If OptionPOSPara.Value = True Then objTransfCaixa.objMovCaixaPara.iTipoCartao = TIPO_POS
        
        Case TIPOMEIOPAGTOLOJA_CARTAO_DEBITO

            'mover o tipo de movimento
            objTransfCaixa.objMovCaixaPara.iTipo = MOVIMENTOCAIXA_ENTRADA_TRANSF_CARTAO_DEBITO

            'move o codigo
            If Len(Trim(AdmCrtPara.Text)) <> 0 Then objTransfCaixa.objMovCaixaPara.iAdmMeioPagto = Codigo_Extrai(AdmCrtPara.Text)

            'move o parcelamento
            If Len(Trim(ParcelamentoCrtPara.Text)) <> 0 Then objTransfCaixa.objMovCaixaPara.iParcelamento = Codigo_Extrai(ParcelamentoCrtPara.Text)
            
            objTransfCaixa.objMovCaixaPara.iTipoCartao = TIPO_POS

        Case TIPOMEIOPAGTOLOJA_VALE_TICKET, TIPOMEIOPAGTOLOJA_VALE_REFEICAO, TIPOMEIOPAGTOLOJA_VALE_PRESENTE, TIPOMEIOPAGTOLOJA_VALE_COMBUSTIVEL

            'mover o tipo de movimento
            objTransfCaixa.objMovCaixaPara.iTipo = MOVIMENTOCAIXA_ENTRADA_TRANSF_VALETICKET

            'move o codigo
            If Len(Trim(AdmTktPara.Text)) <> 0 Then objTransfCaixa.objMovCaixaPara.iAdmMeioPagto = Codigo_Extrai(AdmTktPara.Text)

        Case TIPOMEIOPAGTOLOJA_OUTROS

            'mover o tipo de movimento
            objTransfCaixa.objMovCaixaPara.iTipo = MOVIMENTOCAIXA_ENTRADA_TRANSF_OUTROS

            'move o codigo
            If Len(Trim(AdmOutPara.Text)) <> 0 Then objTransfCaixa.objMovCaixaPara.iAdmMeioPagto = Codigo_Extrai(AdmOutPara.Text)

    End Select

    Move_Tela_MemoriaPara = SUCESSO

    Exit Function

Erro_Move_Tela_MemoriaPara:

    Move_Tela_MemoriaPara = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175351)

    End Select

    Exit Function

End Function

Private Function Move_Dados_ChequePara(objTransfCaixa As ClassTransfCaixa, objCheque As ClassChequePre) As Long

Dim lTamanho As Long
Dim sRetorno As String
Dim lErro As Long

On Error GoTo Erro_Move_Dados_ChequePara

    objCheque.dtDataDeposito = StrParaDate(DataBomParaChqPara.Text)
    objCheque.dValor = objTransfCaixa.objMovCaixaDe.dValor
    objCheque.iBanco = StrParaInt(BancoChqPara.Text)
    objCheque.iCaixa = giCodCaixa
    objCheque.iFilialEmpresa = giFilialEmpresa
    objCheque.iFilialEmpresaLoja = giFilialEmpresa
    
    objCheque.lCupomFiscal = StrParaLong(CupomFiscalChqPara.Text)
    objCheque.iECF = StrParaInt(ECFChqPara.Text)
    
    objCheque.lNumero = StrParaLong(NumeroChqPara.Text)
    
    lTamanho = 50
    sRetorno = String(lTamanho, 0)

    Call GetPrivateProfileString(APLICACAO_CAIXA, "NumProxCheque", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    If sRetorno <> String(lTamanho, 0) Then objCheque.lSequencialCaixa = StrParaLong(sRetorno)
    
    If objCheque.lSequencialCaixa = 0 Then objCheque.lSequencialCaixa = 1
    
    'Atualiza o sequencial de arquivo
    lErro = WritePrivateProfileString(APLICACAO_CAIXA, "NumProxCheque", CStr(objCheque.lSequencialCaixa + 1), NOME_ARQUIVO_CAIXA)
    If lErro = 0 Then gError 105782
    
    objCheque.sAgencia = AgenciaChqPara.Text
    objCheque.sContaCorrente = ContaChqPara.Text
    objCheque.sCPFCGC = ClienteChqPara.ClipText
   
    If Len(Trim(BancoChqPara.ClipText)) = 0 Then
        objCheque.iNaoEspecificado = CHEQUE_NAO_ESPECIFICADO
    Else
        objCheque.iNaoEspecificado = CHEQUE_ESPECIFICADO
    End If
   
    objCheque.iLocalizacao = CHEQUEPRE_LOCALIZACAO_CAIXA
   
    Move_Dados_ChequePara = SUCESSO
   
    Exit Function
    
Erro_Move_Dados_ChequePara:

    Move_Dados_ChequePara = gErr
    
    Select Case gErr
    
        Case 105782
            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_NAO_ENCONTRADO1, gErr, APLICACAO_CAIXA, "NumProxCheque", NOME_ARQUIVO_CAIXA)
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175352)
    
    End Select
    
    Exit Function
   
End Function

Private Sub Move_Tela_Memoria1(objTransfCaixa As ClassTransfCaixa)
'objTransfCaixa eh parametro de input

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria1

    With objTransfCaixa
        
        'move a data
        .objMovCaixaDe.dtDataMovimento = Date
        .objMovCaixaPara.dtDataMovimento = Date

        'move filialempresa
        .objMovCaixaDe.iFilialEmpresa = giFilialEmpresa
        .objMovCaixaPara.iFilialEmpresa = giFilialEmpresa

        'move o cod do caixa
        .objMovCaixaDe.iCaixa = giCodCaixa
        .objMovCaixaPara.iCaixa = giCodCaixa

        'move o cod do operador
        .objMovCaixaDe.iCodOperador = giCodOperador
        .objMovCaixaPara.iCodOperador = giCodOperador

        'move a hora
        .objMovCaixaDe.dHora = CDbl(Time)
        .objMovCaixaPara.dHora = CDbl(Time)

        'move o codigo da tela
        .objMovCaixaDe.lTransferencia = StrParaLong(Codigo.Text)
        .objMovCaixaPara.lTransferencia = StrParaLong(Codigo.Text)


    End With

    Exit Sub

Erro_Move_Tela_Memoria1:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175353)

    End Select

    Exit Sub

End Sub

Public Function TransfCaixa_Grava(objTransfCaixa As ClassTransfCaixa, objChequeDe As ClassChequePre, objChequePara As ClassChequePre) As Long

Dim lErro As Long
Dim lSequencial As Long
Dim objTransfCaixaAux As New ClassTransfCaixa
Dim colRegistros As New Collection
Dim lErro1 As Long

On Error GoTo Erro_TransfCaixa_Grava

    'abre a transação
    lErro = CF_ECF("Caixa_Transacao_Abrir", lSequencial)
    If lErro <> SUCESSO Then gError 101922

    'preenche os dados chave para uma busca de transferência de caixa
    objTransfCaixaAux.objMovCaixaDe.lTransferencia = objTransfCaixa.objMovCaixaDe.lTransferencia
    objTransfCaixaAux.objMovCaixaPara.lTransferencia = objTransfCaixa.objMovCaixaPara.lTransferencia
       
    'busca se já existe uma transfcaixa com o código da tela
    'i.e. busca por um movimento de caixa casado (de/Para) com lTransferencia preenchidos
    lErro = CF_ECF("TransfCaixa_Le", objTransfCaixaAux)
    If lErro <> SUCESSO And lErro <> 109404 Then gError 109459
    
    'critica a transferrencia atual
    lErro1 = TransfCaixa_Critica(objTransfCaixa, objChequeDe, objTransfCaixaAux)
    If lErro1 <> SUCESSO Then gError 101921
    
    'se achou->alteração
    If lErro = SUCESSO Then
    
        lErro = TransfCaixa_Exclui_EmTrans(objTransfCaixaAux)
        If lErro <> SUCESSO Then gError 109460
        
    Else
    
        If objChequeDe.iStatus = STATUS_EXCLUIDO Then gError 105341
        
    End If
    
    'mover lsequencial para o objtransfcaixa, movimento origem
    objTransfCaixa.objMovCaixaDe.lSequencial = lSequencial
    
    lSequencial = lSequencial + 1
    
    'mover lsequencial para o objtransfcaixa, movimento destino
    objTransfCaixa.objMovCaixaPara.lSequencial = lSequencial
    
    'grava o arquivo de log
    Call TransfCaixa_Gera_Log_Inclusao(colRegistros, objTransfCaixa, objChequePara)
    
    'Função que Grava o Arquivo de Movimento Caixa
    lErro = CF_ECF("MovimentoCaixaECF_Grava", colRegistros)
    If lErro <> SUCESSO Then gError 101947

    lErro = CF_ECF("TransfCaixa_Grava_Memoria", objTransfCaixa, objChequePara)
    If lErro <> SUCESSO Then gError 101924
    
    'fecha a transação
    lErro = CF_ECF("Caixa_Transacao_Fechar", lSequencial)
    If lErro <> SUCESSO Then gError 101923
    
    TransfCaixa_Grava = SUCESSO
    
    Exit Function

Erro_TransfCaixa_Grava:

    TransfCaixa_Grava = gErr
    
    Select Case gErr
    
        Case 101921, 101922, 101923, 101947, 109459, 109460
        
        'se de erro na gravação da memória->sai
        Case 101924
            Call Rotina_ErroECF(vbOKOnly, ERRO_NECESSARIO_FECHAR_APP, gErr)
            GL_objMDIForm.mnuSair_click

        Case 105341
            Call Rotina_ErroECF(vbOKOnly, ERRO_CHEQUEPRE_EXCLUIDO, gErr, objChequeDe.lSequencialCaixa)

        Case 105554
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_JA_TRANSMITIDO, gErr)
       
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175354)
    
    End Select
    
    Call CF_ECF("Caixa_Transacao_Rollback", glTransacaoPAFECF)
    
    Exit Function
    
End Function

Private Function TransfCaixa_Critica(objTransfCaixa As ClassTransfCaixa, objCheque As ClassChequePre, objTransfCaixaExc As ClassTransfCaixa) As Long
'objTransfCaixa eh parametro de input
'objCheque eh parametro de input

Dim lNumIntCheque As Long
Dim lNumInt As Long
Dim lErro As Long
Dim dValor As Double
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto
Dim sNomeAdmMeioPagto As String

Dim dSaldoEmDinheiro As Double

On Error GoTo Erro_TransfCaixa_Critica

    objAdmMeioPagtoCondPagto.iAdmMeioPagto = objTransfCaixa.objMovCaixaDe.iAdmMeioPagto
    objAdmMeioPagtoCondPagto.iParcelamento = objTransfCaixa.objMovCaixaDe.iParcelamento
    Call CF_ECF("Obtem_Nome_AdmMeioPagto", objTransfCaixa.objMovCaixaDe, sNomeAdmMeioPagto)
    objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = sNomeAdmMeioPagto
    objAdmMeioPagtoCondPagto.iTipoCartao = objTransfCaixa.objMovCaixaDe.iTipoCartao

    'seleciona o tipo de transferencia origem
    Select Case objTransfCaixa.objMovCaixaDe.iTipo
    
        Case MOVIMENTOCAIXA_SAIDA_TRANSF_DINHEIRO
        
            '??? 24/08/2016 If gdSaldoDinheiro < objTransfCaixa.objMovCaixaDe.dValor - objTransfCaixaExc.objMovCaixaDe.dValor Then gError 101867
            
            lErro = CF_ECF("SaldoEmDinheiro_Le", dSaldoEmDinheiro)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            If dSaldoEmDinheiro < objTransfCaixa.objMovCaixaDe.dValor - objTransfCaixaExc.objMovCaixaDe.dValor Then gError 101867
        
        Case MOVIMENTOCAIXA_SAIDA_TRANSF_CHEQUE
        
        
        Case MOVIMENTOCAIXA_SAIDA_TRANSF_CARTAO_CREDITO, MOVIMENTOCAIXA_SAIDA_TRANSF_CARTAO_DEBITO

            'se o saldo para o cartao/parc. em questao eh menor q o valor da tela => erro
            Call Obtem_Saldo_AdmMeioPagto(objAdmMeioPagtoCondPagto, gcolCartao, dValor)
            
            If dValor < objTransfCaixa.objMovCaixaDe.dValor - objTransfCaixaExc.objMovCaixaDe.dValor Then gError 101935
        
        Case MOVIMENTOCAIXA_SAIDA_TRANSF_VALETICKET
        
            'se o saldo do ticket quem questao for menor q o valor da tela
            Call Obtem_Saldo_AdmMeioPagto(objAdmMeioPagtoCondPagto, gcolTicket, dValor)
            
            If dValor < objTransfCaixa.objMovCaixaDe.dValor - objTransfCaixaExc.objMovCaixaDe.dValor Then gError 101937
        
        Case MOVIMENTOCAIXA_SAIDA_TRANSF_OUTROS
        
            'verificar se o saldo para a adm/parcelamento em questao eh < q o da tela, se for => erro
            Call Obtem_Saldo_AdmMeioPagto(objAdmMeioPagtoCondPagto, gcolOutros, dValor)
            
            If dValor < objTransfCaixa.objMovCaixaDe.dValor - objTransfCaixaExc.objMovCaixaDe.dValor Then gError 101936
        
    End Select
    
    TransfCaixa_Critica = SUCESSO
    
    Exit Function

Erro_TransfCaixa_Critica:

    TransfCaixa_Critica = gErr
    
    Select Case gErr
    
        Case 101867
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALORDIN_MAIOR_SALDODIN, gErr, Format(objTransfCaixa.objMovCaixaDe.dValor, "STANDARD"), Format(dSaldoEmDinheiro + objTransfCaixaExc.objMovCaixaDe.dValor, "STANDARD"))
        
        Case 101868, ERRO_SEM_MENSAGEM
        
        Case 101935
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALORCRT_MAIOR_SALDOCRT, gErr, Format(objTransfCaixa.objMovCaixaDe.dValor, "STANDARD"), Format(dValor + objTransfCaixaExc.objMovCaixaDe.dValor, "STANDARD"), objTransfCaixa.objMovCaixaDe.iAdmMeioPagto, objTransfCaixa.objMovCaixaDe.iParcelamento)
        
        Case 101936
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALORADM_MAIOR_SALDOADM, gErr, Format(objTransfCaixa.objMovCaixaDe.dValor, "STANDARD"), Format(dValor + objTransfCaixaExc.objMovCaixaDe.dValor, "STANDARD"), objTransfCaixa.objMovCaixaDe.iAdmMeioPagto, objTransfCaixa.objMovCaixaDe.iParcelamento)
        
        Case 101937
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALORTKT_MAIOR_SALDOTKT, gErr, Format(objTransfCaixa.objMovCaixaDe.dValor, "STANDARD"), Format(dValor, "STANDARD"), objTransfCaixa.objMovCaixaDe.iAdmMeioPagto, objTransfCaixa.objMovCaixaDe.iParcelamento)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175355)
        
    End Select
    
    Exit Function

End Function

Private Sub Obtem_Saldo_AdmMeioPagto(objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto, colAdmMeioPagtoCondPagto As Collection, dSaldo As Double)
'Esta função recebe um admmeiopagto, um parcelamento e uma coleção de admmeiopagtocondpagto
'e retorna o saldo de uma determinada combinação AdmMeioPagto/Parcelamento

Dim objAdmMeioPagtoCondPagtoAux As ClassAdmMeioPagtoCondPagto

    dSaldo = 0
    
    For Each objAdmMeioPagtoCondPagtoAux In colAdmMeioPagtoCondPagto
        
        If objAdmMeioPagtoCondPagto.iAdmMeioPagto = objAdmMeioPagtoCondPagtoAux.iAdmMeioPagto And _
           objAdmMeioPagtoCondPagto.iParcelamento = objAdmMeioPagtoCondPagtoAux.iParcelamento And _
           objAdmMeioPagtoCondPagto.iTipoCartao = objAdmMeioPagtoCondPagtoAux.iTipoCartao And _
           objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = objAdmMeioPagtoCondPagtoAux.sNomeAdmMeioPagto And _
           objAdmMeioPagtoCondPagtoAux.iOrigem = OPTION_CF Then
        
            dSaldo = objAdmMeioPagtoCondPagtoAux.dSaldo
            
            Exit Sub
        
        End If
    
    Next

End Sub


Public Sub Desmembra_Dados_TransfCaixa(sLog As String, objTransfCaixa As ClassTransfCaixa, objCheque As ClassChequePre)

Dim iInicio As Integer
Dim iFim As Integer
Dim iMeio As Integer
Dim iTipoRegistro As Integer
Dim iIndice As Integer

    iInicio = 1
    
    'busca o primeiro control
    iMeio = InStr(iInicio, sLog, Chr(vbKeyControl))
    
    'verifica se é inclusão ou exclusão
    iTipoRegistro = StrParaInt(Mid(sLog, iInicio, iMeio - 1))
    
    iIndice = 0
    
    'começo da string( ignorando o tipo de registro)
    iInicio = iMeio + 1
    
    'primeiro escape
    iMeio = InStr(iInicio, sLog, Chr(vbKeyEscape))
    
    iFim = InStr(iInicio, sLog, Chr(vbKeyControl))
    
    If iFim = 0 Then iFim = InStr(iInicio, sLog, Chr(vbKeyEnd))
    
    If iMeio = 0 Then iMeio = iFim
    
    If iTipoRegistro = TIPOREGISTROECF_TRANSFCAIXA_INCLUSAO Then
    
        Do While iMeio <> 0
        
            iIndice = iIndice + 1
            
            Select Case iIndice
            
                'desmembra o movimento de origem
                Case 1: objTransfCaixa.objMovCaixaDe.iTipo = StrParaInt(Mid(sLog, iInicio, iMeio - iInicio))
                Case 2: objTransfCaixa.objMovCaixaDe.lSequencial = StrParaLong(Mid(sLog, iInicio, iMeio - iInicio))
                Case 3: objTransfCaixa.objMovCaixaDe.iAdmMeioPagto = StrParaInt(Mid(sLog, iInicio, iMeio - iInicio))
                Case 4: objTransfCaixa.objMovCaixaDe.iParcelamento = StrParaInt(Mid(sLog, iInicio, iMeio - iInicio))
                Case 5: objTransfCaixa.objMovCaixaDe.dHora = StrParaDbl(Mid(sLog, iInicio, iMeio - iInicio))
                Case 6: objTransfCaixa.objMovCaixaDe.iTipoCartao = StrParaInt(Mid(sLog, iInicio, iMeio - iInicio))
                Case 7: objTransfCaixa.objMovCaixaPara.lCupomFiscal = StrParaLong(Mid(sLog, iInicio, iMeio - iInicio))
                
                'desmembra o movimento de destino
                Case 8: objTransfCaixa.objMovCaixaPara.iTipo = StrParaInt(Mid(sLog, iInicio, iMeio - iInicio))
                Case 9: objTransfCaixa.objMovCaixaPara.lSequencial = StrParaLong(Mid(sLog, iInicio, iMeio - iInicio))
                Case 10: objTransfCaixa.objMovCaixaPara.iAdmMeioPagto = StrParaInt(Mid(sLog, iInicio, iMeio - iInicio))
                Case 11: objTransfCaixa.objMovCaixaPara.iParcelamento = StrParaInt(Mid(sLog, iInicio, iMeio - iInicio))
                Case 12: objTransfCaixa.objMovCaixaPara.dHora = StrParaDbl(Mid(sLog, iInicio, iMeio - iInicio))
                Case 13: objTransfCaixa.objMovCaixaPara.iTipoCartao = StrParaInt(Mid(sLog, iInicio, iMeio - iInicio))
                Case 14: objTransfCaixa.objMovCaixaPara.lCupomFiscal = StrParaLong(Mid(sLog, iInicio, iMeio - iInicio))
                
                'coloca os dados comuns aos dois movimentos
                Case 15
                    objTransfCaixa.objMovCaixaDe.iCaixa = StrParaInt(Mid(sLog, iInicio, iMeio - iInicio))
                    objTransfCaixa.objMovCaixaPara.iCaixa = StrParaInt(Mid(sLog, iInicio, iMeio - iInicio))
                Case 16
                    objTransfCaixa.objMovCaixaDe.iFilialEmpresa = StrParaInt(Mid(sLog, iInicio, iMeio - iInicio))
                    objTransfCaixa.objMovCaixaPara.iFilialEmpresa = StrParaInt(Mid(sLog, iInicio, iMeio - iInicio))
                Case 17
                    objTransfCaixa.objMovCaixaDe.iCodOperador = StrParaInt(Mid(sLog, iInicio, iMeio - iInicio))
                    objTransfCaixa.objMovCaixaPara.iCodOperador = StrParaInt(Mid(sLog, iInicio, iMeio - iInicio))
                Case 18
                    objTransfCaixa.objMovCaixaDe.dtDataMovimento = StrParaDate(Mid(sLog, iInicio, iMeio - iInicio))
                    objTransfCaixa.objMovCaixaPara.dtDataMovimento = StrParaDate(Mid(sLog, iInicio, iMeio - iInicio))
                Case 19
                    objTransfCaixa.objMovCaixaDe.dValor = StrParaDbl(Mid(sLog, iInicio, iMeio - iInicio))
                    objTransfCaixa.objMovCaixaPara.dValor = StrParaDbl(Mid(sLog, iInicio, iMeio - iInicio))
                Case 20
                    objTransfCaixa.objMovCaixaDe.lTransferencia = StrParaLong(Mid(sLog, iInicio, iMeio - iInicio))
                    objTransfCaixa.objMovCaixaPara.lTransferencia = StrParaLong(Mid(sLog, iInicio, iMeio - iInicio))
                
                'preenche o cheque, se for o caso
                Case 21: objCheque.dtDataDeposito = StrParaDate(Mid(sLog, iInicio, iMeio - iInicio))
                Case 22: objCheque.iBanco = StrParaInt(Mid(sLog, iInicio, iMeio - iInicio))
                Case 23: objCheque.iNaoEspecificado = StrParaInt(Mid(sLog, iInicio, iMeio - iInicio))
                Case 24: objCheque.iStatus = StrParaInt(Mid(sLog, iInicio, iMeio - iInicio))
                Case 25: objCheque.lNumero = StrParaLong(Mid(sLog, iInicio, iMeio - iInicio))
                Case 26: objCheque.lSequencialCaixa = StrParaLong(Mid(sLog, iInicio, iMeio - iInicio))
                Case 27: objCheque.sAgencia = Mid(sLog, iInicio, iMeio - iInicio)
                Case 28: objCheque.sContaCorrente = Mid(sLog, iInicio, iMeio - iInicio)
                Case 29
                    objCheque.sCPFCGC = Mid(sLog, iInicio, iMeio - iInicio)
                    
                    'preenche os dados complementares do cheque
                    objCheque.iCaixa = objTransfCaixa.objMovCaixaPara.iCaixa
                    objCheque.iFilialEmpresa = objTransfCaixa.objMovCaixaPara.iFilialEmpresa
                    objCheque.iFilialEmpresaLoja = objTransfCaixa.objMovCaixaPara.iFilialEmpresa
                    objCheque.lCupomFiscal = objTransfCaixa.objMovCaixaPara.lCupomFiscal
                    objCheque.dValor = objTransfCaixa.objMovCaixaPara.dValor
            
            End Select
            
            'atualizo para a primeira posição após o último escape
            iInicio = iMeio + 1
            
            'atualizo o meio
            iMeio = InStr(iInicio, sLog, Chr(vbKeyEscape))
            
            'se o meio ultrapassou o fim atual, significa que é hora de pegar outro control
            If iMeio > iFim Then
                
                'atualizo o meio para o fim atual para pegar o último dado do trecho atual
                iMeio = iFim
                
                'pego o próximo control
                iFim = InStr(iFim + 1, sLog, Chr(vbKeyControl))
                
                'se não encontrar, pego o end
                If iFim = 0 Then iFim = InStr(iMeio, sLog, Chr(vbKeyEnd))
            
            End If
            
            'se não tiver acabado e o meio retornou 0 (não achou o escape) então o meio se torna o fim
            If iInicio < iFim And iMeio = 0 Then iMeio = iFim
        
        Loop
        
        'chamar função que grava a memória
    
    Else
    
        Do While iMeio <> 0
        
            iIndice = iIndice + 1
            
            Select Case iIndice
                Case 1
                    objTransfCaixa.objMovCaixaDe.lTransferencia = StrParaInt(Mid(sLog, iInicio, iMeio - iInicio))
                    objTransfCaixa.objMovCaixaPara.lTransferencia = StrParaInt(Mid(sLog, iInicio, iMeio - iInicio))
            
            End Select
            
            'atualizo para a primeira posição após o último escape
            iInicio = iMeio + 1
            
            'atualizo o meio
            iMeio = InStr(iInicio, sLog, Chr(vbKeyEscape))
            
            'se o meio ultrapassou o fim atual, significa que é hora de pegar outro control
            If iMeio > iFim Then
                
                'atualizo o meio para o fim atual para pegar o último dado do trecho atual
                iMeio = iFim
                
                'pego o próximo control
                iFim = InStr(iFim + 1, sLog, Chr(vbKeyControl))
                
                'se não encontrar, pego o end
                If iFim = 0 Then iFim = InStr(iMeio, sLog, Chr(vbKeyEnd))
            
            End If
            
            'se não tiver acabado e o meio retornou 0 (não achou o escape) então o meio se torna o fim
            If iInicio < iFim And iMeio = 0 Then iMeio = iFim
        
        Loop
    
        'chamar função que exclui da memória
    
    End If

End Sub

Public Sub TransfCaixa_Gera_Log_Inclusao(ByVal colRegistros As Collection, ByVal objTransfCaixa As ClassTransfCaixa, ByVal objChequePara As ClassChequePre)

Dim sLog As String
    
    'inclui o movimento de
    sLog = TIPOREGISTROECF_TRANSFCAIXA_INCLUSAO & Chr(vbKeyControl) & _
           objTransfCaixa.objMovCaixaDe.iTipo & Chr(vbKeyEscape) & _
           objTransfCaixa.objMovCaixaDe.lSequencial & Chr(vbKeyEscape) & _
           objTransfCaixa.objMovCaixaDe.iAdmMeioPagto & Chr(vbKeyEscape) & _
           objTransfCaixa.objMovCaixaDe.iParcelamento & Chr(vbKeyEscape) & _
           objTransfCaixa.objMovCaixaDe.dHora & Chr(vbKeyEscape) & _
           objTransfCaixa.objMovCaixaDe.iTipoCartao & Chr(vbKeyEscape) & _
           objTransfCaixa.objMovCaixaPara.lCupomFiscal & Chr(vbKeyEscape) & _
           objTransfCaixa.objMovCaixaDe.lNumRefInterna & Chr(vbKeyEscape)
           
    'inclui o movimento PAra
    sLog = sLog & _
           objTransfCaixa.objMovCaixaPara.iTipo & Chr(vbKeyEscape) & _
           objTransfCaixa.objMovCaixaPara.lSequencial & Chr(vbKeyEscape) & _
           objTransfCaixa.objMovCaixaPara.iAdmMeioPagto & Chr(vbKeyEscape) & _
           objTransfCaixa.objMovCaixaPara.iParcelamento & Chr(vbKeyEscape) & _
           objTransfCaixa.objMovCaixaPara.dHora & Chr(vbKeyEscape) & _
           objTransfCaixa.objMovCaixaPara.iTipoCartao & Chr(vbKeyEscape) & _
           objTransfCaixa.objMovCaixaPara.lCupomFiscal & Chr(vbKeyEscape) & _
           objTransfCaixa.objMovCaixaPara.lNumRefInterna & Chr(vbKeyEscape)

    sLog = sLog & _
    objTransfCaixa.objMovCaixaDe.iCaixa & Chr(vbKeyEscape) & _
    objTransfCaixa.objMovCaixaDe.iFilialEmpresa & Chr(vbKeyEscape) & _
    objTransfCaixa.objMovCaixaDe.iCodOperador & Chr(vbKeyEscape) & _
    objTransfCaixa.objMovCaixaDe.dtDataMovimento & Chr(vbKeyEscape) & _
    objTransfCaixa.objMovCaixaDe.dValor & Chr(vbKeyEscape) & _
    objTransfCaixa.objMovCaixaDe.lTransferencia & Chr(vbKeyEscape)

    'se entrou um cheque
    If objTransfCaixa.objMovCaixaPara.iTipo = MOVIMENTOCAIXA_ENTRADA_TRANSF_CHEQUE Then
    
        sLog = sLog & _
               objChequePara.dtDataDeposito & Chr(vbKeyEscape) & _
               objChequePara.iBanco & Chr(vbKeyEscape) & _
               objChequePara.iNaoEspecificado & Chr(vbKeyEscape) & _
               objChequePara.iStatus & Chr(vbKeyEscape) & _
               objChequePara.lNumero & Chr(vbKeyEscape) & _
               objChequePara.lSequencialCaixa & Chr(vbKeyEscape) & _
               objChequePara.sAgencia & Chr(vbKeyEscape) & _
               objChequePara.sContaCorrente & Chr(vbKeyEscape) & _
               objChequePara.sCPFCGC & Chr(vbKeyEscape) & _
               objChequePara.iLocalizacao & Chr(vbKeyEscape) & _
               objChequePara.iECF & Chr(vbKeyEscape)
    
    End If
    
    sLog = sLog & Chr(vbKeyEnd)
    
    colRegistros.Add sLog
    
End Sub

Public Sub TransfCaixa_Gera_Log_Exclusao(ByVal colRegistros As Collection, ByVal objTransfCaixa As ClassTransfCaixa)

Dim sLog As String

    sLog = TIPOREGISTROECF_TRANSFCAIXA_EXCLUSAO & Chr(vbKeyControl) & _
           objTransfCaixa.objMovCaixaDe.iFilialEmpresa & Chr(vbKeyEscape) & _
           objTransfCaixa.objMovCaixaDe.iCaixa & Chr(vbKeyEscape) & _
           objTransfCaixa.objMovCaixaDe.lTransferencia & Chr(vbKeyEscape) & _
           Chr(vbKeyEnd)
           
    colRegistros.Add sLog
    
End Sub

Private Function Limpa_Tela_TransfCaixa() As Long
'limpa a tela.. duh...

Dim objCombo As Object

    Call Limpa_Tela(Me)

    'limpa as combos
    For Each objCombo In Me.Controls

      'se for realmente combo
      If TypeName(objCombo) = "ComboBox" Then
         objCombo.ListIndex = -1
      End If
    Next
    
    OptionManualDe.Value = True
    OptionManualPara.Value = True

    Call Limpa_ChequeDe

    'coloca todos os frames invisiveis
    Call Invisibiliza_Frames

End Function

Private Sub Limpa_ChequeDe()

    BancoChqDe.Caption = ""
    AgenciaChqDe.Caption = ""
    ContaChqDe.Caption = ""
    NumeroChqDe.Caption = ""
    ValorChqDe.Caption = ""
    ClienteChqDe.Caption = ""
    CupomFiscalChqDe.Caption = ""
    DataBomParaChqDe.Caption = ""
    

End Sub

Private Sub Invisibiliza_Frames()

    Call Invisibiliza_FramesDe
    Call Invisibiliza_FramesPara

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

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objTransfCaixa As New ClassTransfCaixa
Dim vbMsgResp As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 207997

    'se o codigo não estiver preenchido-> erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 109467
    
    'preenche o movimento de caixa de e o movimento de caixa para
    objTransfCaixa.objMovCaixaDe.lTransferencia = StrParaLong(Codigo.Text)
    
    'procura por movimento de caixa com esse codigo de transferencia
    lErro = CF_ECF("TransfCaixa_Le", objTransfCaixa)
    If lErro <> SUCESSO And lErro <> 109404 Then gError 109468
    
    If lErro = 109404 Then gError 109469
    
    'se já tiver sido transferido para o log ==> nao pode alterar
    If objTransfCaixa.objMovCaixaDe.lNumIntDocLog = MARCADO Then gError 105553
    
    'pergunta se tem certeza que deseja excluir o cara
    vbMsgResp = Rotina_AvisoECF(vbYesNo, AVISO_EXCLUSAO_TRANSFCAIXA, objTransfCaixa.objMovCaixaDe.lTransferencia)
    
    'se sim...
    If vbMsgResp = vbYes Then
    
        'exclui a transfcaixa
        lErro = TransfCaixa_Exclui(objTransfCaixa)
        If lErro <> SUCESSO Then gError 109470
        
        'limpa a tela
        Call Limpa_Tela_TransfCaixa
        
        iAlterado = 0
    
    End If

    Exit Sub
    
Erro_BotaoExcluir_Click:
    
    Select Case gErr
    
        Case 105553
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_JA_TRANSMITIDO, gErr)
    
        Case 109467
            Call Rotina_ErroECF(vbOKOnly, ERRO_TRANSFCAIXA_NAO_PREENCHIDO, gErr)
            
        Case 109469
            Call Rotina_ErroECF(vbOKOnly, ERRO_TRANSFCAIXA_NAO_ENCONTRADO, gErr)
        
        Case 109468, 109470, 207997
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175356)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
   Call Limpa_Tela_TransfCaixa
End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Public Function TransfCaixa_Critica_Exclusao(objTransfCaixa As ClassTransfCaixa) As Long

Dim objCheque As New ClassChequePre
Dim lErro As Long
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto
Dim dValor As Double
Dim sNomeAdmMeioPagto As String

Dim dSaldoEmDinheiro As Double

On Error GoTo Erro_TransfCaixa_Critica_Exclusao

    objAdmMeioPagtoCondPagto.iAdmMeioPagto = objTransfCaixa.objMovCaixaPara.iAdmMeioPagto
    objAdmMeioPagtoCondPagto.iParcelamento = objTransfCaixa.objMovCaixaPara.iParcelamento
    Call CF_ECF("Obtem_Nome_AdmMeioPagto", objTransfCaixa.objMovCaixaPara, sNomeAdmMeioPagto)
    objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = sNomeAdmMeioPagto
    objAdmMeioPagtoCondPagto.iTipoCartao = objTransfCaixa.objMovCaixaPara.iTipoCartao
    
    'seleciona o tipo de transferencia origem
    Select Case objTransfCaixa.objMovCaixaPara.iTipo
    
        Case MOVIMENTOCAIXA_ENTRADA_TRANSF_DINHEIRO
        
            '??? 24/08/2016 If gdSaldoDinheiro < objTransfCaixa.objMovCaixaPara.dValor Then gError 101945
        
            lErro = CF_ECF("SaldoEmDinheiro_Le", dSaldoEmDinheiro)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            If dSaldoEmDinheiro < objTransfCaixa.objMovCaixaPara.dValor Then gError 101945
        
        Case MOVIMENTOCAIXA_ENTRADA_TRANSF_CHEQUE
        
            objCheque.lSequencialCaixa = objTransfCaixa.objMovCaixaPara.lNumRefInterna
            
            lErro = CF_ECF("Cheque_Le_SequencialCaixa", objCheque)
            If lErro <> SUCESSO And lErro <> 109462 Then gError 109463
            
            If lErro = 109462 Then gError 109464
            
            'o cheque para nao pode estar excluido, se estiver ==> erro
            If objCheque.iStatus = STATUS_EXCLUIDO Then gError 105339
            
            'se o cheque tiver sido sangrado ==> nao pode ser transferido
            If objCheque.lNumMovtoSangria <> 0 Then gError 105511
        
        Case MOVIMENTOCAIXA_ENTRADA_TRANSF_CARTAO_CREDITO, MOVIMENTOCAIXA_ENTRADA_TRANSF_CARTAO_DEBITO
        
            'se o saldo para o cartao/parc. em questao eh menor q o valor da tela => erro
            Call Obtem_Saldo_AdmMeioPagto(objAdmMeioPagtoCondPagto, gcolCartao, dValor)
            
            If dValor < objTransfCaixa.objMovCaixaPara.dValor Then gError 101942
        
        Case MOVIMENTOCAIXA_ENTRADA_TRANSF_VALETICKET
        
            'se o saldo do ticket quem questao for menor q o valor da tela
            Call Obtem_Saldo_AdmMeioPagto(objAdmMeioPagtoCondPagto, gcolTicket, dValor)
            
            If dValor < objTransfCaixa.objMovCaixaPara.dValor Then gError 101943
        
        Case MOVIMENTOCAIXA_ENTRADA_TRANSF_OUTROS
        
            'verificar se o saldo para a adm/parcelamento em questao eh < q o da tela, se for => erro
            Call Obtem_Saldo_AdmMeioPagto(objAdmMeioPagtoCondPagto, gcolOutros, dValor)
            
            If dValor < objTransfCaixa.objMovCaixaPara.dValor Then gError 101944
        
        Case Else
    
            gError 105346
    
    End Select
    
    TransfCaixa_Critica_Exclusao = SUCESSO
    
    Exit Function

Erro_TransfCaixa_Critica_Exclusao:

    TransfCaixa_Critica_Exclusao = gErr
    
    Select Case gErr
    
        Case 101942
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALORCRT_MAIOR_SALDOCRT, gErr, Format(objTransfCaixa.objMovCaixaPara.dValor, "STANDARD"), Format(dValor, "STANDARD"), objTransfCaixa.objMovCaixaPara.iAdmMeioPagto, objTransfCaixa.objMovCaixaPara.iParcelamento)
        
        Case 101943
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALORTKT_MAIOR_SALDOTKT, gErr, Format(objTransfCaixa.objMovCaixaPara.dValor, "STANDARD"), Format(dValor, "STANDARD"), objTransfCaixa.objMovCaixaPara.iAdmMeioPagto, objTransfCaixa.objMovCaixaPara.iParcelamento)
        
        Case 101944
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALORADM_MAIOR_SALDOADM, gErr, Format(objTransfCaixa.objMovCaixaPara.dValor, "STANDARD"), Format(dValor, "STANDARD"), objTransfCaixa.objMovCaixaPara.iAdmMeioPagto, objTransfCaixa.objMovCaixaPara.iParcelamento)
        
        Case 101945
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALORDIN_MAIOR_SALDODIN, gErr, Format(objTransfCaixa.objMovCaixaPara.dValor, "STANDARD"), Format(dSaldoEmDinheiro, "STANDARD"))
        
        Case 105339
            Call Rotina_ErroECF(vbOKOnly, ERRO_CHEQUEPRE_EXCLUIDO, gErr, objCheque.lSequencialCaixa)
        
        Case 105346
            Call Rotina_ErroECF(vbOKOnly, ERRO_TIPOMOVTOCAIXA_INVALIDO, gErr, objTransfCaixa.objMovCaixaPara.iTipo)

        Case 105511
            Call Rotina_ErroECF(vbOKOnly, ERRO_CHEQUE_SANGRADO, gErr, objCheque.lSequencialCaixa)

        Case 109463, ERRO_SEM_MENSAGEM

        Case 109464
            Call Rotina_ErroECF(vbOKOnly, ERRO_CHEQUEPRE_NAO_ENCONTRADO2, gErr, objCheque.lSequencialCaixa)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175357)
    
    End Select
    
    Exit Function

End Function

'Private Function Checa_Frame_ChqDe2() As Long
'
'On Error GoTo Erro_Checa_Frame_ChqDe2
'
''    'ou cupom fiscal ou carne devem estar preenchidos
''    If Len(Trim(CarneDe.ClipText)) = 0 And Len(Trim(CupomFiscalChqDe.ClipText)) = 0 Then gError 109452
'
''    'se o cupomfiscal não estiver preenchido->erro
''    If Len(Trim(CupomFiscalChqDe.ClipText)) = 0 Then gError 109480
''
''    'só o cupom fiscal ou (exclusivo) carne devem estar estar preenchidos
''    If Len(Trim(CarneDe.ClipText)) > 0 And Len(Trim(CupomFiscalChqDe.Text)) > 0 Then gError 109453
'
''    '... se o numero, conta,banco ou agencia estiverem preenchidos-> todos os complementares devem estar preenchidos
''    If (Len(Trim(BancoChqDe.ClipText)) > 0 Or Len(Trim(AgenciaChqDe.ClipText)) > 0 Or Len(Trim(NumeroChqDe.ClipText)) > 0 Or Len(Trim(ContaChqDe.ClipText)) > 0) And _
''       (Len(Trim(BancoChqDe.ClipText)) = 0 Or Len(Trim(AgenciaChqDe.ClipText)) = 0 Or Len(Trim(NumeroChqDe.ClipText)) = 0 Or Len(Trim(ContaChqDe.ClipText)) = 0) Then gError 109454
'
'    Checa_Frame_ChqDe2 = SUCESSO
'
'    Exit Function
'
'Erro_Checa_Frame_ChqDe2:
'
'    Checa_Frame_ChqDe2 = gErr
'
'    Select Case gErr
'
''        Case 109452
''            Call Rotina_ErroECF(vbOKOnly, ERRO_CARNE_CUPOMFISCAL_NAO_PREENHIDOS, gErr)
'
'        Case 109480
'            Call Rotina_ErroECF(vbOKOnly, ERRO_CUPOMFISCAL_NAO_PREENCHIDO, gErr)
'
''        Case 109453
''            Call Rotina_ErroECF(vbOKOnly, ERRO_CARNE_CUPOMFISCAL_PREENCHIDOS, gErr)
'
'        Case 109454
'            Call Rotina_ErroECF(vbOKOnly, ERRO_GRUPOCHQDE_NAO_PREENCHIDO, gErr)
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175358)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Sub Carrega_AdmMeioPagto_Cartao(AdmCrt As ComboBox, iTipoMeioPagto As Integer)

Dim objAdmMeioPagto As ClassAdmMeioPagto

On Error GoTo Erro_Carrega_AdmMeioPagto_Cartao

    'limpa a combo recebida por parâmetro
    AdmCrt.Clear

    'varre a coleção global buscando admmeiopagtos
    'que possuem tipomeiopagto = ao tipo passado por parâmetro
    'e as adiciona à combo passada por parâmetro
    For Each objAdmMeioPagto In gcolAdmMeioPagto

        If objAdmMeioPagto.iTipoMeioPagto = iTipoMeioPagto Then
            AdmCrt.AddItem (objAdmMeioPagto.iCodigo & SEPARADOR & objAdmMeioPagto.sNome)
            AdmCrt.ItemData(AdmCrt.NewIndex) = objAdmMeioPagto.iCodigo

        End If

    Next

    Exit Sub

Erro_Carrega_AdmMeioPagto_Cartao:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175359)

    End Select

    Exit Sub

End Sub


Private Sub ParcelamentoCrtDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_ParcelamentoCrtDe_Validate

    'Se foi preenchida a Combo
    If Len(Trim(ParcelamentoCrtDe.Text)) > 0 Then

        'Se o parcelamento atual for diferente do selecionado
        If ParcelamentoCrtDe.ListIndex = -1 Then

            'Verifica se existe o item na List da Combo. Se existir seleciona.
            lErro = Combo_Seleciona(ParcelamentoCrtDe, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 101853

            'Nao existe o item com o CODIGO na List da ComboBox
            If lErro = 6730 Then gError 101854

            'Nao existe o item com a STRING na List da ComboBox
            If lErro = 6731 Then gError 101855

        End If

    End If

    Exit Sub

Erro_ParcelamentoCrtDe_Validate:

    Cancel = True

    Select Case gErr

        Case 101853

        Case 101854
            Call Rotina_ErroECF(vbOKOnly, ERRO_CODIGO_PARCELAMENTO_NAO_EXISTENTE, gErr, ParcelamentoCrtDe.Text)

        Case 101855
            Call Rotina_ErroECF(vbOKOnly, ERRO_NOME_PARCELAMENTO_NAO_EXISTENTE, gErr, ParcelamentoCrtDe.Text)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175360)

    End Select

    Exit Sub

End Sub


Private Sub AdmCrtDe_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AdmCrtPara_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AdmOutDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AdmOutDe_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AdmOutPara_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AdmOutPara_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AdmTktDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AdmTktDe_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AdmTktPara_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AdmTktPara_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AgenciaChqDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AgenciaChqPara_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AgenciaChqPara_GotFocus()
    Call MaskEdBox_TrataGotFocus(AgenciaChqPara, iAlterado)
End Sub

'Private Sub BancoChqDe_Change()
'    iAlterado = REGISTRO_ALTERADO
'End Sub
'
'Private Sub BancoChqDe_GotFocus()
'    Call MaskEdBox_TrataGotFocus(BancoChqDe, iAlterado)
'End Sub
'
'Private Sub BancoChqDe_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_BancoChqDe_Validate
'
'    'se o campo estiver preenchido
'    If Len(Trim(BancoChqDe.ClipText)) > 0 Then
'
'        'critica o conteudo
'        lErro = Inteiro_Critica(BancoChqDe.Text)
'        If lErro <> SUCESSO Then gError 101806
'
'    End If
'
'    Exit Sub
'
'Erro_BancoChqDe_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 101806
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175361)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'$$$$$$$$$$$

Private Sub LabelCodigo_Click()

Dim objTransfCaixa As New ClassTransfCaixa
Dim lErro As Long

On Error GoTo Erro_LabelCodigo_Click

    'chamar o browser TransfCaixaLista
    Call Chama_TelaECF_Modal("TransfCaixaLista", objTransfCaixa)
    
    'se o obj tiver preenchido, traz pra tela..
    If giRetornoTela = vbOK Then
    
        lErro = CF_ECF("TransfCaixa_Le", objTransfCaixa)
        If lErro <> SUCESSO And lErro <> 109404 Then gError 109483
        
        If lErro = 109404 Then gError 109484
        
        lErro = Traz_TransfCaixa_Tela(objTransfCaixa)
        If lErro <> SUCESSO Then gError 101802
    
    End If
    
    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr
    
        Case 101802, 109483
        
        Case 109484
            Call Rotina_ErroECF(vbOKOnly, ERRO_TRANSFCAIXA_NAO_ENCONTRADO, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175362)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BancoChqPara_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub BancoChqPara_GotFocus()
    Call MaskEdBox_TrataGotFocus(BancoChqPara, iAlterado)
End Sub

Private Sub BancoChqPara_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_BancoChqPara_Validate

    'se o campo estiver preenchido
    If Len(Trim(BancoChqPara.ClipText)) > 0 Then

        'critica o conteudo
        lErro = Inteiro_Critica(BancoChqPara.Text)
        If lErro <> SUCESSO Then gError 101812

    End If

    Exit Sub

Erro_BancoChqPara_Validate:

    Cancel = True

    Select Case gErr

        Case 101812

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175363)

    End Select

    Exit Sub

End Sub

Private Sub NumeroChqDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NumeroChqPara_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NumeroChqPara_GotFocus()
    Call MaskEdBox_TrataGotFocus(NumeroChqPara, iAlterado)
End Sub

Private Sub NumeroChqPara_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NumeroChqPara_Validate

    'se campo estiver preenchido
    If Len(Trim(NumeroChqPara.ClipText)) > 0 Then

        'critica o conteudo
        lErro = Long_Critica(NumeroChqPara.Text)
        If lErro <> SUCESSO Then gError 101814

    End If

    Exit Sub

Erro_NumeroChqPara_Validate:

    Cancel = True

    Select Case gErr

        Case 101814

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175364)

    End Select

    Exit Sub

End Sub

Private Sub ParcelamentoCrtDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ParcelamentoCrtDe_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ParcelamentoCrtPara_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ParcelamentoCrtPara_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub SeqChqDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub SeqChqDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(SeqChqDe, iAlterado)
End Sub

Private Sub SeqChqDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCheque As New ClassChequePre

On Error GoTo Erro_SeqChqDe_Validate

    If Len(Trim(SeqChqDe.Text)) > 0 Then

        'preenche os dados chave do cheque
        objCheque.lSequencialCaixa = StrParaLong(SeqChqDe.Text)
    
        'lê o cheque na coleção global
        lErro = CF_ECF("Cheque_Le_SequencialCaixa", objCheque)
        If lErro <> SUCESSO And lErro <> 109462 Then gError 109402
    
        'se não achou, erro
        If lErro = 109462 Then gError 109403
    
        If objCheque.lNumMovtoSangria <> 0 Then gError 105816
    
        'preenche a tela
        Call Preenche_Frame_ChqDe(objCheque)

    Else
    
        Call Limpa_ChequeDe

    End If

    Exit Sub

Erro_SeqChqDe_Validate:

    Cancel = True

    Select Case gErr

        Case 109401, 109402

        Case 109403
            Call Rotina_ErroECF(vbOKOnly, ERRO_CHEQUEPRE_NAO_ENCONTRADO2, gErr, objCheque.lSequencialCaixa)

        Case 105816
            Call Rotina_ErroECF(vbOKOnly, ERRO_CHEQUE_SANGRADO, gErr, objCheque.lSequencialCaixa)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175365)

    End Select

    Exit Sub

End Sub

Private Sub ValorChqDe_Change()
    iAlterado = REGISTRO_ALTERADO
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
        If lErro <> SUCESSO Then gError 101811

    End If

    Exit Sub

Erro_ValorCrtDe_Validate:

    Cancel = True

    Select Case gErr

        Case 101811

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175366)

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
        If lErro <> SUCESSO Then gError 105347

    End If

    Exit Sub

Erro_ValorDinDe_Validate:

    Cancel = True

    Select Case gErr

        Case 105347

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175367)

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
        If lErro <> SUCESSO Then gError 105348

    End If

    Exit Sub

Erro_ValorOutDe_Validate:

    Cancel = True

    Select Case gErr

        Case 105348

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175368)

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
        If lErro <> SUCESSO Then gError 105349

    End If

    Exit Sub

Erro_ValorTktDe_Validate:

    Cancel = True

    Select Case gErr

        Case 105349

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175369)

    End Select

    Exit Sub

End Sub

'Private Sub CarneDe_Change()
'    iAlterado = REGISTRO_ALTERADO
'End Sub

'Private Sub CarneDe_GotFocus()
'    Call MaskEdBox_TrataGotFocus(CarneDe, iAlterado)
'End Sub

'Private Sub CarnePara_Change()
'    iAlterado = REGISTRO_ALTERADO
'End Sub

'Private Sub CarnePara_GotFocus()
'    Call MaskEdBox_TrataGotFocus(CarnePara, iAlterado)
'End Sub

Private Sub ClienteChqDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ClienteChqPara_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ClienteChqPara_GotFocus()
    Call MaskEdBox_TrataGotFocus(ClienteChqPara, iAlterado)
End Sub

Private Sub Codigo_Change()
    iAlterado = REGISTRO_ALTERADO
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
        If lErro <> SUCESSO Then gError 101803

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 101803

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175370)

    End Select

    Exit Sub

End Sub

Private Sub ContaChqDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ContaChqPara_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ContaChqPara_GotFocus()
    Call MaskEdBox_TrataGotFocus(ContaChqPara, iAlterado)
End Sub

Private Sub CupomFiscalChqDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CupomFiscalChqPara_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CupomFiscalChqPara_GotFocus()
    Call MaskEdBox_TrataGotFocus(CupomFiscalChqPara, iAlterado)
End Sub

Private Sub DataBomParaChqDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataBomParaChqPara_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataBomParaChqPara_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataBomParaChqPara, iAlterado)
End Sub

Private Sub DataBomParaChqPara_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NumeroChqPara_Validate

    'se campo estiver preenchido
    If Len(Trim(DataBomParaChqPara.ClipText)) > 0 Then

        'critica o conteudo
        lErro = Data_Critica(DataBomParaChqPara.Text)
        If lErro <> SUCESSO Then gError 101813

    End If

    Exit Sub

Erro_NumeroChqPara_Validate:

    Cancel = True

    Select Case gErr

        Case 101813

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175371)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'Função que Incrementa o Código Atraves da Tecla F2

Dim lErro As Long

On Error GoTo Erro_UserControl_KeyDown

    Select Case KeyCode

    'Função que Incrementa a transferencia
        Case KEYCODE_PROXIMO_NUMERO
            BotaoProxNum.SetFocus
            Call BotaoProxNum_Click

        Case KEYCODE_BROWSER
    
            If Me.ActiveControl Is Codigo Then
                Call LabelCodigo_Click
            ElseIf Me.ActiveControl Is SeqChqDe Then
                Call BotaoChequeDe_Click
            ElseIf Me.ActiveControl Is CupomFiscalChqPara Then
                Call LabelCupomFiscalChqPara_Click
            ElseIf Me.ActiveControl Is ECFChqPara Then
                Call LabelCupomFiscalChqPara_Click
            End If
    
        Case vbKeyF5
            If Not TrocaFoco(Me, BotaoGravar) Then Exit Sub
            Call BotaoGravar_Click
            
        Case vbKeyF6
            If Not TrocaFoco(Me, BotaoExcluir) Then Exit Sub
            Call BotaoExcluir_Click
            
        Case vbKeyF7
            If Not TrocaFoco(Me, BotaoLimpar) Then Exit Sub
            Call BotaoLimpar_Click
            
        Case vbKeyF8
            If Not TrocaFoco(Me, BotaoFechar) Then Exit Sub
            Call BotaoFechar_Click
    
        Case vbKeyF9
            If Not TrocaFoco(Me, BotaoChequeDe) Then Exit Sub
            Call BotaoChequeDe_Click

    End Select
    

    Exit Sub

Erro_UserControl_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175372)

    End Select

    Exit Sub

End Sub

Private Sub BotaoChequeDe_Click()

Dim objCheque As New ClassChequePre
Dim lErro As Long

On Error GoTo Erro_BotaoChequeDe_Click

    'chama o browser chequecaixalista passando objcheque
    Call Chama_TelaECF_Modal("ChequeLista", objCheque)

    'se objcheque tiver preenchido, chama pra tela..
    If giRetornoTela = vbOK Then

        'limpa a frame do cheque
        Call Limpa_Frame_ChqDe

        'preenche a frame do cheque
        Call Preenche_Frame_ChqDe(objCheque)

    End If

    Exit Sub

Erro_BotaoChequeDe_Click:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175373)

    End Select

    Exit Sub

End Sub

Private Sub TipoMeioPagtoDe_Click()

    iAlterado = REGISTRO_ALTERADO

    'chama o tratamento de frames
    If TipoMeioPagtoDe.ListIndex <> -1 Then Call TipoMeioPagto_TrataFramesDe

End Sub


Private Sub TipoMeioPagtoPara_Click()

    iAlterado = REGISTRO_ALTERADO

    'chama o tratamento de frames
    If TipoMeioPagtoPara.ListIndex <> -1 Then Call TipoMeioPagto_TrataFramesPara

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Transferência de Caixa"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TransfCaixa"

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

'????Procurar saber com o Luiz
Private Sub BotaoProxNum_Click()
'Coloca o próximo número a ser gerado na tela

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Gera número automático.
    lErro = CF_ECF("Caixa_Obtem_NumAutomatico", lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 101800

    'Joga o código na tela
    Codigo.Text = CStr(lCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 101800

        Case Else
             Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175374)

    End Select

    Exit Sub

End Sub

Public Function TransfCaixa_Exclui(objTransfCaixa As ClassTransfCaixa) As Long

Dim lErro As Long
Dim objCheque As New ClassChequePre
Dim lSequencial As Long

On Error GoTo Erro_TransfCaixa_Exclui

    lErro = CF_ECF("Caixa_Transacao_Abrir", lSequencial)
    If lErro <> SUCESSO Then gError 101911
    
    lErro = TransfCaixa_Exclui_EmTrans(objTransfCaixa)
    If lErro <> SUCESSO Then gError 101912
    
    lErro = CF_ECF("Caixa_Transacao_Fechar", lSequencial)
    If lErro <> SUCESSO Then gError 101914
    
    TransfCaixa_Exclui = SUCESSO
    
    Exit Function
    
Erro_TransfCaixa_Exclui:
    
    TransfCaixa_Exclui = gErr
    
    Select Case gErr
    
        Case 101911, 101912, 101914
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175375)
    
    End Select
    
    'Rollback na transação
    Call CF_ECF("Caixa_Transacao_Rollback", glTransacaoPAFECF)
    
    Exit Function
    
End Function

Public Function TransfCaixa_Exclui_EmTrans(objTransfCaixa As ClassTransfCaixa) As Long
    
Dim lErro As Long
Dim objCheque As New ClassChequePre
Dim lSequencial As Long
Dim colRegistros As New Collection
Dim lTamanho As Long
Dim sRetorno As String
Dim sArquivo As String

On Error GoTo Erro_TransfCaixa_Exclui_EmTrans
    
    lTamanho = 255
    sRetorno = String(lTamanho, 0)
    
    'Obtém o diretório onde deve ser armazenado o arquivo com dados do backoffice
    Call GetPrivateProfileString(APLICACAO_DADOS, "DirDadosCC", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    
    'Retira os espaços no final da string
    sRetorno = StringZ(sRetorno)
    
    'Se não encontrou
    If Len(Trim(sRetorno)) = 0 Or sRetorno = CStr(CONSTANTE_ERRO) Then gError 127097
    
    If right(sRetorno, 1) <> "\" Then sRetorno = sRetorno & "\"
    
    sArquivo = sRetorno & giCodEmpresa & "_" & giFilialEmpresa & "_" & NOME_ARQUIVOCC
    
    'Abre o arquivo de retorno
    Open sArquivo For Input Lock Read Write As #10

    'faz a critica para verificar se a exclusao podera ser feita
    lErro = TransfCaixa_Critica_Exclusao(objTransfCaixa)
    If lErro <> SUCESSO Then gError 101910
    
    lTamanho = 255
    sRetorno = String(lTamanho, 0)
        
    'Obtém a ultima transacao transferida
    Call GetPrivateProfileString(APLICACAO_DADOS, "UltimaTransacaoTransf", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    
    'se o numero da ultima transacao transferida ultrapassar o numero da transacao do movimento de caixa
    If objTransfCaixa.objMovCaixaDe.lSequencial <> 0 And StrParaLong(sRetorno) > objTransfCaixa.objMovCaixaDe.lSequencial Then gError 133851

    'se o numero da ultima transacao transferida ultrapassar o numero da transacao do movimento de caixa
    If objTransfCaixa.objMovCaixaPara.lSequencial <> 0 And StrParaLong(sRetorno) > objTransfCaixa.objMovCaixaPara.lSequencial Then gError 133852
    
    'chama a gravação de arquivo, indicando que é uma exclusão
    'por isso, o booleano é false
    Call TransfCaixa_Gera_Log_Exclusao(colRegistros, objTransfCaixa)
    
    'Função que Grava o Arquivo de Movimento Caixa
    lErro = CF_ECF("MovimentoCaixaECF_Grava", colRegistros)
    If lErro <> SUCESSO Then gError 101947
    
    'chama funcao q ira cuidar de adicionar/subtrair os totais consolidados.
    'apesar do nome induzir, ela nao gravara na memoria, mas sim retirara...
    lErro = CF_ECF("TransfCaixa_Exclui_Memoria", objTransfCaixa)
    If lErro <> SUCESSO Then gError 101915
    
    Close #10
    
    TransfCaixa_Exclui_EmTrans = SUCESSO
    
    Exit Function
    
Erro_TransfCaixa_Exclui_EmTrans:
    
    Close #10
    
    TransfCaixa_Exclui_EmTrans = gErr
    
    Select Case gErr
    
        Case 101910, 101947
        
        Case 101915
            Call Rotina_ErroECF(vbOKOnly, ERRO_NECESSARIO_FECHAR_APP, gErr)
            GL_objMDIForm.mnuSair_click
        
        Case 133851, 133852
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_JA_TRANSMITIDO, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175376)
    
    End Select
    
    Exit Function

End Function


Private Sub Limpa_Frame_ChqDe()

On Error GoTo Erro_Limpa_Frame_ChqDe

    BancoChqDe.Caption = ""
    AgenciaChqDe.Caption = ""
    ContaChqDe.Caption = ""
    NumeroChqDe.Caption = ""
    ClienteChqDe.Caption = ""
    CupomFiscalChqDe.Caption = ""
    ValorChqDe.Caption = ""
    DataBomParaChqDe.Caption = ""


    Exit Sub

Erro_Limpa_Frame_ChqDe:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175377)

    End Select

    Exit Sub

End Sub

Public Function Nome_Extrai(sTexto As String) As String
'Função que retira de um texto no formato "Codigo - Nome" apenas o nome.

Dim iPosicao As Integer
Dim sString As String

    iPosicao = InStr(1, sTexto, "-")
    sString = Mid(sTexto, iPosicao + 1)
    
    Nome_Extrai = sString
    
    Exit Function

End Function

