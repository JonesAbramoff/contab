VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl OperacaoCaixaECF 
   ClientHeight    =   5580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9585
   ScaleHeight     =   5580
   ScaleWidth      =   9585
   Begin VB.Frame FrameOp 
      BorderStyle     =   0  'None
      Height          =   3825
      Index           =   5
      Left            =   210
      TabIndex        =   22
      Top             =   1605
      Visible         =   0   'False
      Width           =   9150
      Begin VB.CommandButton BotaoSangriaBol 
         Caption         =   "Sangria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   7110
         TabIndex        =   54
         Top             =   3075
         Width           =   1830
      End
      Begin VB.Frame FrameCartao 
         Caption         =   "Comprovantes de Venda"
         Height          =   2970
         Left            =   75
         TabIndex        =   23
         Top             =   -15
         Width           =   9090
         Begin VB.CommandButton BotaoDesmarcarTodosBol 
            Caption         =   "Desmarcar Todos"
            Height          =   585
            Left            =   1815
            Picture         =   "OperacaoCaixaECF.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   2265
            Width           =   1440
         End
         Begin VB.ComboBox TipoParcelamentoBol 
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
            Height          =   315
            ItemData        =   "OperacaoCaixaECF.ctx":11E2
            Left            =   5145
            List            =   "OperacaoCaixaECF.ctx":11E4
            TabIndex        =   88
            Top             =   570
            Width           =   1500
         End
         Begin VB.ComboBox AdmBol 
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
            Height          =   315
            Left            =   2220
            TabIndex        =   87
            Top             =   525
            Width           =   1560
         End
         Begin VB.ComboBox TipoTerminalBol 
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
            Height          =   315
            ItemData        =   "OperacaoCaixaECF.ctx":11E6
            Left            =   6705
            List            =   "OperacaoCaixaECF.ctx":11E8
            TabIndex        =   86
            Top             =   570
            Width           =   1260
         End
         Begin VB.CheckBox SelecionadoBol 
            Height          =   195
            Left            =   1155
            TabIndex        =   85
            Top             =   555
            Width           =   1020
         End
         Begin MSMask.MaskEdBox ValorBol 
            Height          =   300
            Left            =   3885
            TabIndex        =   89
            Top             =   525
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin VB.CommandButton BotaoMarcarTodosBol 
            Caption         =   "Marcar Todos"
            Height          =   585
            Left            =   210
            Picture         =   "OperacaoCaixaECF.ctx":11EA
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   2265
            Width           =   1440
         End
         Begin MSFlexGridLib.MSFlexGrid GridCartoes 
            Height          =   1815
            Left            =   1005
            TabIndex        =   84
            Top             =   345
            Width           =   7305
            _ExtentX        =   12885
            _ExtentY        =   3201
            _Version        =   393216
            Rows            =   5
            Cols            =   5
            FixedCols       =   0
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Parcto"
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
            Left            =   5265
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   93
            Top             =   90
            Width           =   555
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
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
            Left            =   4080
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   92
            Top             =   105
            Width           =   450
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Administradora"
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
            Left            =   2460
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   91
            Top             =   75
            Width           =   1260
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Terminal"
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
            Left            =   6750
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   90
            Top             =   105
            Width           =   1170
         End
         Begin VB.Label LabelTotalCComprovBol 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7305
            TabIndex        =   25
            Top             =   2355
            Width           =   1350
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
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
            Left            =   6735
            TabIndex        =   24
            Top             =   2385
            Width           =   510
         End
      End
      Begin MSMask.MaskEdBox ValorSComprovBol 
         Height          =   285
         Left            =   2295
         TabIndex        =   55
         Top             =   3420
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label LabelValorCComprovBol 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5805
         TabIndex        =   60
         Top             =   3420
         Width           =   1125
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Valor com Comprovante:"
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
         Left            =   3630
         TabIndex        =   59
         Top             =   3465
         Width           =   2085
      End
      Begin VB.Label LabelSaldoSComprovBol 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2280
         TabIndex        =   58
         Top             =   3045
         Width           =   1125
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Total sem Comprovante:"
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
         Left            =   165
         TabIndex        =   57
         Top             =   3075
         Width           =   2070
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Valor sem Comprovante:"
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
         Left            =   150
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   56
         Top             =   3465
         Width           =   2070
      End
   End
   Begin VB.Frame FrameOp 
      BorderStyle     =   0  'None
      Height          =   3825
      Index           =   1
      Left            =   195
      TabIndex        =   26
      Top             =   1545
      Width           =   9150
      Begin VB.Frame Frame3 
         Caption         =   "Caixa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   240
         TabIndex        =   127
         Top             =   105
         Width           =   3945
         Begin VB.CommandButton BotaoReducaoZ 
            Caption         =   "Redução Z"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2145
            TabIndex        =   131
            Top             =   915
            Width           =   1695
         End
         Begin VB.CommandButton BotaoLeituraX 
            Caption         =   "Leitura X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   195
            TabIndex        =   130
            Top             =   915
            Width           =   1695
         End
         Begin VB.CommandButton BotaoFechamento 
            Caption         =   "Fechamento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2145
            TabIndex        =   129
            Top             =   465
            Width           =   1695
         End
         Begin VB.CommandButton BotaoAbertura 
            Caption         =   "Abertura"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   195
            TabIndex        =   128
            Top             =   465
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Sessão"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   240
         TabIndex        =   123
         Top             =   2025
         Width           =   4035
         Begin VB.CommandButton BotaoSuspendeSessao 
            Caption         =   "Suspende"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2145
            TabIndex        =   126
            Top             =   465
            Width           =   1695
         End
         Begin VB.CommandButton BotaoEncerraSessao 
            Caption         =   "Encerra"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   210
            TabIndex        =   125
            Top             =   930
            Width           =   1695
         End
         Begin VB.CommandButton BotaoIniciaSessao 
            Caption         =   "Inicia"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   225
            TabIndex        =   124
            Top             =   465
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Saldos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   4440
         TabIndex        =   27
         Top             =   135
         Width           =   4350
         Begin VB.Label Dinheiro 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2250
            TabIndex        =   39
            Top             =   300
            Width           =   1695
         End
         Begin VB.Label ChequeVista 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "98398398,67"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2250
            TabIndex        =   38
            Top             =   780
            Width           =   1695
         End
         Begin VB.Label CartaoVista 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2250
            TabIndex        =   37
            Top             =   1275
            Width           =   1695
         End
         Begin VB.Label CartaoPrazo 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2250
            TabIndex        =   36
            Top             =   2310
            Width           =   1695
         End
         Begin VB.Label ChequePre 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2250
            TabIndex        =   35
            Top             =   1770
            Width           =   1695
         End
         Begin VB.Label Outros 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2250
            TabIndex        =   34
            Top             =   2850
            Width           =   1695
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Cheque Pré:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   615
            TabIndex        =   33
            Top             =   1275
            Width           =   1500
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Cartão a vista:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            TabIndex        =   32
            Top             =   1755
            Width           =   1755
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Cheque a vista:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   240
            TabIndex        =   31
            Top             =   765
            Width           =   1875
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Dinheiro:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1020
            TabIndex        =   30
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Vales/Tickets:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   420
            TabIndex        =   29
            Top             =   2835
            Width           =   1695
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Cartão a prazo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   240
            TabIndex        =   28
            Top             =   2265
            Width           =   1875
         End
      End
   End
   Begin VB.Frame FrameOp 
      BorderStyle     =   0  'None
      Height          =   3825
      Index           =   4
      Left            =   195
      TabIndex        =   49
      Top             =   1545
      Visible         =   0   'False
      Width           =   9150
      Begin VB.CommandButton BotaoSangriaCh 
         Caption         =   "Sangria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   7185
         TabIndex        =   62
         Top             =   3015
         Width           =   1830
      End
      Begin VB.Frame FrameCheque 
         Caption         =   "Cheques"
         Height          =   2745
         Left            =   225
         TabIndex        =   61
         Top             =   30
         Width           =   8850
         Begin VB.CommandButton BotaoDesmarcarTodosCh 
            Caption         =   "Desmarcar Todos"
            Height          =   585
            Left            =   1785
            Picture         =   "OperacaoCaixaECF.ctx":2204
            Style           =   1  'Graphical
            TabIndex        =   98
            Top             =   2055
            Width           =   1440
         End
         Begin VB.CommandButton BotaoMarcarTodosCh 
            Caption         =   "Marcar Todos"
            Height          =   585
            Left            =   180
            Picture         =   "OperacaoCaixaECF.ctx":33E6
            Style           =   1  'Graphical
            TabIndex        =   97
            Top             =   2040
            Width           =   1440
         End
         Begin VB.CheckBox SelecionadoCh 
            Height          =   195
            Left            =   315
            TabIndex        =   94
            Top             =   510
            Width           =   1020
         End
         Begin MSMask.MaskEdBox ClienteCh 
            Height          =   300
            Left            =   7185
            TabIndex        =   70
            Top             =   510
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataDepositoCh 
            Height          =   300
            Left            =   4875
            TabIndex        =   71
            Top             =   510
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ContaCh 
            Height          =   300
            Left            =   3120
            TabIndex        =   72
            Top             =   510
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NumChequeCh 
            Height          =   300
            Left            =   4020
            TabIndex        =   73
            Top             =   510
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox AgenciaCh 
            Height          =   300
            Left            =   2265
            TabIndex        =   74
            Top             =   510
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox BancoCh 
            Height          =   300
            Left            =   1365
            TabIndex        =   75
            Top             =   510
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorCh 
            Height          =   300
            Left            =   5910
            TabIndex        =   76
            Top             =   510
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridCheques 
            Height          =   1650
            Left            =   150
            TabIndex        =   69
            Top             =   270
            Width           =   8580
            _ExtentX        =   15134
            _ExtentY        =   2910
            _Version        =   393216
            Rows            =   5
            Cols            =   5
            FixedCols       =   0
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
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
            Left            =   6690
            TabIndex        =   100
            Top             =   2175
            Width           =   510
         End
         Begin VB.Label LabelTotalCComprovCh 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7260
            TabIndex        =   99
            Top             =   2145
            Width           =   1350
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            Height          =   195
            Left            =   7620
            TabIndex        =   83
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Número"
            Height          =   195
            Left            =   4260
            TabIndex        =   82
            Top             =   135
            Width           =   555
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            Height          =   195
            Left            =   6345
            TabIndex        =   81
            Top             =   135
            Width           =   360
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Banco"
            Height          =   195
            Left            =   1650
            TabIndex        =   80
            Top             =   165
            Width           =   465
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Bom para"
            Height          =   195
            Left            =   5070
            TabIndex        =   79
            Top             =   105
            Width           =   675
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Conta"
            Height          =   195
            Left            =   3435
            TabIndex        =   78
            Top             =   135
            Width           =   420
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Agência"
            Height          =   195
            Left            =   2580
            TabIndex        =   77
            Top             =   135
            Width           =   585
         End
      End
      Begin MSMask.MaskEdBox ValorSComprovCh 
         Height          =   285
         Left            =   2355
         TabIndex        =   63
         Top             =   3360
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "Valor sem Comprovante:"
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
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   68
         Top             =   3405
         Width           =   2070
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "Total sem Comprovante:"
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
         TabIndex        =   67
         Top             =   3000
         Width           =   2070
      End
      Begin VB.Label LabelSaldoSComprovCh 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2355
         TabIndex        =   66
         Top             =   2970
         Width           =   1125
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "Valor com Comprovante:"
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
         Left            =   3705
         TabIndex        =   65
         Top             =   3405
         Width           =   2085
      End
      Begin VB.Label LabelValorCComprovCh 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5880
         TabIndex        =   64
         Top             =   3360
         Width           =   1125
      End
   End
   Begin VB.Frame FrameOp 
      BorderStyle     =   0  'None
      Height          =   3825
      Index           =   7
      Left            =   210
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   9150
      Begin VB.CommandButton BotaoLeituraMemoriaFiscal 
         Caption         =   "Leitura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   6990
         TabIndex        =   10
         Top             =   3060
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Intervalo de Datas"
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
         TabIndex        =   2
         Top             =   1230
         Width           =   2160
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Intervalo de Reduções"
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
         TabIndex        =   1
         Top             =   1605
         Width           =   2340
      End
      Begin VB.Frame FrameReducoes 
         Caption         =   "Intervalo de Reduções"
         Height          =   885
         Left            =   3180
         TabIndex        =   11
         Top             =   1950
         Width           =   4020
         Begin MSMask.MaskEdBox ReducaoInicial 
            Height          =   300
            Left            =   675
            TabIndex        =   12
            Top             =   360
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "dd/mm/yyyy"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ReducaoFinal 
            Height          =   300
            Left            =   2490
            TabIndex        =   13
            Top             =   360
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "dd/mm/yyyy"
            PromptChar      =   " "
         End
         Begin VB.Label Label25 
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
            Left            =   2085
            TabIndex        =   15
            Top             =   420
            Width           =   360
         End
         Begin VB.Label Label28 
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
            Height          =   240
            Left            =   285
            TabIndex        =   14
            Top             =   390
            Width           =   345
         End
      End
      Begin VB.Frame FrameDatas 
         Caption         =   "Intervalo de Datas"
         Height          =   885
         Left            =   3195
         TabIndex        =   3
         Top             =   1005
         Width           =   4020
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   315
            Left            =   1575
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   390
            Width           =   240
            _ExtentX        =   318
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataInicial 
            Height          =   300
            Left            =   630
            TabIndex        =   5
            Top             =   390
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDown2 
            Height          =   315
            Left            =   3390
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   390
            Width           =   240
            _ExtentX        =   318
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataFinal 
            Height          =   300
            Left            =   2445
            TabIndex        =   7
            Top             =   390
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label dIni 
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
            Height          =   240
            Left            =   240
            TabIndex        =   9
            Top             =   420
            Width           =   345
         End
         Begin VB.Label dFim 
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
            Left            =   2040
            TabIndex        =   8
            Top             =   450
            Width           =   360
         End
      End
   End
   Begin VB.Frame FrameOp 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   3825
      Index           =   6
      Left            =   225
      TabIndex        =   48
      Top             =   1545
      Visible         =   0   'False
      Width           =   9150
      Begin VB.CommandButton BotaoSangriaOutro 
         Caption         =   "Sangria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   7170
         TabIndex        =   115
         Top             =   3120
         Width           =   1830
      End
      Begin VB.Frame FrameOutros 
         Caption         =   "Outros"
         Height          =   2940
         Left            =   150
         TabIndex        =   95
         Top             =   30
         Width           =   8925
         Begin VB.CheckBox SelecionadoOutro 
            Height          =   195
            Left            =   345
            TabIndex        =   122
            Top             =   450
            Width           =   1020
         End
         Begin VB.CommandButton BotaDesmarcarTodosOutro 
            Caption         =   "Desmarcar Todos"
            Height          =   585
            Left            =   1935
            Picture         =   "OperacaoCaixaECF.ctx":4400
            Style           =   1  'Graphical
            TabIndex        =   112
            Top             =   2250
            Width           =   1440
         End
         Begin VB.CommandButton BotaoMarcarTodosOutro 
            Caption         =   "Marcar Todos"
            Height          =   585
            Left            =   345
            Picture         =   "OperacaoCaixaECF.ctx":55E2
            Style           =   1  'Graphical
            TabIndex        =   111
            Top             =   2250
            Width           =   1440
         End
         Begin VB.ComboBox AdmMeioPagtoOutro 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1380
            TabIndex        =   102
            Text            =   "AdmMeioPagto"
            Top             =   420
            Width           =   1650
         End
         Begin MSMask.MaskEdBox QuantidadeOutro 
            Height          =   300
            Left            =   3060
            TabIndex        =   101
            Top             =   420
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorOutro 
            Height          =   300
            Left            =   4320
            TabIndex        =   103
            Top             =   435
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TotalOutro 
            Height          =   300
            Left            =   5550
            TabIndex        =   104
            Top             =   435
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ClienteOutro 
            Height          =   300
            Left            =   6795
            TabIndex        =   105
            Top             =   435
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridOutros 
            Height          =   1830
            Left            =   270
            TabIndex        =   96
            Top             =   330
            Width           =   8370
            _ExtentX        =   14764
            _ExtentY        =   3228
            _Version        =   393216
            Rows            =   5
            FixedCols       =   0
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
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
            Left            =   6855
            TabIndex        =   114
            Top             =   2370
            Width           =   510
         End
         Begin VB.Label LabelTotalCComprovOutro 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7425
            TabIndex        =   113
            Top             =   2340
            Width           =   1350
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Administradora"
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
            Left            =   1545
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   110
            Top             =   90
            Width           =   1260
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
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
            Left            =   4995
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   109
            Top             =   135
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade"
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
            Left            =   3495
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   108
            Top             =   120
            Width           =   990
         End
         Begin VB.Label Label33 
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
            Height          =   195
            Left            =   5880
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   107
            Top             =   120
            Width           =   450
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
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
            Left            =   7140
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   106
            Top             =   90
            Width           =   600
         End
      End
      Begin MSMask.MaskEdBox ValorSComprovOutro 
         Height          =   285
         Left            =   2340
         TabIndex        =   116
         Top             =   3465
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "Valor sem Comprovante:"
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
         Left            =   210
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   121
         Top             =   3510
         Width           =   2070
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "Total sem Comprovante:"
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
         TabIndex        =   120
         Top             =   3120
         Width           =   2070
      End
      Begin VB.Label LabelSaldoSComprovOutro 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2340
         TabIndex        =   119
         Top             =   3090
         Width           =   1125
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Valor com Comprovante:"
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
         TabIndex        =   118
         Top             =   3510
         Width           =   2085
      End
      Begin VB.Label LabelValorCComprovOutro 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5865
         TabIndex        =   117
         Top             =   3465
         Width           =   1125
      End
   End
   Begin VB.Frame FrameOp 
      BorderStyle     =   0  'None
      Height          =   3825
      Index           =   3
      Left            =   210
      TabIndex        =   16
      Top             =   1560
      Visible         =   0   'False
      Width           =   9150
      Begin VB.CommandButton BotaoSangriaDinheiro 
         Caption         =   "Sangria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   5280
         TabIndex        =   17
         Top             =   2610
         Width           =   3345
      End
      Begin MSMask.MaskEdBox SangriaDinheiro 
         Height          =   645
         Left            =   2865
         TabIndex        =   18
         Top             =   660
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   1138
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "A9"
         PromptChar      =   " "
      End
      Begin VB.Label SaldoDinheiro 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   1
         Left            =   2880
         TabIndex        =   21
         Top             =   1470
         Width           =   3765
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Saldo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1215
         TabIndex        =   20
         Top             =   1470
         Width           =   1470
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1305
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   19
         Top             =   690
         Width           =   1380
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   8175
      ScaleHeight     =   495
      ScaleWidth      =   1140
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   150
      Width           =   1200
      Begin VB.CommandButton BotaoLimpar 
         Height          =   345
         Left            =   120
         Picture         =   "OperacaoCaixaECF.ctx":65FC
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   585
         Picture         =   "OperacaoCaixaECF.ctx":6B2E
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4485
      Left            =   180
      TabIndex        =   47
      Top             =   960
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   7911
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   7
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Abertura / Fechamento / Sessão"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Suprimento"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sangria de Dinheiro"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sangria de Cheques"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sangria de Boletos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sangria de Outros"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Memória Fiscal"
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
   Begin VB.Frame FrameOp 
      BorderStyle     =   0  'None
      Height          =   3825
      Index           =   2
      Left            =   195
      TabIndex        =   40
      Top             =   1605
      Visible         =   0   'False
      Width           =   9150
      Begin VB.CommandButton BotaoSuprimento 
         Caption         =   "Suprimento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   5280
         TabIndex        =   41
         Top             =   2610
         Width           =   3345
      End
      Begin MSMask.MaskEdBox SuprimentoDinheiro 
         Height          =   645
         Left            =   2865
         TabIndex        =   42
         Top             =   660
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   1138
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "A9"
         PromptChar      =   " "
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1305
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   45
         Top             =   690
         Width           =   1380
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Saldo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1215
         TabIndex        =   44
         Top             =   1470
         Width           =   1470
      End
      Begin VB.Label SaldoDinheiro 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   0
         Left            =   2880
         TabIndex        =   43
         Top             =   1470
         Width           =   3765
      End
   End
End
Attribute VB_Name = "OperacaoCaixaECF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()
Dim iFrameAtual As Integer

Private Sub BotaoResetECF_Click()
'Serve para dar reset em caso de erro (Bematech tem)

End Sub


Private Sub BotaoFechamento_Click()

'Quando apertar esse botão faz a transferencia de todos os meios de pagto para o Caixa Central.

End Sub

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

Public Sub Form_Load()
    iFrameAtual = 1
    lErro_Chama_Tela = SUCESSO

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

End Sub

Public Sub Form_Unload(Cancel As Integer)

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Operação de Caixa"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "OperacaoCaixaECF"
    
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

Private Sub TabStrip1_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        FrameOp(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        FrameOp(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index

    End If

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

