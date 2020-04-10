VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.UserControl CobradoresOcx 
   ClientHeight    =   5700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8865
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   8865
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4575
      Index           =   0
      Left            =   210
      TabIndex        =   13
      Top             =   885
      Width           =   8415
      Begin VB.Frame FrameCNAB 
         Caption         =   "Cobrança Eletrônica"
         Height          =   675
         Left            =   165
         TabIndex        =   80
         Top             =   3825
         Width           =   5295
         Begin MSMask.MaskEdBox CNABProxSeqArqCobr 
            Height          =   300
            Left            =   4050
            TabIndex        =   12
            Top             =   240
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   9
            Mask            =   "#########"
            PromptChar      =   " "
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Próximo número do arquivo:"
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
            Left            =   1575
            TabIndex        =   81
            Top             =   270
            Width           =   2355
         End
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   2445
         Picture         =   "CobradoresOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Numeração Automática"
         Top             =   285
         Width           =   300
      End
      Begin VB.ComboBox Filial 
         Height          =   315
         Left            =   4200
         TabIndex        =   10
         Top             =   2985
         Width           =   1440
      End
      Begin VB.Frame Frame5 
         Caption         =   "Tipo do Cobrador"
         Height          =   675
         Left            =   165
         TabIndex        =   15
         Top             =   615
         Width           =   5295
         Begin VB.OptionButton OpcaoOutros 
            Caption         =   "Outros"
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
            Left            =   3915
            TabIndex        =   5
            Top             =   300
            Width           =   960
         End
         Begin VB.OptionButton OpcaoBanco 
            Caption         =   "Banco"
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
            Left            =   135
            TabIndex        =   3
            Top             =   315
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.ComboBox Banco 
            Height          =   315
            Left            =   1065
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   255
            Width           =   2625
         End
      End
      Begin VB.CheckBox CobradorInativo 
         Caption         =   "Inativo"
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
         TabIndex        =   2
         Top             =   315
         Width           =   975
      End
      Begin VB.ListBox Lista_Cobradores 
         Height          =   3960
         Left            =   5775
         Sorted          =   -1  'True
         TabIndex        =   17
         Top             =   555
         Width           =   2415
      End
      Begin VB.TextBox NomeReduzido 
         Height          =   300
         Left            =   1755
         MaxLength       =   20
         TabIndex        =   6
         Top             =   1440
         Width           =   1920
      End
      Begin VB.ComboBox ContaCorrente 
         Height          =   315
         Left            =   1755
         TabIndex        =   8
         Top             =   2460
         Width           =   3120
      End
      Begin VB.CheckBox CobrancaEletronica 
         Caption         =   "Utiliza Cobrança Eletrônica"
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
         Left            =   1470
         TabIndex        =   11
         Top             =   3480
         Width           =   2835
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   300
         Left            =   1845
         TabIndex        =   1
         Top             =   270
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Nome 
         Height          =   300
         Left            =   1755
         TabIndex        =   7
         Top             =   1935
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   50
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Fornecedor 
         Height          =   315
         Left            =   1755
         TabIndex        =   9
         Top             =   2985
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   " Filial:"
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
         Left            =   3600
         TabIndex        =   72
         Top             =   3030
         Width           =   525
      End
      Begin VB.Label LabelFornecedor 
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
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   630
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   71
         Top             =   3030
         Width           =   1035
      End
      Begin VB.Label Label1 
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
         Index           =   0
         Left            =   1125
         TabIndex        =   30
         Top             =   330
         Width           =   660
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Cobradores"
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
         Left            =   5760
         TabIndex        =   31
         Top             =   315
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
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
         Left            =   1110
         TabIndex        =   32
         Top             =   1980
         Width           =   540
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Nome Reduzido:"
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
         Left            =   285
         TabIndex        =   33
         Top             =   1485
         Width           =   1395
      End
      Begin VB.Label ContaCorrenteLabel 
         AutoSize        =   -1  'True
         Caption         =   "Conta corrente:"
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
         Left            =   360
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   34
         Top             =   2520
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4560
      Index           =   1
      Left            =   210
      TabIndex        =   18
      Top             =   885
      Visible         =   0   'False
      Width           =   8415
      Begin VB.Frame FrameCarteira 
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   3435
         Index           =   1
         Left            =   255
         TabIndex        =   41
         Top             =   960
         Width           =   7965
         Begin VB.Frame FrameCarteiraEstatistica 
            Caption         =   "Dados Estatísticos"
            Height          =   2790
            Index           =   3
            Left            =   4440
            TabIndex        =   61
            Top             =   315
            Width           =   3150
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "Qtde pelo Banco:"
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
               TabIndex        =   69
               Top             =   1680
               Width           =   1500
            End
            Begin VB.Label QtdBanco 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   1890
               TabIndex        =   68
               Top             =   1640
               Width           =   615
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Saldo pelo Banco:"
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
               Top             =   2400
               Width           =   1575
            End
            Begin VB.Label SaldoBanco 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   1875
               TabIndex        =   66
               Top             =   2310
               Width           =   975
            End
            Begin VB.Label SaldoValor 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   1890
               TabIndex        =   65
               Top             =   970
               Width           =   975
            End
            Begin VB.Label QtdParcelas 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   1920
               TabIndex        =   64
               Top             =   300
               Width           =   615
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Saldo em Valor:"
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
               TabIndex        =   63
               Top             =   1020
               Width           =   1350
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Qtde de Parcelas:"
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
               Left            =   285
               TabIndex        =   62
               Top             =   345
               Width           =   1545
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Desconto"
            Height          =   1200
            Left            =   525
            TabIndex        =   51
            Top             =   1890
            Width           =   3180
            Begin MSMask.MaskEdBox ContaContabilDup 
               Height          =   315
               Left            =   1545
               TabIndex        =   52
               Top             =   750
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   393216
               AllowPrompt     =   -1  'True
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox TaxaDesconto 
               Height          =   300
               Left            =   720
               TabIndex        =   53
               Top             =   300
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   7
               Format          =   "##0.#0\%"
               PromptChar      =   " "
            End
            Begin VB.Label Label13 
               Caption         =   "Taxa:"
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
               Left            =   165
               TabIndex        =   55
               Top             =   330
               Width           =   585
            End
            Begin VB.Label LabelCtaDesconto 
               AutoSize        =   -1  'True
               Caption         =   "Conta Contábil:"
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
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   54
               Top             =   810
               Width           =   1320
            End
         End
         Begin MSMask.MaskEdBox DiasdeRetencao 
            Height          =   300
            Left            =   2205
            TabIndex        =   56
            Top             =   892
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "9999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TaxaCobranca 
            Height          =   300
            Left            =   2205
            TabIndex        =   57
            Top             =   390
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   7
            Format          =   "##0.#0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Conta 
            Height          =   315
            Left            =   2205
            TabIndex        =   70
            Top             =   1395
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label LabelCtaCarteira 
            AutoSize        =   -1  'True
            Caption         =   "Conta Contábil:"
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
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   60
            Top             =   1455
            Width           =   1320
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Dias de Retenção:"
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
            TabIndex        =   59
            Top             =   945
            Width           =   1605
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Taxa de Cobrança:"
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
            Left            =   495
            TabIndex        =   58
            Top             =   420
            Width           =   1635
         End
      End
      Begin VB.Frame FrameCarteira 
         BorderStyle     =   0  'None
         Caption         =   "Utilizando formulário pré-impresso"
         Height          =   3435
         Index           =   2
         Left            =   180
         TabIndex        =   21
         Top             =   975
         Visible         =   0   'False
         Width           =   8085
         Begin VB.CheckBox FormularioPreImpresso 
            Caption         =   "Utiliza formulário pré-impresso"
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
            Left            =   4800
            TabIndex        =   82
            Top             =   225
            Width           =   3435
         End
         Begin VB.Frame Frame2 
            Caption         =   "Numeração dos Títulos"
            Height          =   1080
            Left            =   150
            TabIndex        =   43
            Top             =   645
            Width           =   7860
            Begin VB.CheckBox GeraNossoNumero 
               Caption         =   "Gerar Automáticamente"
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
               Left            =   105
               TabIndex        =   44
               Top             =   300
               Width           =   4620
            End
            Begin MSMask.MaskEdBox NumeroInicial 
               Height          =   300
               Left            =   780
               TabIndex        =   45
               Top             =   600
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox NumeroFinal 
               Height          =   300
               Left            =   3240
               TabIndex        =   46
               Top             =   600
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox NumeroProx 
               Height          =   300
               Left            =   6000
               TabIndex        =   47
               Top             =   585
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin VB.Label LabelNumProxCartBco 
               AutoSize        =   -1  'True
               Caption         =   "Próximo:"
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
               Left            =   5205
               TabIndex        =   50
               Top             =   660
               Width           =   735
            End
            Begin VB.Label LabelNumFimCartBco 
               AutoSize        =   -1  'True
               Caption         =   "Final:"
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
               Left            =   2745
               TabIndex        =   49
               Top             =   675
               Width           =   480
            End
            Begin VB.Label LabelNumIniCartBco 
               AutoSize        =   -1  'True
               Caption         =   "Inicial:"
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
               Left            =   105
               TabIndex        =   48
               Top             =   660
               Width           =   585
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Identificação da carteira no Banco"
            Height          =   735
            Left            =   150
            TabIndex        =   22
            Top             =   1980
            Width           =   7860
            Begin MSMask.MaskEdBox CodCarteiraNoBanco 
               Height          =   300
               Left            =   870
               TabIndex        =   23
               Top             =   270
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   1
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox NumCarteiraNoBanco 
               Height          =   300
               Left            =   2475
               TabIndex        =   24
               Top             =   285
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "9999"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox NomeNoBanco 
               Height          =   300
               Left            =   3990
               TabIndex        =   25
               Top             =   300
               Width           =   3480
               _ExtentX        =   6138
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   50
               PromptChar      =   " "
            End
            Begin VB.Label LabelNomeCartBco 
               AutoSize        =   -1  'True
               Caption         =   "Nome:"
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
               Left            =   3345
               TabIndex        =   35
               Top             =   315
               Width           =   555
            End
            Begin VB.Label LabelNumCartBco 
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
               Height          =   195
               Left            =   1680
               TabIndex        =   36
               Top             =   330
               Width           =   720
            End
            Begin VB.Label LabelCodCartBco 
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
               Height          =   195
               Left            =   135
               TabIndex        =   37
               Top             =   330
               Width           =   660
            End
         End
         Begin VB.CheckBox Registrada 
            Caption         =   "Com Registro"
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
            Left            =   270
            TabIndex        =   26
            Top             =   225
            Width           =   1605
         End
         Begin VB.CheckBox BcoImprimeBoleta 
            Caption         =   "Banco Imprime a Boleta"
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
            Left            =   2205
            TabIndex        =   27
            Top             =   225
            Width           =   2595
         End
      End
      Begin VB.CommandButton BotaoRemoverCarteira 
         Caption         =   "Remover"
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
         Left            =   7065
         TabIndex        =   14
         Top             =   30
         Width           =   1260
      End
      Begin VB.CommandButton BotaoAdicionarCarteira 
         Caption         =   "Adicionar..."
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
         Left            =   5565
         TabIndex        =   16
         Top             =   30
         Width           =   1395
      End
      Begin VB.CheckBox Desativada 
         Caption         =   "Desativada"
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
         Left            =   4095
         TabIndex        =   28
         Top             =   105
         Width           =   1320
      End
      Begin VB.ComboBox Carteira 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   60
         Width           =   2775
      End
      Begin MSComctlLib.TabStrip TabStripCartCobr 
         Height          =   3960
         Left            =   15
         TabIndex        =   42
         Top             =   585
         Width           =   8325
         _ExtentX        =   14684
         _ExtentY        =   6985
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Dados Básicos"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Cobrança Eletrônica"
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
      Begin VB.Label Label3 
         Caption         =   "Carteira:"
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
         Left            =   315
         TabIndex        =   38
         Top             =   90
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4545
      Index           =   2
      Left            =   195
      TabIndex        =   19
      Top             =   765
      Visible         =   0   'False
      Width           =   8415
      Begin TelasCpr.TabEndereco TabEnd 
         Height          =   3660
         Index           =   0
         Left            =   15
         TabIndex        =   79
         Top             =   1035
         Width           =   8445
         _ExtentX        =   14896
         _ExtentY        =   6456
      End
      Begin VB.Frame SSFrame1 
         Height          =   510
         Left            =   30
         TabIndex        =   20
         Top             =   345
         Width           =   8355
         Begin VB.Label Label15 
            Caption         =   "Cobrador:"
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
            Left            =   180
            TabIndex        =   39
            Top             =   180
            Width           =   870
         End
         Begin VB.Label Cobrador 
            Height          =   210
            Index           =   1
            Left            =   1065
            TabIndex        =   40
            Top             =   180
            Width           =   3300
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6570
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   75
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "CobradoresOcx.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "CobradoresOcx.ctx":0244
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "CobradoresOcx.ctx":03CE
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "CobradoresOcx.ctx":0900
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5220
      Left            =   90
      TabIndex        =   0
      Top             =   420
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   9208
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Carteiras"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Endereço"
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
Attribute VB_Name = "CobradoresOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'observacao: a colecao de carteiras de cobrador gcolCarteirasCobrador tem como chave o cstr do codigo da carteira e seus elementos podem nao estar na mesma ordem dos elementos na combo que contem no itemdata o codigo da carteira de cobranca

'Property Variables:
Dim m_Caption As String
Event Unload()

Public gobjTabEnd As New ClassTabEndereco

'DECLARAÇÃO DAS VARIÁVEIS GLOBAIS
Public iAlterado As Integer
Dim iFrameAtual As Integer
Dim iFrameAtualCart As Integer
Dim iCarteiraAnterior As Integer
Dim gcolCarteirasCobrador As Collection
Dim iFornecedorAlterado As Integer

Private WithEvents objEventoContaCorrenteInt As AdmEvento
Attribute objEventoContaCorrenteInt.VB_VarHelpID = -1
Private WithEvents objEventoCarteiraCobrador As AdmEvento
Attribute objEventoCarteiraCobrador.VB_VarHelpID = -1
Private WithEvents objEventoCtaCarteira As AdmEvento
Attribute objEventoCtaCarteira.VB_VarHelpID = -1
Private WithEvents objEventoCtaDesconto As AdmEvento
Attribute objEventoCtaDesconto.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_Carteiras = 2
Private Const TAB_Endereco = 3

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    'Gera Código da proximo cobrador
    lErro = CF("Cobrador_Automatico", iCodigo)
    If lErro <> SUCESSO Then Error 57545
    
    Codigo.PromptInclude = False
    Codigo.Text = CStr(iCodigo)
    Codigo.PromptInclude = True

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57545
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154282)
    
    End Select

    Exit Sub

End Sub

Private Sub Banco_Click()

Dim iPosicao As Integer

    iAlterado = REGISTRO_ALTERADO

    If Banco.ListIndex <> -1 Then
        'Coloca o Nome Reduzido do Cobrador na Label
        iPosicao = InStr(Banco.Text, SEPARADOR) + 1
        NomeReduzido.Text = Trim(Mid(Banco.Text, iPosicao))
    End If
    
End Sub

Private Sub BotaoAdicionarCarteira_Click()

Dim objCarteiraCobranca As New ClassCarteiraCobranca
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoAdicionarCarteira_Click

    If Len(Trim(Codigo.Text)) = 0 Then Error 62110
    
    If OpcaoBanco.Value Then
        colSelecao.Add 0
    Else
        colSelecao.Add 1
    End If

    Call Chama_Tela("CarteirasCobrancaLista", colSelecao, objCarteiraCobranca, objEventoCarteiraCobrador)

    Exit Sub
    
Erro_BotaoAdicionarCarteira_Click:

    Select Case Err
        
        Case 62110
            Call Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_INFORMADO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154283)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objCobrador As New ClassCobrador
Dim colCodNomeFiliais As New AdmColCodigoNome
Dim iCodigo As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o codigo foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then Error 23502

    objCobrador.iCodigo = CInt(Codigo.Text)

    'Lê os dados do Cobrador a ser excluido
    lErro = CF("Cobrador_Le", objCobrador)
    If lErro <> SUCESSO And lErro <> 19294 Then Error 23503

    'Verifica se Cobrador está cadastrado
    If lErro <> SUCESSO Then Error 23504

    'Envia aviso perguntando se realmente deseja excluir Cobrador
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_COBRADOR", objCobrador.iCodigo)

    If vbMsgRes = vbYes Then

        'Exclui Cobrador
        lErro = CF("Cobrador_Exclui", objCobrador)
        If lErro <> SUCESSO Then Error 23505

        'Exclui da ListBox
        Call Lista_Cobradores_Remove(objCobrador)

        'Limpa a Tela
        Call Limpa_Tela_Cobrador

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 23502
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_INFORMADO", Err)

        Case 23503, 23505, 23506

        Case 23504   'Cobrador com codigo %i nao esta cadastrada
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_CADASTRADO", Err, objCobrador.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154284)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoGravar_Click

    'Grava o Cobrador
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 23463

    'Limpa a Tela
    Call Limpa_Tela_Cobrador

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 23463, 23464

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154285)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 80459

    'Limpa a Tela
    Call Limpa_Tela_Cobrador

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 80459, 23491

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154286)

    End Select

End Sub

Private Sub BotaoRemoverCarteira_Click()

Dim lErro As Long
Dim objCarteiraCobrador As New ClassCarteiraCobrador
Dim vbMsgRes  As VbMsgBoxResult

On Error GoTo Erro_BotaoRemoverCarteira_Click

    If Carteira.ListIndex = -1 Then Exit Sub

    If Len(Trim(Codigo.Text)) > 0 Then
        objCarteiraCobrador.iCobrador = CInt(Codigo.Text)
    Else
        Error 19414
    End If

    objCarteiraCobrador.iCodCarteiraCobranca = Codigo_Extrai(Carteira.Text)

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_CARTEIRA", objCarteiraCobrador.iCodCarteiraCobranca)
    
    If vbMsgRes = vbNo Then Error 62105

    lErro = CF("Cobrador_Exclui_Carteira", objCarteiraCobrador)
    If lErro <> SUCESSO And lErro <> 23524 Then Error 23461

    'Excluir da colecao
    gcolCarteirasCobrador.Remove CStr(objCarteiraCobrador.iCodCarteiraCobranca)

    'Excluir da combobox carteira
    Carteira.RemoveItem (Carteira.ListIndex)

    iCarteiraAnterior = -1
    
    'Desseleciona Combo
    Carteira.ListIndex = -1

    'Limpar o frame do tab de carteiras
    lErro = Limpa_Tab_Carteiras
    If lErro <> SUCESSO Then Error 23462

    Exit Sub

Erro_BotaoRemoverCarteira_Click:

    Select Case Err

        Case 19414
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_INFORMADO", Err)
        
        Case 23459

        Case 23461, 23462, 62105

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154287)

    End Select

    Exit Sub

End Sub

Private Sub Carteira_Click()

Dim lErro As Long
Dim iCodigo As Integer
Dim objCarteiraCobrador As New ClassCarteiraCobrador

On Error GoTo Erro_Carteira_Click

    'Verifica se há algum item selecionado na combo
    If Carteira.ListIndex = -1 Then
        iCarteiraAnterior = -1
        Exit Sub
    End If

    'Atualiza carteira atual
    If iCarteiraAnterior <> -1 And Carteira.List(iCarteiraAnterior) <> Carteira.List(Carteira.ListIndex) Then
    
        'transfere dados da tela p/obj
        lErro = Move_Carteira_Memoria(objCarteiraCobrador)
        If lErro <> SUCESSO Then Error 23578
        
        objCarteiraCobrador.iCodCarteiraCobranca = Carteira.ItemData(iCarteiraAnterior)
        
        'Atualiza carteira na coleção
        gcolCarteirasCobrador.Remove CStr(Carteira.ItemData(iCarteiraAnterior))
        gcolCarteirasCobrador.Add objCarteiraCobrador, CStr(Carteira.ItemData(iCarteiraAnterior))
    
        iCarteiraAnterior = -1
        
    End If

    'Pega dados da carteira selecionada e joga na tela
    For Each objCarteiraCobrador In gcolCarteirasCobrador
        
        iCodigo = objCarteiraCobrador.iCodCarteiraCobranca

        If Carteira.ItemData(Carteira.ListIndex) = iCodigo Then
        
            lErro = Exibe_CarteiraCobrador(objCarteiraCobrador)
            If lErro <> SUCESSO Then Error 23561

            Exit For
            
        End If
        
    Next

    iCarteiraAnterior = Carteira.ListIndex

    Exit Sub

Erro_Carteira_Click:

    Select Case Err

        Case 23561, 23578

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154288)

    End Select

    Exit Sub

End Sub

Private Sub CobrancaEletronica_Click()

Dim bHabilita As Boolean
    
    If CobrancaEletronica.Value = vbChecked Then
        bHabilita = True
    Else
        bHabilita = False
    End If
    
    Call Habilita_Campos_CobrancaEletronica(bHabilita)
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long, iTrocouCobrador As Integer
Dim objCobrador As New ClassCobrador
Dim objCarteiraCobrador As ClassCarteiraCobrador

On Error GoTo Erro_Codigo_Validate
        
    'Verifica se o Código foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then Exit Sub

    'Critica se é do tipo inteiro positivo
    lErro = Inteiro_Critica(Codigo.Text)
    If lErro <> SUCESSO Then Error 23542
    
    'Preenche o código no objeto
    objCobrador.iCodigo = Codigo.Text
    
    If objCobrador.iCodigo = COBRADOR_PROPRIA_EMPRESA Then Error 43369
    
    'Se o codigo do cobrador na colecao é diferente,
    'limpá-la e ao tab de carteiras
    iTrocouCobrador = 0
    For Each objCarteiraCobrador In gcolCarteirasCobrador
    
        If objCarteiraCobrador.iCobrador <> objCobrador.iCodigo Then
            iTrocouCobrador = 1
            Exit For
        End If
        
    Next
    
    If iTrocouCobrador Then
    
        Set gcolCarteirasCobrador = New Collection
        Carteira.Clear
        Call Limpa_Tab_Carteiras
    
    End If
    
    'Lê o Cobrador no BD , a consulta sera o codigo
    lErro = CF("Cobrador_Le", objCobrador)
    If lErro <> SUCESSO And lErro <> 19294 Then Error 40588
   
    If lErro = SUCESSO Then
                                         
        If giFilialEmpresa <> EMPRESA_TODA Then
        
            If objCobrador.iFilialEmpresa <> giFilialEmpresa Then Error 49555
        
        End If
        
        'Exibe os dados do Cobrador
        lErro = Exibe_Dados_Cobrador(objCobrador)
        If lErro <> SUCESSO Then Error 40590
    
        'Fecha o comando das setas se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)
        
        'Zerar iAlterado
        iAlterado = 0

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True


    Select Case Err

        Case 23542
            
        Case 40588, 40590
        
        Case 43369
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_PROPRIA_EMPRESA", Err, objCobrador.iCodigo)
            
        Case 49555
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_PERTENCE_FILIAL", Err, objCobrador.iCodigo, giFilialEmpresa)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154289)

    End Select

    Exit Sub

End Sub

Private Sub Conta_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Conta_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaMascarada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Conta_Validate
    
    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica_Modulo", Conta.Text, Conta.ClipText, objPlanoConta, MODULO_CONTASARECEBER)
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then Error 19371
    
    If lErro = SUCESSO Then
    
        sContaFormatada = objPlanoConta.sConta
        
        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)
        
        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 19372
        
        Conta.PromptInclude = False
        Conta.Text = sContaMascarada
        Conta.PromptInclude = True
    
    'se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then
  
        lErro = CF("Conta_Critica", Conta.Text, sContaFormatada, objPlanoConta, MODULO_CONTASARECEBER)
        If lErro <> SUCESSO And lErro <> 5700 Then Error 41528
    
        If lErro = 5700 Then Error 41529
    
    End If
    
    Exit Sub

Erro_Conta_Validate:

    Cancel = True


    Select Case Err

    Case 19371, 41528
    
    Case 19372
        lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)

    Case 41529
        lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", Err, Conta.Text)

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154290)

    End Select

    Exit Sub

End Sub

Private Sub ContaContabilDup_Validate(Cancel As Boolean)
Dim lErro As Long
Dim sContaFormatada As String
Dim sContaMascarada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ContaContabilDup_Validate
    
    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica_Modulo", ContaContabilDup.Text, ContaContabilDup.ClipText, objPlanoConta, MODULO_CONTASARECEBER)
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then Error 39788
    
    If lErro = SUCESSO Then
    
        sContaFormatada = objPlanoConta.sConta
        
        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)
        
        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 39796
        
        ContaContabilDup.PromptInclude = False
        ContaContabilDup.Text = sContaMascarada
        ContaContabilDup.PromptInclude = True
    
    
    'se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then
    
        lErro = CF("Conta_Critica", ContaContabilDup.Text, sContaFormatada, objPlanoConta, MODULO_CONTASARECEBER)
        If lErro <> SUCESSO And lErro <> 5700 Then Error 41530
    
        If lErro = 5700 Then Error 41531
    
    End If
    
    Exit Sub

Erro_ContaContabilDup_Validate:

    Cancel = True


    Select Case Err
    
    Case 39796
        lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)
        
    Case 39788, 41530

    Case 41531
        lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", Err, ContaContabilDup.Text)

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154291)

    End Select

    Exit Sub

End Sub

Private Sub ContaCorrente_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaCorrente_Validate(bCancel As Boolean)

Dim lErro As Long
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_ContaCorrente_Validate

    If Len(Trim(ContaCorrente.Text)) = 0 Then Exit Sub

    'Verifica se esta preenchida com o item selecionado na ComboBox CodContacOrrente
    If ContaCorrente.Text = ContaCorrente.List(ContaCorrente.ListIndex) Then Exit Sub
    
    lErro = Combo_Seleciona(ContaCorrente, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 23553
    
    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then
        
        objContaCorrenteInt.iCodigo = iCodigo
        
        lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
        If lErro <> SUCESSO And lErro <> 11807 Then Error 23554
               
        'Não encontrou a Conta Corrente no BD
        If lErro = 11807 Then Error 23555
        
        If giFilialEmpresa <> EMPRESA_TODA Then
        
            If objContaCorrenteInt.iFilialEmpresa <> giFilialEmpresa Then Error 49551
        
        End If
        
        'Encontrou a Conta Corrente no BD, coloca no Text da Combo
        ContaCorrente.Text = CStr(objContaCorrenteInt.iCodigo) & SEPARADOR & objContaCorrenteInt.sNomeReduzido

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 23556

    Exit Sub

Erro_ContaCorrente_Validate:

    bCancel = True
    
    Select Case Err

        Case 23553, 23554

        Case 23555
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CODCONTACORRENTE_INEXISTENTE", objContaCorrenteInt.iCodigo)

            If vbMsgRes = vbYes Then
                'Chama a tela de Contas Correntes
                Call Chama_Tela("CtaCorrenteInt", objContaCorrenteInt)
            End If

        Case 23556
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE1", Err, ContaCorrente.Text)
        
        Case 49551
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_PERTENCE_FILIAL", Err, ContaCorrente.Text, giFilialEmpresa)
             
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154292)

    End Select

    Exit Sub

End Sub

Private Sub Desativada_Click()

    iAlterado = REGISTRO_ALTERADO
     

End Sub

Private Sub DiasdeRetencao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DiasdeRetencao_GotFocus()

    Call MaskEdBox_TrataGotFocus(DiasdeRetencao, iAlterado)
    
End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome
Dim colCodigo As New Collection
Dim ColCobrador As New Collection
Dim objCobrador As ClassCobrador
Dim vCodigo As Variant, iIndice2 As Integer
Dim sListBoxItem As String
Dim objTela As Object

On Error GoTo Erro_Cobradores_Form_Load
    
    iFrameAtual = 1
    iFrameAtualCart = 1

    Set gcolCarteirasCobrador = New Collection
    Set objEventoCarteiraCobrador = New AdmEvento
    Set objEventoContaCorrenteInt = New AdmEvento
    Set objEventoFornecedor = New AdmEvento
    
    'Preenche a ComboBox banco
    'Lê cada código e Nome Reduzido da tabela bancos
    lErro = CF("Cod_Nomes_Le", "Bancos", "CodBanco", "NomeReduzido", STRING_NOME_REDUZIDO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 23449

    For Each objCodigoDescricao In colCodigoDescricao

        'Concatena Código e nome reduzido do banco
        sListBoxItem = CStr(objCodigoDescricao.iCodigo)
        sListBoxItem = sListBoxItem & SEPARADOR & Trim(objCodigoDescricao.sNome)

        Banco.AddItem sListBoxItem
        Banco.ItemData(Banco.NewIndex) = objCodigoDescricao.iCodigo

    Next
    
    'Carrega a Combo de Contas
    lErro = Carrega_ContasCorrente()
    If lErro <> SUCESSO Then Error 43447

    'Preenche a listbox de Cobradores
    Set colCodigoDescricao = New AdmColCodigoNome

    'Lê cada código e Nome Reduzido da tabela cobradores
    lErro = CF("Cobradores_Le_Todos_Filial", ColCobrador)
    If lErro <> SUCESSO Then Error 23451

    For Each objCobrador In ColCobrador
    
        If objCobrador.iCodigo <> COBRADOR_PROPRIA_EMPRESA Then
            Lista_Cobradores.AddItem objCobrador.sNomeReduzido
            Lista_Cobradores.ItemData(Lista_Cobradores.NewIndex) = objCobrador.iCodigo
        End If
    Next

    'Lê cada código da tabela Estados
    Set objTela = Me
    lErro = gobjTabEnd.Inicializa(objTela, TabEnd(0))
    If lErro <> SUCESSO Then Error 23452

    'Inicializa mascaras de contas
    lErro = Inicializa_Contabilidade_Tela()
    If lErro <> SUCESSO Then Error 40595

    Call Habilita_Campos_CobrancaEletronica(False)

    iAlterado = 0
    iCarteiraAnterior = -1
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Cobradores_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 23449, 23451, 23452, 23453, 40595, 40596, 43447

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154294)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Carrega_ContasCorrente() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim colCodigoNomeRed As New AdmColCodigoNome
Dim objCodigoNome As New AdmCodigoNome

On Error GoTo Erro_Carrega_ContasCorrente

    'Leitura dos códigos e descrições das Contas no BD
    lErro = CF("ContasCorrentesInternas_Le_CodigosNomesRed", colCodigoNomeRed)
    If lErro <> SUCESSO Then Error 43446

    'Preenche listbox com descrição das contas
    For iIndice = 1 To colCodigoNomeRed.Count
        
        Set objCodigoNome = colCodigoNomeRed(iIndice)
        
        ContaCorrente.AddItem objCodigoNome.iCodigo & SEPARADOR & objCodigoNome.sNome
        ContaCorrente.ItemData(ContaCorrente.NewIndex) = objCodigoNome.iCodigo

    Next

    Carrega_ContasCorrente = SUCESSO

    Exit Function

Erro_Carrega_ContasCorrente:

    Carrega_ContasCorrente = Err

    Select Case Err

        Case 43446

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154295)

    End Select

    Exit Function
    
End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_UnLoad(Cancel As Integer)
Dim lErro As Long

    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
 
    Set objEventoCarteiraCobrador = Nothing
    Set objEventoContaCorrenteInt = Nothing
    Set objEventoCtaCarteira = Nothing
    Set objEventoCtaDesconto = Nothing
    Set gcolCarteirasCobrador = Nothing
    Set objEventoFornecedor = Nothing
    
    Call gobjTabEnd.Finaliza
    Set gobjTabEnd = Nothing
    
End Sub

Private Sub CobradorInativo_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaCorrenteLabel_Click()
'chama browse de conta corrente

Dim objContasCorrentesInternas As New ClassContasCorrentesInternas
Dim colSelecao As New Collection

    If Len(Trim(ContaCorrente.Text)) > 0 Then objContasCorrentesInternas.iCodigo = Codigo_Extrai(ContaCorrente.Text)

    Call Chama_Tela("CtaCorrenteLista", colSelecao, objContasCorrentesInternas, objEventoContaCorrenteInt)

End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.ListIndex >= 0 Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 111366

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        'Verifica se o Fornecedor foi digitado
        If Len(Trim(Fornecedor.Text)) = 0 Then Error 111367

        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then Error 111368

        If lErro = 18272 Then

            objFornecedor.sNomeReduzido = Fornecedor.Text

            'Le o Código do Fornecedor --> Para Passar para a Tela de Filiais
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then Error 111369

            'Passa o Código do Fornecedor
            objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo

            'Sugere cadastrar nova Filial
            Error 111370

        End If

        'Coloca na tela
        Filial.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then Error 111371

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 111366, 111368, 111369 'Tratados nas Rotinas chamadas

        Case 111370
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 111367
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 111371
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", Err, Filial.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154296)

    End Select

    Exit Sub

End Sub

Private Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Filial_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Fornecedor_Change()

    iAlterado = REGISTRO_ALTERADO
    iFornecedorAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)
'Função que valida os dados do Fornecedor

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Validate

    'Verifica se o Fornecedor foi Alterado
    If iFornecedorAlterado = REGISTRO_ALTERADO Then

        'Verifica se o Fornecedor está preenchido
        If Len(Trim(Fornecedor.Text)) <> 0 Then
            'Função que Lê o Fornecedor Pelo Nome/Codigo/CPF/CGC , iCodFilial Traz o Codigo da Filial Matriz
            lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
            If lErro <> SUCESSO Then gError 111350
            'obs se não Encontrar o fornecedor pelo nome erro 6663
            'obs se não Encontrar o fornecedor pelo Codigo erro 6664
            'obs se não Encontrar o fornecedor pelo Codigo erro 6675
            'obs se não Encontrar o fornecedor pelo Codigo erro 6660
            
            Fornecedor.Text = objFornecedor.sNomeReduzido

            'Função que a Coleção de nomes e Codidos das Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then gError 111351

            'Função que Preenche a Combo de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)

            'Função que Seleciona a Filial na Combo de Filiais
            Call CF("Filial_Seleciona", Filial, iCodFilial)

         Else

            'Se o Fornecedor não estiver preenchido limpa a Combo de Filial
            Filial.Clear

        End If

    End If

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True

    Select Case gErr

        Case 111350, 111351

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154297)

    End Select

    Exit Sub


End Sub

Private Sub GeraNossoNumero_Click()

    If GeraNossoNumero.Value = vbChecked Then
    
        LabelNumIniCartBco.Enabled = False
        NumeroInicial.Enabled = False
        LabelNumFimCartBco.Enabled = False
        NumeroFinal.Enabled = False
        LabelNumProxCartBco.Enabled = False
        NumeroProx.Enabled = False
    Else
        LabelNumIniCartBco.Enabled = True
        NumeroInicial.Enabled = True
        LabelNumFimCartBco.Enabled = True
        NumeroFinal.Enabled = True
        LabelNumProxCartBco.Enabled = True
        NumeroProx.Enabled = True
    End If

    Exit Sub

End Sub

Private Sub LabelCtaCarteira_Click()
'chama browse do plano de contas

Dim lErro As Long
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim iContaPreenchida As Integer
Dim sConta As String

On Error GoTo Erro_LabelCtaCarteira_Click

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", Conta.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 43358

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    Call Chama_Tela("PlanoContaCRLista", colSelecao, objPlanoConta, objEventoCtaCarteira)

    Exit Sub

Erro_LabelCtaCarteira_Click:

    Select Case Err

        Case 43358
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154298)

    End Select

    Exit Sub

End Sub

Private Sub LabelCtaDesconto_Click()
'chama browse do plano de contas

Dim lErro As Long
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim iContaPreenchida As Integer
Dim sConta As String

On Error GoTo Erro_LabelCtaDesconto_Click

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaContabilDup.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 43360

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    Call Chama_Tela("PlanoContaCRLista", colSelecao, objPlanoConta, objEventoCtaDesconto)

    Exit Sub

Erro_LabelCtaDesconto_Click:

    Select Case Err

        Case 43360
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154299)

    End Select

    Exit Sub

End Sub

Private Sub LabelFornecedor_Click()
'Função que Traz o Browse de Fornecedor para a Tela

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim colSelecao As New Collection

On Error GoTo Erro_LabelFornecedor_Click

    'Verifica se fornecedor esta preenchido
    If Len(Trim(Fornecedor.Text)) <> 0 Then
        objFornecedor = Fornecedor.Text
    End If

    'Chama o Browse de Fornecedor
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)
  
    Exit Sub

Erro_LabelFornecedor_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154300)

    End Select

    Exit Sub

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)
'Traz os Fornecedores

Dim objFornecedor As ClassFornecedor
Dim lErro As Long
Dim bCancel As Boolean

On Error GoTo Erro_objEventoFornecedor_evSelecao

    Set objFornecedor = obj1

    'Move os Dados do fornecedor para a tela
    Fornecedor.Text = objFornecedor.sNomeReduzido

    'Chama a Função que valida o nome
    Call Fornecedor_Validate(bCancel)

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoFornecedor_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154301)

    End Select

    Exit Sub

End Sub

Private Sub Lista_Cobradores_DblClick()

Dim lErro As Long
Dim sListBoxItem As String
Dim objCobrador As New ClassCobrador

On Error GoTo Erro_Lista_Cobradores_DblClick

    'Verifica se há algum Cobrador Selecionada
    If Lista_Cobradores.ListIndex = -1 Then Exit Sub
    
    'Guarda o valor do código do Cobrador selecionado na ListBox Lista_Cobradores
    objCobrador.iCodigo = Lista_Cobradores.ItemData(Lista_Cobradores.ListIndex)

    'Lê o Cobrador no BD
    lErro = CF("Cobrador_Le", objCobrador)
    If lErro <> SUCESSO And lErro <> 19250 Then Error 23539

    'Se Cobrador não está cadastrado, erro
    If lErro <> SUCESSO Then Error 23540

    'Exibe os dados do Cobrador
    lErro = Exibe_Dados_Cobrador(objCobrador)
    If lErro <> SUCESSO Then Error 23541

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Exit Sub

Erro_Lista_Cobradores_DblClick:

    Select Case Err

    Case 23539, 23541

    Case 23540
        lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_CADASTRADO", Err, objCobrador.iCodigo)
        Lista_Cobradores.RemoveItem (Lista_Cobradores.ListIndex)

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154302)

    End Select

    Exit Sub

End Sub

Private Sub ContaContabilDup_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Nome_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NomeNoBanco_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim iIndice As Integer

On Error GoTo Erro_NomeNoBanco_Validate

    'verifica se foi preenchido o campo Codigo Fornecedor
    If Len(Trim(NumCarteiraNoBanco.Text)) = 0 Then Exit Sub

    lErro = Inteiro_Critica(NumCarteiraNoBanco.Text)
    If lErro <> SUCESSO Then Error 49550
    
    Exit Sub

Erro_NomeNoBanco_Validate:

    Cancel = True


    Select Case Err

    Case 49550

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154303)

    End Select

    Exit Sub

End Sub

Private Sub NomeReduzido_Change()

    iAlterado = REGISTRO_ALTERADO
    Cobrador(1).Caption = NomeReduzido.Text

End Sub

Private Sub NomeReduzido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NomeReduzido_Validate
    
    'Se está preenchido, testa se começa por letra
    If Len(Trim(NomeReduzido.Text)) > 0 Then

        If Not IniciaLetra(NomeReduzido.Text) Then Error 57825

    End If
        
    Exit Sub

Erro_NomeReduzido_Validate:

    Cancel = True

    
    Select Case Err
    
        Case 57825
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_COMECA_LETRA", Err, NomeReduzido.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154304)
    
    End Select
    
    Exit Sub

End Sub

Private Sub NumCarteiraNoBanco_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumCarteiraNoBanco, iAlterado)

End Sub

Private Sub NumeroFinal_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub NumeroInicial_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub NumeroProx_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoCarteiraCobrador_evSelecao(obj1 As Object)

Dim objCarteiraCobranca As New ClassCarteiraCobranca
Dim objCarteiraCobrador As New ClassCarteiraCobrador
Dim iIndice As Integer
Dim lErro As Long
Dim sListBoxItem As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_objEventoCarteiraCobrador_evSelecao

    Set objCarteiraCobranca = obj1
    
    'Concatena Código e a Descricao da carteira
    sListBoxItem = CStr(objCarteiraCobranca.iCodigo)
    sListBoxItem = sListBoxItem & SEPARADOR & objCarteiraCobranca.sDescricao
 
    'Verifica se a Carteira já está na combo
    For iIndice = 0 To Carteira.ListCount - 1
        
        If Carteira.ItemData(iIndice) = objCarteiraCobranca.iCodigo Then Error 56600
        
    Next
         
    'Preenche Text da Combo Carteira
    Carteira.AddItem sListBoxItem
    
    Carteira.ItemData(Carteira.NewIndex) = objCarteiraCobranca.iCodigo
    
    objCarteiraCobrador.iCodCarteiraCobranca = objCarteiraCobranca.iCodigo
    objCarteiraCobrador.iCobrador = CInt(Codigo.Text)
    
    'Incluir na Coleção
    gcolCarteirasCobrador.Add objCarteiraCobrador, CStr(objCarteiraCobranca.iCodigo)

    'Selecionar Carteira
    Carteira.ListIndex = Carteira.ListCount - 1
    
    Me.Show

    Exit Sub

Erro_objEventoCarteiraCobrador_evSelecao:

    Select Case Err
   
        Case 56600
            vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_CARTEIRA_JA_ADICIONADA", sListBoxItem)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154305)

    End Select

    Exit Sub

End Sub

Private Sub objEventoContaCorrenteInt_evSelecao(obj1 As Object)

Dim objContaCorrenteInt As ClassContasCorrentesInternas, bCancel As Boolean

    Set objContaCorrenteInt = obj1
    
    ContaCorrente.Text = objContaCorrenteInt.iCodigo
    Call ContaCorrente_Validate(bCancel)
    
    Me.Show

End Sub

Private Sub objEventoCtaCarteira_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoCtaCarteira_evSelecao

    Set objPlanoConta = obj1

    If objPlanoConta.sConta = "" Then
        Conta.Text = ""
    Else
        Conta.PromptInclude = False

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then Error 43359

        Conta.Text = sContaEnxuta

        Conta.PromptInclude = True

    End If

    Me.Show

    Exit Sub

Erro_objEventoCtaCarteira_evSelecao:

    Select Case Err

        Case 43359
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154306)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCtaDesconto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoCtaDesconto_evSelecao

    Set objPlanoConta = obj1

    If objPlanoConta.sConta = "" Then
        ContaContabilDup.Text = ""
    Else
        ContaContabilDup.PromptInclude = False

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then Error 43361

        ContaContabilDup.Text = sContaEnxuta

        ContaContabilDup.PromptInclude = True

    End If

    Me.Show

    Exit Sub

Erro_objEventoCtaDesconto_evSelecao:

    Select Case Err

        Case 43361
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154307)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub Opcao_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then
        
        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        Frame1(Opcao.SelectedItem.Index - 1).Visible = True
        Frame1(iFrameAtual - 1).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index
        
        Select Case iFrameAtual
        
            Case TAB_Identificacao
                Parent.HelpContextID = IDH_COBRADORES_ID
                
            Case TAB_Carteiras
                Parent.HelpContextID = IDH_COBRADORES_CARTEIRAS
                        
            Case TAB_Endereco
                Parent.HelpContextID = IDH_COBRADORES_ENDERECOS
            
        End Select
    
    End If

End Sub

Function Trata_Parametros(Optional objCobrador As ClassCobrador) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se há um Cobrador preenchido
    If Not (objCobrador Is Nothing) Then

        'Verifica se o Cobrador existe, lendo no BD a partir do  codigo
        lErro = CF("Cobrador_Le", objCobrador)
        If lErro <> SUCESSO And lErro <> 19294 Then Error 23454

        'Se o Cobrador existe
        If lErro = SUCESSO Then
            
            If giFilialEmpresa <> EMPRESA_TODA Then
        
                If objCobrador.iFilialEmpresa <> giFilialEmpresa Then Error 49557
        
            End If
                   
            lErro = Exibe_Dados_Cobrador(objCobrador)

        'Se o Cobrador não existe
        Else

            'Mantém o Código do Cobrador na tela
            Codigo.PromptInclude = False
            Codigo.Text = CStr(objCobrador.iCodigo)
            Codigo.PromptInclude = True

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 23454, 23455
        
        Case 49557
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_PERTENCE_FILIAL", Err, objCobrador.iCodigo, giFilialEmpresa)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154308)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Sub Habilita_Campos_Banco(bHabilita As Boolean)

    CobrancaEletronica.Enabled = bHabilita
    Call Habilita_Campos_CobrancaEletronica(bHabilita)
    
End Sub

Private Sub OpcaoBANCO_Click()

    iAlterado = REGISTRO_ALTERADO

    'Habilitar campos específicos de Bancos
    CobrancaEletronica.Enabled = True
    
End Sub

Private Sub OpcaoOutros_Click()

    iAlterado = REGISTRO_ALTERADO

    CobrancaEletronica.Value = vbUnchecked
    
    'Desabilitar campos específicos de Bancos
    Call Habilita_Campos_Banco(False)

End Sub

Private Sub TabStripCartCobr_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStripCartCobr.SelectedItem.Index <> iFrameAtualCart Then
        
        If TabStrip_PodeTrocarTab(iFrameAtualCart, TabStripCartCobr, Me) <> SUCESSO Then Exit Sub

        FrameCarteira(TabStripCartCobr.SelectedItem.Index).Visible = True
        FrameCarteira(iFrameAtualCart).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtualCart = TabStripCartCobr.SelectedItem.Index
        
    End If

End Sub

Private Sub TaxaCobranca_Change()

    iAlterado = REGISTRO_ALTERADO
   
End Sub

Private Sub TaxaCobranca_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TaxaCobranca_Validate

    If Len(Trim(TaxaCobranca.Text)) = 0 Then Exit Sub

    lErro = Porcentagem_Critica(TaxaCobranca.Text)
    
    If lErro = SUCESSO Then
        
        TaxaCobranca.Text = Format(TaxaCobranca.Text, "Fixed")
    
    Else
        
        Error 40622
    
    End If
    
    Exit Sub
    
Erro_TaxaCobranca_Validate:

    Cancel = True

    
    Select Case Err

        Case 40622
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154310)
    
    End Select

    Exit Sub

End Sub

Private Sub TaxaDesconto_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub TaxaDesconto_Validate(Cancel As Boolean)
Dim lErro As Long

On Error GoTo Erro_TaxaDesconto_Validate

    If Len(Trim(TaxaDesconto.Text)) = 0 Then Exit Sub

    lErro = Porcentagem_Critica(TaxaDesconto.Text)
    
    If lErro = SUCESSO Then
        
        TaxaDesconto.Text = Format(TaxaDesconto.Text, "Fixed")
    
    Else
        
        Error 40622
    
    End If
    
    Exit Sub
    
Erro_TaxaDesconto_Validate:

    Cancel = True

    
    Select Case Err

        Case 40622
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154311)
    
    End Select

    Exit Sub
End Sub

Private Sub CodCarteiraNoBanco_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub NumCarteiraNoBanco_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub


Private Sub NomeNoBanco_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objCobrador As New ClassCobrador
Dim objCarteiraCobrador As New ClassCarteiraCobrador
Dim objEndereco As New ClassEndereco
Dim sNumeroInicial As String
Dim sNumeroFinal As String
Dim bInicialMenor As Boolean

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se foi preenchido o Código
    If Len(Trim(Codigo.Text)) = 0 Then Error 23465

    'Se a opção "banco" estiver selecionada
    If OpcaoBanco.Value Then

        If Banco.ListIndex = -1 Then Error 19415
    
    End If
    
    'Verifica se Nome Reduzido foi preenchido
    If Len(Trim(NomeReduzido.Text)) = 0 Then Error 23468

    'verifica se tem pelo menos uma carteira
    If Carteira.ListCount = 0 Then Error 40599

    'Se utiliza cobrança eletrônica
    If CobrancaEletronica.Value = COBRANCA_ELETRONICA Then
        If Len(Trim(ContaCorrente.Text)) = 0 Then Error 62012
    End If
        
    'Preenche os objetos com os dados da tela
    lErro = Move_Tela_Memoria(objCobrador, objEndereco)
    If lErro <> SUCESSO Then Error 23466
    
    lErro = Trata_Alteracao(objCobrador, objCobrador.iCodigo)
    If lErro <> SUCESSO Then Error 23490
    
    'Tranferir dados da Carteira na Tela para objCarteiraCobrador
    lErro = Move_Carteira_Memoria(objCarteiraCobrador)
    If lErro <> SUCESSO Then Error 23579

    If objCarteiraCobrador.iCodCarteiraCobranca > 0 Then
        gcolCarteirasCobrador.Remove CStr(objCarteiraCobrador.iCodCarteiraCobranca)
        gcolCarteirasCobrador.Add objCarteiraCobrador, CStr(objCarteiraCobrador.iCodCarteiraCobranca)
    End If
    
    iIndice = 0
    
    If OpcaoBanco.Value = True Then
    
        For Each objCarteiraCobrador In gcolCarteirasCobrador
            
            
            If (Len(Trim(objCarteiraCobrador.sFaixaNossoNumeroInicial)) > 0 Or Len(Trim(objCarteiraCobrador.sFaixaNossoNumeroFinal)) > 0 Or Len(Trim(objCarteiraCobrador.sFaixaNossoNumeroProx)) > 0) And (Len(Trim(objCarteiraCobrador.sFaixaNossoNumeroInicial)) = 0 Or Len(Trim(objCarteiraCobrador.sFaixaNossoNumeroFinal)) = 0 Or Len(Trim(objCarteiraCobrador.sFaixaNossoNumeroProx)) = 0) Then Error 62107
            
            If Len(Trim(NumeroInicial)) > 0 Then
                sNumeroInicial = objCarteiraCobrador.sFaixaNossoNumeroInicial
                sNumeroFinal = objCarteiraCobrador.sFaixaNossoNumeroFinal
            
                Call Compara_NossoNumero(sNumeroInicial, sNumeroFinal, bInicialMenor)
                
                If Not bInicialMenor Then Error 62108
                
                sNumeroInicial = objCarteiraCobrador.sFaixaNossoNumeroProx
                sNumeroFinal = objCarteiraCobrador.sFaixaNossoNumeroFinal
            
                Call Compara_NossoNumero(sNumeroInicial, sNumeroFinal, bInicialMenor)
                
                If Not bInicialMenor Then Error 62109
                
                sNumeroInicial = objCarteiraCobrador.sFaixaNossoNumeroInicial
                sNumeroFinal = objCarteiraCobrador.sFaixaNossoNumeroProx
            
                Call Compara_NossoNumero(sNumeroInicial, sNumeroFinal, bInicialMenor)
                
                If Not bInicialMenor Then Error 62106
            
            End If
        Next
    End If
        
    'Grava o Cobrador e o seu respectivo endereço no BD
    lErro = CF("Cobrador_Grava", objCobrador, gcolCarteirasCobrador, objEndereco)
    If lErro <> SUCESSO Then Error 23467

    'Atualiza ListBox de Cobradores
    Call Lista_Cobradores_Remove(objCobrador)
    Call Lista_Cobradores_Adiciona(objCobrador)

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 23465
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)

        Case 23466, 23467

        Case 23468
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_PREENCHIDO", Err, Error$)

        Case 23490, 23579

        Case 19415
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BANCO_NAO_INFORMADO", Err)
        
        Case 40599
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CARTEIRA_COBANCA_NAO_INFORMADA", Err)
        
        Case 62012
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_COBRANCA_NAO_INFORMADA", Err)
        
        Case 62106
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PROXNOSSONUMERO_MENOR", Err, objCarteiraCobrador.iCodCarteiraCobranca)
        
        Case 62107
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FAIXA_NOSSONUMERO_IMCOMPLETA", Err, objCarteiraCobrador.iCodCarteiraCobranca)
        
        Case 62108
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOSSONUMERO_INICIAL_MAIOR", Err, objCarteiraCobrador.iCodCarteiraCobranca)
        
        Case 62109
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PROXNOSSONUMERO_MAIOR", Err, objCarteiraCobrador.iCodCarteiraCobranca)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154312)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(objCobrador As ClassCobrador, objEndereco As ClassEndereco) As Long
'Lê os dados que estão na tela Cobrador e coloca em objCobrador

Dim lErro As Long
Dim iPosicao As Integer
Dim objFornecedor As New ClassFornecedor
Dim colEnderecos As New Collection

On Error GoTo Erro_Move_Tela_Memoria

    If Len(Trim(Codigo.Text)) > 0 Then objCobrador.iCodigo = CInt(Codigo.Text)
    
    objCobrador.iFilialEmpresa = giFilialEmpresa
    
    If CobradorInativo.Value = vbChecked Then
        objCobrador.iInativo = Inativo
    Else
        objCobrador.iInativo = Ativo
    End If
    
    If Len(Trim(Fornecedor.Text)) > 0 Then
    
        objFornecedor.sNomeReduzido = Fornecedor.Text
            
        'busca o códigodo fornecedor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 109896
        
        'naum existe fornecedor
        If lErro = 6681 Then gError 109897
        
        objCobrador.lFornecedor = objFornecedor.lCodigo
        
    End If
    
    If Filial.ListIndex <> -1 Then objCobrador.iFilial = Codigo_Extrai(Filial.Text)
    
    If OpcaoBanco.Value Then
    
        If (Len(Trim(Banco.Text))) > 0 Then
            
            objCobrador.iCodBanco = Banco.ItemData(Banco.ListIndex)
    
            'Pegar nome reduzido do Banco
            iPosicao = InStr(Trim(Banco.List(Banco.ListIndex)), SEPARADOR) + 1
    
        End If
        
    Else
    
        objCobrador.iCodBanco = 0 'Cobrador não é um banco

    End If

    objCobrador.sNomeReduzido = Trim(NomeReduzido.Text)
    objCobrador.sNome = Trim(Nome)
    objCobrador.iCodCCI = Codigo_Extrai(ContaCorrente.Text)
    objCobrador.iCobrancaEletronica = CobrancaEletronica.Value
    objCobrador.lCNABProxSeqArqCobr = StrParaLong(CNABProxSeqArqCobr.Text)
    
    lErro = gobjTabEnd.Move_Endereco_Memoria(colEnderecos)
    If lErro <> SUCESSO Then gError 109896

    Set objEndereco = colEnderecos.Item(1)

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
        
        Case 109896
            
        Case 109897
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INEXISTENTE", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154313)

    End Select

    Exit Function

End Function

Sub Limpa_Tela_Cobrador()

Dim lErro As Long
Dim iIndice As Integer

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Limpa TextBox e MaskedEditBox
    Call Limpa_Tela(Me)

    'Limpa os textos das Combos
    Banco.ListIndex = -1
    ContaCorrente.Text = ""

    'Limpar Combo Carteiras
    Carteira.Clear

    'Selecionar opção banco
    OpcaoBanco.Value = True

    CobradorInativo.Value = vbUnchecked
    Desativada.Value = 0

    'Limpa os Label's
    QtdParcelas.Caption = ""
    QtdBanco.Caption = ""
    SaldoValor.Caption = ""
    SaldoBanco.Caption = ""

    'Desseleciona ListBox de Cobradores
    Lista_Cobradores.ListIndex = -1

    'Limpar coleção
    Set gcolCarteirasCobrador = New Collection
    
    'Exibe código na Tela
    Codigo.PromptInclude = False
    Codigo.Text = ""
    Codigo.PromptInclude = True

    BcoImprimeBoleta.Value = vbChecked
    GeraNossoNumero.Value = vbUnchecked
    Registrada.Value = vbUnchecked
    FormularioPreImpresso.Value = vbUnchecked

    Call gobjTabEnd.Limpa_Tela
    
    iCarteiraAnterior = -1
        
    Filial.Clear
    
    iAlterado = 0

End Sub

Function Exibe_Dados_Cobrador(objCobrador As ClassCobrador) As Long

Dim lErro As Long
Dim sListBoxItem As String
Dim iIndice As Integer
Dim iCodigo As Integer
Dim objCarteiraCobrador As New ClassCarteiraCobrador
Dim objEndereco As New ClassEndereco, bCancel As Boolean
Dim objCarteiraCobranca As New ClassCarteiraCobranca
Dim colEnderecos As New Collection

On Error GoTo Erro_Exibe_Dados_Cobrador

    Call Limpa_Tela_Cobrador

    'Carrega lEndereco em objEndereco
    objEndereco.lCodigo = objCobrador.lEndereco

    'Le o endereco a partir do Codigo
    lErro = CF("Endereco_Le", objEndereco)
    If lErro <> SUCESSO And lErro <> 12309 Then Error 23494

    If lErro = 12309 Then Error 23495
    
    colEnderecos.Add objEndereco

    'Exibe os dados de objCobrador na tela
    If objCobrador.iCodigo = 0 Then
        Codigo.PromptInclude = False
        Codigo.Text = ""
        Codigo.PromptInclude = True
    Else
        Codigo.PromptInclude = False
        Codigo.Text = CStr(objCobrador.iCodigo)
        Codigo.PromptInclude = True
        'Codigo_Validate
    End If

    CobradorInativo.Value = IIf(objCobrador.iInativo = Inativo, vbChecked, vbUnchecked)

    If objCobrador.iCodBanco <> 0 Then
        OpcaoBanco.Value = True

        'Exibir dados específicos para bancos
        For iIndice = 0 To Banco.ListCount - 1
            
            iCodigo = Codigo_Extrai(Banco.List(iIndice))
            If iCodigo = objCobrador.iCodBanco Then
                Banco.ListIndex = iIndice
                Exit For
            End If

        Next

    Else
        OpcaoOutros.Value = True

        'Limpar campos relativos a banco
        Banco.ListIndex = -1

    End If

    'Exibir dados específicos para cobrador não banco
    NomeReduzido.Text = objCobrador.sNomeReduzido
    Nome.Text = objCobrador.sNome

    If objCobrador.iCodCCI <> 0 Then
        ContaCorrente.Text = CStr(objCobrador.iCodCCI)
        Call ContaCorrente_Validate(bCancel)
    Else
        ContaCorrente.ListIndex = -1
    End If
    
    CobrancaEletronica.Value = objCobrador.iCobrancaEletronica

    lErro = gobjTabEnd.Traz_Endereco_Tela(colEnderecos)
    If lErro <> SUCESSO Then Error 23496

    'Limpar Combo Carteiras
    Carteira.Clear

    'Limpar coleção
    If Not (gcolCarteirasCobrador Is Nothing) Then

        Set gcolCarteirasCobrador = New Collection

    End If
    
    'Le as carteiras associadas ao Cobrador
    lErro = CF("Cobrador_Le_Carteiras", objCobrador, gcolCarteirasCobrador)
    If lErro <> SUCESSO And lErro <> 23500 Then Error 23496

    If lErro = SUCESSO Then
        'Preencher a Combo
        For Each objCarteiraCobrador In gcolCarteirasCobrador
                       
            objCarteiraCobranca.iCodigo = objCarteiraCobrador.iCodCarteiraCobranca
        
            lErro = CF("CarteiraDeCobranca_Le", objCarteiraCobranca)
            If lErro <> SUCESSO And lErro <> 23413 Then Error 40591
        
            'Carteira não está cadastrado
            If lErro = 23413 Then Error 40592
       
            'Concatena Código e a Descricao da carteira
            sListBoxItem = CStr(objCarteiraCobranca.iCodigo)
            sListBoxItem = sListBoxItem & SEPARADOR & objCarteiraCobranca.sDescricao
              
            Carteira.AddItem sListBoxItem
            Carteira.ItemData(Carteira.NewIndex) = objCarteiraCobranca.iCodigo
        Next

    End If
        
    If objCobrador.lFornecedor > 0 Then
        Fornecedor.Text = objCobrador.lFornecedor
        Call Fornecedor_Validate(False)
    End If
    
    If objCobrador.iFilial > 0 Then
        Filial.Text = objCobrador.iFilial
        Call Filial_Validate(False)
    End If
    
    CNABProxSeqArqCobr.PromptInclude = False
    CNABProxSeqArqCobr.Text = CStr(objCobrador.lCNABProxSeqArqCobr)
    CNABProxSeqArqCobr.PromptInclude = True
    
    'Desselecionar a combo carteira
    Carteira.ListIndex = -1

    iAlterado = 0

    Exibe_Dados_Cobrador = SUCESSO
    
    Exit Function

Erro_Exibe_Dados_Cobrador:

    Exibe_Dados_Cobrador = Err

    Select Case Err

        Case 23494

        Case 23495
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_NAO_CADASTRADO", Err)

        Case 23496
        
        Case 40591
        
        Case 40592
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRANCA_NAO_CADASTRADA", Err, objCarteiraCobranca.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154314)

    End Select

    Exit Function

End Function

Private Sub Lista_Cobradores_Remove(objCobrador As ClassCobrador)
'Percorre a ListBox Lista_Cobradores para remover o Cobrador caso ela exista

Dim iIndice As Integer

For iIndice = 0 To Lista_Cobradores.ListCount - 1

    If Lista_Cobradores.ItemData(iIndice) = objCobrador.iCodigo Then

        Lista_Cobradores.RemoveItem iIndice
        Exit For

    End If

Next

End Sub

Function Limpa_Tab_Carteiras()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tab_Carteiras

    'Limpa Tab
    Desativada.Value = 0
    TaxaCobranca.Text = ""
    DiasdeRetencao.Text = ""
    
    Conta.PromptInclude = False
    Conta.Text = ""
    Conta.PromptInclude = True
    
    TaxaDesconto.Text = ""
    
    ContaContabilDup.PromptInclude = False
    ContaContabilDup.Text = ""
    ContaContabilDup.PromptInclude = True

    'dados estatísticos
    QtdParcelas.Caption = ""
    QtdBanco.Caption = ""
    SaldoValor.Caption = ""
    SaldoBanco.Caption = ""

    'Dados específicos para Cobrador Banco
    If OpcaoBanco.Value Then

        CodCarteiraNoBanco.Text = ""
        NumCarteiraNoBanco.Text = ""
        NomeNoBanco.Text = ""
        NumeroInicial.Text = ""
        NumeroFinal.Text = ""
        NumeroProx.Text = ""

    End If
    
    FormularioPreImpresso.Value = vbUnchecked

    iCarteiraAnterior = -1
    
    Limpa_Tab_Carteiras = SUCESSO
    
    Exit Function

Erro_Limpa_Tab_Carteiras:

    Limpa_Tab_Carteiras = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154315)

    End Select

    Exit Function

End Function

Function Exibe_CarteiraCobrador(objCarteiraCobrador As ClassCarteiraCobrador) As Long
'exibe carteira selecionada

Dim lErro As Long
Dim sContaEnxuta As String
Dim sContaDuplEnxuta As String
Dim iAlteradoAnterior As Integer

On Error GoTo Erro_Exibe_CarteiraCobrador
    
    iAlteradoAnterior = iAlterado

    If objCarteiraCobrador.iDesativada = 1 Then
        Desativada.Value = 1
    Else
        Desativada.Value = 0
    End If
            
    TaxaCobranca.Text = CStr(objCarteiraCobrador.dTaxaCobranca)
    DiasdeRetencao.Text = CStr(objCarteiraCobrador.iDiasDeRetencao)
    TaxaDesconto.Text = CStr(objCarteiraCobrador.dTaxaDesconto)
                
    If objCarteiraCobrador.sContaContabil = "" Then
        
        Conta.PromptInclude = False
        Conta.Text = ""
        Conta.PromptInclude = True
     
    Else
    
        Conta.PromptInclude = False
        lErro = Mascara_RetornaContaEnxuta(objCarteiraCobrador.sContaContabil, sContaEnxuta)
        If lErro <> SUCESSO Then Error 40593

        Conta.Text = sContaEnxuta
        Conta.PromptInclude = True
    
    End If
    
    If objCarteiraCobrador.sContaDuplDescontadas = "" Then

        ContaContabilDup.PromptInclude = False
        ContaContabilDup.Text = ""
        ContaContabilDup.PromptInclude = True

    Else

        ContaContabilDup.PromptInclude = False
        lErro = Mascara_RetornaContaEnxuta(objCarteiraCobrador.sContaDuplDescontadas, sContaDuplEnxuta)
        If lErro <> SUCESSO Then Error 40594
    
        ContaContabilDup.Text = sContaDuplEnxuta
        ContaContabilDup.PromptInclude = True
    
    End If
        
    'dados estatísticos
    QtdParcelas.Caption = CStr(objCarteiraCobrador.lQuantidadeAtual)
    QtdBanco.Caption = CStr(objCarteiraCobrador.lQuantidadeAtualBanco)
    SaldoValor.Caption = Format(objCarteiraCobrador.dSaldoAtual, "Standard")
    SaldoBanco.Caption = CStr(objCarteiraCobrador.dSaldoAtualBanco)
    If objCarteiraCobrador.iComRegistro = CARTEIRA_COM_REGISTRO Then
        Registrada.Value = vbChecked
    Else
        Registrada.Value = vbUnchecked
    End If

    'Dados específicos de Cobrador Banco
    If (OpcaoBanco.Value) And CobrancaEletronica.Value = vbChecked Then

        CodCarteiraNoBanco.Text = objCarteiraCobrador.sCodCarteiraNoBanco
        NumCarteiraNoBanco.Text = CStr(objCarteiraCobrador.iNumCarteiraNoBanco)
        NomeNoBanco.Text = objCarteiraCobrador.sNomeNoBanco
        NumeroInicial.Text = objCarteiraCobrador.sFaixaNossoNumeroInicial
        NumeroFinal.Text = objCarteiraCobrador.sFaixaNossoNumeroFinal
        NumeroProx.Text = objCarteiraCobrador.sFaixaNossoNumeroProx
        If objCarteiraCobrador.iGeraNossoNumero = EMPRESA_GERA_NOSSONUMERO Then
            GeraNossoNumero.Value = vbChecked
        Else
            GeraNossoNumero.Value = vbUnchecked
        End If
        If objCarteiraCobrador.iImprimeBoleta = EMPRESA_IMPRIME_BOLETA Then
            BcoImprimeBoleta.Value = vbUnchecked
        Else
            BcoImprimeBoleta.Value = vbChecked
        End If
        If objCarteiraCobrador.iFormPreImp = DESMARCADO Then
            FormularioPreImpresso.Value = vbUnchecked
        Else
            FormularioPreImpresso.Value = vbChecked
        End If
    End If
    
    iAlterado = iAlteradoAnterior

    Exibe_CarteiraCobrador = SUCESSO

    Exit Function

Erro_Exibe_CarteiraCobrador:

    Exibe_CarteiraCobrador = Err

    Select Case Err
    
        Case 40593
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objCarteiraCobrador.sContaContabil)
        
        Case 40594
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objCarteiraCobrador.sContaDuplDescontadas)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154316)

    End Select

    Exit Function

End Function

Function Move_Carteira_Memoria(objCarteiraCobrador As ClassCarteiraCobrador) As Long

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaDupFormatada As String
Dim iContaPreenchida As Integer
Dim iContaDupPreenchida As Integer

On Error GoTo Erro_Move_Carteira_Memoria

    If Len(Trim(Codigo.Text)) > 0 Then objCarteiraCobrador.iCobrador = CInt(Codigo.Text)
    
    If Len(Trim(Carteira.List(Carteira.ListIndex))) > 0 Then objCarteiraCobrador.iCodCarteiraCobranca = Carteira.ItemData(Carteira.ListIndex)
    
    If Desativada.Value Then
        objCarteiraCobrador.iDesativada = 1
    Else
        objCarteiraCobrador.iDesativada = 0
    End If

    If Len(Trim(TaxaCobranca.Text)) > 0 Then objCarteiraCobrador.dTaxaCobranca = CDbl(TaxaCobranca.Text)
    If Len(Trim(DiasdeRetencao.Text)) > 0 Then objCarteiraCobrador.iDiasDeRetencao = CInt(DiasdeRetencao.Text)
    
    If Len(Trim(Conta.ClipText)) > 0 Then
    
        'Guarda a conta corrente
        lErro = CF("Conta_Formata", Conta.Text, sContaFormatada, iContaPreenchida)
        If lErro <> SUCESSO Then Error 11843
            
        objCarteiraCobrador.sContaContabil = sContaFormatada
        
    End If
    
    If Len(Trim(ContaContabilDup.ClipText)) > 0 Then
    
        'Guarda a conta corrente
        lErro = CF("Conta_Formata", ContaContabilDup.Text, sContaDupFormatada, iContaDupPreenchida)
        If lErro <> SUCESSO Then Error 11843
    
        objCarteiraCobrador.sContaDuplDescontadas = sContaDupFormatada
    
    End If
    
    If Len(Trim(TaxaDesconto.Text)) > 0 Then objCarteiraCobrador.dTaxaDesconto = CDbl(TaxaDesconto.Text)

    'DADOS ESPECÍFICOS PARA COBRADOR BANCO
    If (OpcaoBanco.Value) And CobrancaEletronica.Value = vbChecked Then
        If Registrada.Value Then
            objCarteiraCobrador.iComRegistro = CARTEIRA_COM_REGISTRO
        Else
            objCarteiraCobrador.iComRegistro = CARTEIRA_SEM_REGISTRO
        End If
        
        If BcoImprimeBoleta.Value = vbChecked Then
            objCarteiraCobrador.iImprimeBoleta = BANCO_IMPRIME_BOLETA
        Else
            objCarteiraCobrador.iImprimeBoleta = EMPRESA_IMPRIME_BOLETA
        End If
        
        If FormularioPreImpresso.Value = vbChecked Then
            objCarteiraCobrador.iFormPreImp = MARCADO
        Else
            objCarteiraCobrador.iFormPreImp = DESMARCADO
        End If
        
        objCarteiraCobrador.sCodCarteiraNoBanco = Trim(CodCarteiraNoBanco.Text)
        If Len(Trim(NumCarteiraNoBanco.Text)) > 0 Then objCarteiraCobrador.iNumCarteiraNoBanco = CInt(NumCarteiraNoBanco.Text)
        objCarteiraCobrador.sNomeNoBanco = Trim(NomeNoBanco.Text)
        If GeraNossoNumero.Value = vbChecked Then
            objCarteiraCobrador.iGeraNossoNumero = EMPRESA_GERA_NOSSONUMERO
            objCarteiraCobrador.sFaixaNossoNumeroInicial = 0
            objCarteiraCobrador.sFaixaNossoNumeroFinal = 0
            objCarteiraCobrador.sFaixaNossoNumeroProx = 0
        Else
            objCarteiraCobrador.iGeraNossoNumero = BANCO_GERA_NOSSONUMERO
            objCarteiraCobrador.sFaixaNossoNumeroInicial = Trim(NumeroInicial.Text)
            objCarteiraCobrador.sFaixaNossoNumeroFinal = Trim(NumeroFinal.Text)
            objCarteiraCobrador.sFaixaNossoNumeroProx = Trim(NumeroProx.Text)
        End If
        If Len(Trim(QtdParcelas.Caption)) > 0 Then objCarteiraCobrador.lQuantidadeAtual = CLng(QtdParcelas.Caption)
        If Len(Trim(QtdBanco.Caption)) > 0 Then objCarteiraCobrador.lQuantidadeAtualBanco = CLng(QtdBanco.Caption)
        If Len(Trim(SaldoValor.Caption)) > 0 Then objCarteiraCobrador.dSaldoAtual = CDbl(SaldoValor.Caption)
        If Len(Trim(SaldoBanco.Caption)) > 0 Then objCarteiraCobrador.dSaldoAtualBanco = CDbl(SaldoBanco.Caption)

    End If

    Move_Carteira_Memoria = SUCESSO

    Exit Function

Erro_Move_Carteira_Memoria:

    Move_Carteira_Memoria = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154317)

    End Select

    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objCobrador As New ClassCobrador
Dim objCarteiraCobrador As New ClassCarteiraCobrador
Dim objEndereco As New ClassEndereco

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Cobradores"

    'Le os dados da Tela Cobradores
    lErro = Move_Tela_Memoria(objCobrador, objEndereco)
    If lErro <> SUCESSO Then Error 23492

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objCobrador.iCodigo, 0, "Codigo"
    colCampoValor.Add "Inativo", objCobrador.iInativo, 0, "Inativo"
    colCampoValor.Add "NomeReduzido", objCobrador.sNomeReduzido, STRING_COBRADOR_NOME_REDUZIDO, "NomeReduzido"
    colCampoValor.Add "Nome", objCobrador.sNome, STRING_COBRADORES_NOME, "Nome"
    'colCampoValor.Add "Endereco", objCobrador.lEndereco, 0, "Endereco"
    colCampoValor.Add "CodBanco", objCobrador.iCodBanco, 0, "CodBanco"
    colCampoValor.Add "CobrancaEletronica", objCobrador.iCobrancaEletronica, 0, "CobrancaEletronica"
    colCampoValor.Add "CodCCI", objCobrador.iCodCCI, 0, "CodCCI"
    
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "Codigo", OP_DIFERENTE, COBRADOR_PROPRIA_EMPRESA
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 23492, 23590

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154318)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objCobrador As New ClassCobrador

On Error GoTo Erro_Tela_Preenche

    objCobrador.iCodigo = colCampoValor.Item("Codigo").vValor

    If objCobrador.iCodigo > 0 Then

        'Lê o Cobrador no BD
        lErro = CF("Cobrador_Le", objCobrador)
        If lErro <> SUCESSO And lErro <> 19250 Then Error 41519

        'Se Cobrador não está cadastrado, erro
        If lErro <> SUCESSO Then Error 41520

        'Traz dados do Cobrador para a Tela
        lErro = Exibe_Dados_Cobrador(objCobrador)
        If lErro <> SUCESSO Then Error 23493

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 23493, 41519, 41520

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154319)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Private Function Inicializa_Contabilidade_Tela()

Dim sMascaraConta As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Contabilidade_Tela

    'Verifica se o modulo de contabilidade esta ativo antes das inicializacoes
    If (gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO) Then
    
        Set objEventoCtaCarteira = New AdmEvento
        Set objEventoCtaDesconto = New AdmEvento
        
        'le a mascara das contas
        lErro = MascaraConta(sMascaraConta)
        If lErro <> SUCESSO Then Error 40597
        
        Conta.Mask = sMascaraConta
        ContaContabilDup.Mask = sMascaraConta
    
    Else
    
        'Incluido a inicialização da máscara para não dar erro na gravação de clientes com conta mas que o módulo de contabilidade foi desabilitado
        lErro = MascaraConta(sMascaraConta)
        If lErro <> SUCESSO Then Error 40597
        
        Conta.Mask = sMascaraConta
        ContaContabilDup.Mask = sMascaraConta
        
        Conta.Enabled = False
        ContaContabilDup.Enabled = False
        LabelCtaCarteira.Enabled = False
        LabelCtaDesconto.Enabled = False
        
    End If
    
    Inicializa_Contabilidade_Tela = SUCESSO
     
    Exit Function
    
Erro_Inicializa_Contabilidade_Tela:

    Inicializa_Contabilidade_Tela = Err
     
    Select Case Err
          
        Case 40597
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154320)
     
    End Select
     
    Exit Function

End Function

Private Sub Lista_Cobradores_Adiciona(objCobrador As ClassCobrador)
'Inclui Cobrador na List

    Lista_Cobradores.AddItem objCobrador.sNomeReduzido
    Lista_Cobradores.ItemData(Lista_Cobradores.NewIndex) = objCobrador.iCodigo

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object
    
    Parent.HelpContextID = IDH_COBRADORES_ID
    Set Form_Load_Ocx = Me
    Caption = "Cobradores"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Cobradores"
    
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

'***** fim do trecho a ser copiado ******

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is ContaCorrente Then
            Call ContaCorrenteLabel_Click
        ElseIf Me.ActiveControl Is ContaContabilDup Then
            Call LabelCtaDesconto_Click
        ElseIf Me.ActiveControl Is Conta Then
            Call LabelCtaCarteira_Click
        End If
    
    End If
    
End Sub


Sub Compara_NossoNumero(sNossoNumeroInicial As String, sNossoNumeroFinal As String, bInicialMenor As Boolean)

Dim iIndice As Integer
Dim iDigito1 As Integer
Dim iDigito2 As Integer
Dim bMaior As Boolean

    sNossoNumeroInicial = FormataCpoNum(sNossoNumeroInicial, STRING_NOSSO_NUMERO)
    sNossoNumeroFinal = FormataCpoNum(sNossoNumeroFinal, STRING_NOSSO_NUMERO)
    
    bMaior = False
    
    For iIndice = 1 To STRING_NOSSO_NUMERO
    
        iDigito1 = Mid(sNossoNumeroInicial, iIndice, 1)
        iDigito2 = Mid(sNossoNumeroFinal, iIndice, 1)

        If iDigito2 < iDigito1 Then
            bMaior = True
            Exit For
        ElseIf iDigito2 > iDigito1 Then
            Exit For
        End If
    Next
    
    bInicialMenor = Not bMaior

    Exit Sub

End Sub
Private Function FormataCpoNum(vData As Variant, iTam As Integer) As String
'formata campo numerico alinhado-o à direita sem ponto e decimais, com zeros a esquerda

Dim iData As Integer
Dim sData As String

    If Len(vData) = iTam Then

        FormataCpoNum = vData
        Exit Function

    End If

    iData = iTam - Len(vData)
    
    If iData > 0 Then sData = String(iData, "0")

    FormataCpoNum = sData & vData

    Exit Function

End Function

Private Sub Habilita_Campos_CobrancaEletronica(bHabilita As Boolean)

    CodCarteiraNoBanco.Enabled = bHabilita
    NumCarteiraNoBanco.Enabled = bHabilita
    NomeNoBanco.Enabled = bHabilita
    NumeroInicial.Enabled = bHabilita
    NumeroFinal.Enabled = bHabilita
    NumeroProx.Enabled = bHabilita
    LabelCodCartBco.Enabled = bHabilita
    LabelNumCartBco.Enabled = bHabilita
    LabelNomeCartBco.Enabled = bHabilita
    LabelNumIniCartBco.Enabled = bHabilita
    LabelNumFimCartBco.Enabled = bHabilita
    LabelNumProxCartBco.Enabled = bHabilita
    BcoImprimeBoleta.Enabled = bHabilita
    FormularioPreImpresso.Enabled = bHabilita
    GeraNossoNumero.Enabled = bHabilita
    Registrada.Enabled = bHabilita
    FrameCNAB.Enabled = bHabilita
    
End Sub


Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub

Private Sub Cobrador_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Cobrador(Index), Source, X, Y)
End Sub

Private Sub Cobrador_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Cobrador(Index), Button, Shift, X, Y)
End Sub


Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub ContaCorrenteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContaCorrenteLabel, Source, X, Y)
End Sub

Private Sub ContaCorrenteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContaCorrenteLabel, Button, Shift, X, Y)
End Sub

Private Sub LabelCtaDesconto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCtaDesconto, Source, X, Y)
End Sub

Private Sub LabelCtaDesconto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCtaDesconto, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeCartBco_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeCartBco, Source, X, Y)
End Sub

Private Sub LabelNomeCartBco_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeCartBco, Button, Shift, X, Y)
End Sub

Private Sub LabelNumCartBco_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNumCartBco, Source, X, Y)
End Sub

Private Sub LabelNumCartBco_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNumCartBco, Button, Shift, X, Y)
End Sub

Private Sub LabelCodCartBco_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodCartBco, Source, X, Y)
End Sub

Private Sub LabelCodCartBco_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodCartBco, Button, Shift, X, Y)
End Sub

Private Sub LabelNumIniCartBco_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNumIniCartBco, Source, X, Y)
End Sub

Private Sub LabelNumIniCartBco_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNumIniCartBco, Button, Shift, X, Y)
End Sub

Private Sub LabelNumFimCartBco_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNumFimCartBco, Source, X, Y)
End Sub

Private Sub LabelNumFimCartBco_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNumFimCartBco, Button, Shift, X, Y)
End Sub

Private Sub LabelNumProxCartBco_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNumProxCartBco, Source, X, Y)
End Sub

Private Sub LabelNumProxCartBco_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNumProxCartBco, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub QtdParcelas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QtdParcelas, Source, X, Y)
End Sub

Private Sub QtdParcelas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QtdParcelas, Button, Shift, X, Y)
End Sub

Private Sub SaldoValor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SaldoValor, Source, X, Y)
End Sub

Private Sub SaldoValor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SaldoValor, Button, Shift, X, Y)
End Sub

Private Sub SaldoBanco_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SaldoBanco, Source, X, Y)
End Sub

Private Sub SaldoBanco_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SaldoBanco, Button, Shift, X, Y)
End Sub

Private Sub Label19_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label19, Source, X, Y)
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label19, Button, Shift, X, Y)
End Sub

Private Sub QtdBanco_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QtdBanco, Source, X, Y)
End Sub

Private Sub QtdBanco_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QtdBanco, Button, Shift, X, Y)
End Sub

Private Sub Label24_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label24, Source, X, Y)
End Sub

Private Sub Label24_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label24, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub LabelCtaCarteira_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCtaCarteira, Source, X, Y)
End Sub

Private Sub LabelCtaCarteira_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCtaCarteira, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label15, Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15, Button, Shift, X, Y)
End Sub

Private Sub TabStripCartCobr_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStripCartCobr)
End Sub

Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub
