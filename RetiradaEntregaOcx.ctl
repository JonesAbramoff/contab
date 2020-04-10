VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl RetiradaEntregaOcx 
   ClientHeight    =   5970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9390
   ScaleHeight     =   5970
   ScaleWidth      =   9390
   Begin VB.CommandButton BotaoAnexos 
      Height          =   465
      Left            =   1020
      Picture         =   "RetiradaEntregaOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   106
      ToolTipText     =   "Anexar Arquivos"
      Top             =   5400
      Width           =   585
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4935
      Index           =   1
      Left            =   150
      TabIndex        =   59
      Top             =   360
      Width           =   9045
      Begin VB.Frame FrameNatureza 
         Caption         =   "Natureza"
         Height          =   525
         Left            =   30
         TabIndex        =   102
         Top             =   2175
         Width           =   9030
         Begin MSMask.MaskEdBox Natureza 
            Height          =   315
            Left            =   1785
            TabIndex        =   103
            Top             =   150
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label LabelNatureza 
            AutoSize        =   -1  'True
            Caption         =   "Natureza:"
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
            Left            =   915
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   105
            Top             =   180
            Width           =   840
         End
         Begin VB.Label LabelNaturezaDesc 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3315
            TabIndex        =   104
            Top             =   150
            Width           =   5640
         End
      End
      Begin VB.Frame FrameCcl 
         Caption         =   "Centro de Custos/Lucro"
         Height          =   570
         Index           =   0
         Left            =   30
         TabIndex        =   98
         Top             =   2745
         Width           =   9030
         Begin MSMask.MaskEdBox Ccl 
            Height          =   300
            Left            =   1785
            TabIndex        =   99
            Top             =   195
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin VB.Label CclLabel 
            AutoSize        =   -1  'True
            Caption         =   "Ccl:"
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
            Left            =   1395
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   101
            Top             =   225
            Width           =   345
         End
         Begin VB.Label DescCcl 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3330
            TabIndex        =   100
            Top             =   195
            Width           =   5640
         End
      End
      Begin VB.Frame FramePRJ 
         Caption         =   "Projeto"
         Height          =   1560
         Left            =   30
         TabIndex        =   82
         Top             =   3345
         Width           =   9030
         Begin VB.CommandButton BotaoProjetos 
            Caption         =   "..."
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
            Left            =   3420
            TabIndex        =   84
            Top             =   165
            Width           =   495
         End
         Begin VB.ComboBox Etapa 
            Height          =   315
            Left            =   1020
            Style           =   2  'Dropdown List
            TabIndex        =   83
            Top             =   855
            Width           =   7980
         End
         Begin MSMask.MaskEdBox Projeto 
            Height          =   285
            Left            =   1020
            TabIndex        =   85
            Top             =   180
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Etapa:"
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
            Height          =   180
            Index           =   62
            Left            =   390
            TabIndex        =   97
            Top             =   915
            Width           =   570
         End
         Begin VB.Label LabelProjeto 
            AutoSize        =   -1  'True
            Caption         =   "Projeto:"
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
            Left            =   285
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   96
            Top             =   210
            Width           =   675
         End
         Begin VB.Label Label2 
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
            Left            =   4410
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   95
            Top             =   210
            Width           =   660
         End
         Begin VB.Label PRJCli 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5130
            TabIndex        =   94
            Top             =   180
            Width           =   3840
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Descrição:"
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
            Left            =   30
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   93
            Top             =   555
            Width           =   915
         End
         Begin VB.Label PRJDesc 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1020
            TabIndex        =   92
            Top             =   510
            Width           =   7965
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Respons.:"
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
            Left            =   90
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   91
            Top             =   1245
            Width           =   870
         End
         Begin VB.Label PRJResp 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1020
            TabIndex        =   90
            Top             =   1215
            Width           =   3840
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Início:"
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
            Left            =   5070
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   89
            Top             =   1245
            Width           =   555
         End
         Begin VB.Label PRJDtIni 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5655
            TabIndex        =   88
            Top             =   1215
            Width           =   1275
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Fim:"
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
            Left            =   7305
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   87
            Top             =   1230
            Width           =   360
         End
         Begin VB.Label PRJDtFim 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7710
            TabIndex        =   86
            Top             =   1200
            Width           =   1275
         End
      End
      Begin VB.Frame FrameCompra 
         Caption         =   "Compra"
         Height          =   1230
         Left            =   30
         TabIndex        =   63
         Top             =   930
         Width           =   9030
         Begin VB.TextBox NotaEmpenho 
            Height          =   315
            Left            =   1800
            MaxLength       =   17
            TabIndex        =   36
            Top             =   165
            Width           =   7170
         End
         Begin VB.TextBox Pedido 
            Height          =   315
            Left            =   1800
            MaxLength       =   60
            TabIndex        =   37
            Top             =   510
            Width           =   7170
         End
         Begin VB.TextBox Contrato 
            Height          =   315
            Left            =   1800
            MaxLength       =   60
            TabIndex        =   38
            Top             =   855
            Width           =   7170
         End
         Begin VB.Label Label1Ret 
            AutoSize        =   -1  'True
            Caption         =   "Nota de Empenho:"
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
            Left            =   165
            TabIndex        =   66
            Top             =   210
            Width           =   1590
         End
         Begin VB.Label Label1Ret 
            AutoSize        =   -1  'True
            Caption         =   "Pedido:"
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
            Left            =   1095
            TabIndex        =   65
            Top             =   555
            Width           =   660
         End
         Begin VB.Label Label1Ret 
            AutoSize        =   -1  'True
            Caption         =   "Contrato:"
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
            Left            =   960
            TabIndex        =   64
            Top             =   900
            Width           =   795
         End
      End
      Begin VB.Frame FrameExp 
         Caption         =   "Exportação"
         Height          =   915
         Left            =   30
         TabIndex        =   60
         Top             =   15
         Width           =   9030
         Begin VB.ComboBox NumRE 
            Height          =   315
            Left            =   7125
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   165
            Width           =   1845
         End
         Begin VB.ComboBox UFEmbarque 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1800
            TabIndex        =   34
            Top             =   540
            Width           =   630
         End
         Begin VB.TextBox LocalEmbarque 
            Height          =   315
            Left            =   4365
            MaxLength       =   60
            TabIndex        =   35
            Top             =   525
            Width           =   4605
         End
         Begin MSMask.MaskEdBox NumDE 
            Height          =   315
            Left            =   1800
            TabIndex        =   32
            Tag             =   "Número da declaração de exportação padrão caso a informação não seja preenchida no item"
            Top             =   195
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   11
            Mask            =   "###########"
            PromptChar      =   " "
         End
         Begin VB.Label Label1Ret 
            AutoSize        =   -1  'True
            Caption         =   "Núm.RE Padrão:"
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
            Left            =   5670
            TabIndex        =   108
            Top             =   225
            Width           =   1410
         End
         Begin VB.Label LabelDE 
            AutoSize        =   -1  'True
            Caption         =   "Núm.DE Padrão:"
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
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   107
            Top             =   225
            Width           =   1410
         End
         Begin VB.Label Label1Ret 
            AutoSize        =   -1  'True
            Caption         =   "UF de embarque:"
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
            Left            =   285
            TabIndex        =   62
            Top             =   585
            Width           =   1470
         End
         Begin VB.Label Label1Ret 
            AutoSize        =   -1  'True
            Caption         =   "Local de embarque:"
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
            Left            =   2625
            TabIndex        =   61
            Top             =   585
            Width           =   1695
         End
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4740
      Index           =   2
      Left            =   150
      TabIndex        =   41
      Top             =   465
      Visible         =   0   'False
      Width           =   9045
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   4020
         Index           =   2
         Left            =   120
         TabIndex        =   67
         Top             =   300
         Visible         =   0   'False
         Width           =   8610
         Begin VB.TextBox LogradouroEnt 
            Height          =   315
            Left            =   4785
            MaxLength       =   40
            TabIndex        =   27
            Top             =   2535
            Width           =   3705
         End
         Begin VB.ComboBox PaisEnt 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4785
            TabIndex        =   22
            Top             =   1515
            Width           =   2535
         End
         Begin VB.ComboBox EstadoEnt 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7845
            TabIndex        =   23
            Top             =   1515
            Width           =   630
         End
         Begin VB.ComboBox TipoLogradouroEnt 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1605
            TabIndex        =   26
            Top             =   2535
            Width           =   2085
         End
         Begin VB.Frame FrameEnt 
            Caption         =   "Dados do Cliente/Fornecedor"
            Height          =   1050
            Left            =   135
            TabIndex        =   68
            Top             =   195
            Width           =   8310
            Begin VB.ComboBox FilialEnt 
               Height          =   315
               Left            =   4650
               TabIndex        =   20
               Top             =   630
               Width           =   1635
            End
            Begin VB.OptionButton OptionClienteEnt 
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
               Height          =   300
               Left            =   2655
               TabIndex        =   16
               Top             =   255
               Width           =   1185
            End
            Begin VB.OptionButton OptionFornecedorEnt 
               Caption         =   "Fornecedor"
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
               Left            =   4215
               TabIndex        =   17
               Top             =   255
               Width           =   1455
            End
            Begin MSMask.MaskEdBox FornecedorEnt 
               Height          =   300
               Left            =   1470
               TabIndex        =   18
               Top             =   660
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ClienteEnt 
               Height          =   300
               Left            =   1470
               TabIndex        =   19
               Top             =   660
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin VB.Label LabelFilialEnt 
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
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   4125
               TabIndex        =   71
               Top             =   690
               Width           =   465
            End
            Begin VB.Label ClienteLabelEnt 
               Alignment       =   1  'Right Justify
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
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   360
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   70
               Top             =   705
               Width           =   1035
            End
            Begin VB.Label FornecedorLabelEnt 
               Alignment       =   1  'Right Justify
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
               Height          =   195
               Left            =   360
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   69
               Top             =   705
               Width           =   1035
            End
         End
         Begin VB.CommandButton BotaoLimpaEnt 
            Height          =   330
            Left            =   8100
            Picture         =   "RetiradaEntregaOcx.ctx":0196
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Limpar"
            Top             =   3585
            Width           =   390
         End
         Begin MSMask.MaskEdBox BairroEnt 
            Height          =   315
            Left            =   4785
            TabIndex        =   25
            Top             =   2025
            Width           =   3690
            _ExtentX        =   6509
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CidadeEnt 
            Height          =   315
            Left            =   1605
            TabIndex        =   24
            Top             =   2025
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   12
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CEPEnt 
            Height          =   315
            Left            =   1605
            TabIndex        =   21
            Top             =   1515
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#####-###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NumeroEnt 
            Height          =   315
            Left            =   1605
            TabIndex        =   28
            Top             =   3045
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ComplementoEnt 
            Height          =   315
            Left            =   4785
            TabIndex        =   29
            Top             =   3045
            Width           =   3690
            _ExtentX        =   6509
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   12
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CNPJCPFEnt 
            Height          =   315
            Left            =   1590
            TabIndex        =   30
            Top             =   3555
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label PaisLabelEnt 
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
            Left            =   4290
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   81
            Top             =   1545
            Width           =   495
         End
         Begin VB.Label Label1Ent 
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
            Index           =   4
            Left            =   1155
            TabIndex        =   80
            Top             =   1575
            Width           =   465
         End
         Begin VB.Label LabelCidadeEnt 
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
            Left            =   930
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   79
            Top             =   2070
            Width           =   660
         End
         Begin VB.Label Label1Ent 
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
            Index           =   6
            Left            =   7500
            TabIndex        =   78
            Top             =   1560
            Width           =   315
         End
         Begin VB.Label Label1Ent 
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
            Index           =   14
            Left            =   4185
            TabIndex        =   77
            Top             =   2070
            Width           =   570
         End
         Begin VB.Label Label1Ent 
            AutoSize        =   -1  'True
            Caption         =   "Logradouro:"
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
            Left            =   3735
            TabIndex        =   76
            Top             =   2580
            Width           =   1035
         End
         Begin VB.Label Label1Ent 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Logradouro:"
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
            Left            =   120
            TabIndex        =   75
            Top             =   2580
            Width           =   1470
         End
         Begin VB.Label Label1Ent 
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
            Index           =   2
            Left            =   870
            TabIndex        =   74
            Top             =   3090
            Width           =   705
         End
         Begin VB.Label Label1Ent 
            AutoSize        =   -1  'True
            Caption         =   "Complemento:"
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
            Index           =   3
            Left            =   3570
            TabIndex        =   73
            Top             =   3090
            Width           =   1200
         End
         Begin VB.Label Label1Ent 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ/CPF:"
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
            Left            =   525
            TabIndex        =   72
            Top             =   3600
            Width           =   990
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4020
         Index           =   1
         Left            =   120
         TabIndex        =   42
         Top             =   300
         Width           =   8610
         Begin VB.ComboBox TipoLogradouroRet 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1605
            TabIndex        =   10
            Top             =   2550
            Width           =   2085
         End
         Begin VB.ComboBox EstadoRet 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7845
            TabIndex        =   7
            Top             =   1560
            Width           =   630
         End
         Begin VB.ComboBox PaisRet 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4785
            TabIndex        =   6
            Top             =   1560
            Width           =   2535
         End
         Begin VB.TextBox LogradouroRet 
            Height          =   315
            Left            =   4785
            MaxLength       =   40
            TabIndex        =   11
            Top             =   2550
            Width           =   3705
         End
         Begin VB.Frame Frame3 
            Caption         =   "Dados do Cliente/Fornecedor"
            Height          =   1050
            Left            =   150
            TabIndex        =   43
            Top             =   255
            Width           =   8310
            Begin VB.OptionButton OptionFornecedorRet 
               Caption         =   "Fornecedor"
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
               Left            =   4215
               TabIndex        =   1
               Top             =   255
               Width           =   1455
            End
            Begin VB.OptionButton OptionClienteRet 
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
               Height          =   300
               Left            =   2655
               TabIndex        =   0
               Top             =   255
               Width           =   1185
            End
            Begin VB.ComboBox FilialRet 
               Height          =   315
               Left            =   4635
               TabIndex        =   4
               Top             =   630
               Width           =   1635
            End
            Begin MSMask.MaskEdBox FornecedorRet 
               Height          =   300
               Left            =   1470
               TabIndex        =   2
               Top             =   660
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ClienteRet 
               Height          =   300
               Left            =   1470
               TabIndex        =   3
               Top             =   660
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin VB.Label FornecedorLabelRet 
               Alignment       =   1  'Right Justify
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
               Height          =   195
               Left            =   405
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   46
               Top             =   705
               Width           =   1035
            End
            Begin VB.Label ClienteLabelRet 
               Alignment       =   1  'Right Justify
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
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   405
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   45
               Top             =   705
               Width           =   1035
            End
            Begin VB.Label LabelFilialRet 
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
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   4110
               TabIndex        =   44
               Top             =   690
               Width           =   465
            End
         End
         Begin VB.CommandButton BotaoLimpaRet 
            Height          =   330
            Left            =   8085
            Picture         =   "RetiradaEntregaOcx.ctx":06C8
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Limpar"
            Top             =   3570
            Width           =   390
         End
         Begin MSMask.MaskEdBox BairroRet 
            Height          =   315
            Left            =   4785
            TabIndex        =   9
            Top             =   2055
            Width           =   3690
            _ExtentX        =   6509
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CidadeRet 
            Height          =   315
            Left            =   1605
            TabIndex        =   8
            Top             =   2055
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   12
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CEPRet 
            Height          =   315
            Left            =   1605
            TabIndex        =   5
            Top             =   1560
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#####-###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NumeroRet 
            Height          =   315
            Left            =   1605
            TabIndex        =   12
            Top             =   3060
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ComplementoRet 
            Height          =   315
            Left            =   4785
            TabIndex        =   13
            Top             =   3060
            Width           =   3690
            _ExtentX        =   6509
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   12
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CNPJCPFRet 
            Height          =   315
            Left            =   1590
            TabIndex        =   14
            Top             =   3570
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label Label1Ret 
            AutoSize        =   -1  'True
            Caption         =   "Complemento:"
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
            Index           =   3
            Left            =   3570
            TabIndex        =   56
            Top             =   3105
            Width           =   1200
         End
         Begin VB.Label Label1Ret 
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
            Index           =   2
            Left            =   870
            TabIndex        =   55
            Top             =   3105
            Width           =   705
         End
         Begin VB.Label Label1Ret 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Logradouro:"
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
            Left            =   120
            TabIndex        =   54
            Top             =   2595
            Width           =   1470
         End
         Begin VB.Label Label1Ret 
            AutoSize        =   -1  'True
            Caption         =   "Logradouro:"
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
            Left            =   3735
            TabIndex        =   53
            Top             =   2595
            Width           =   1035
         End
         Begin VB.Label Label1Ret 
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
            Index           =   14
            Left            =   4185
            TabIndex        =   52
            Top             =   2100
            Width           =   570
         End
         Begin VB.Label Label1Ret 
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
            Index           =   6
            Left            =   7500
            TabIndex        =   51
            Top             =   1605
            Width           =   315
         End
         Begin VB.Label LabelCidadeRet 
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
            Left            =   930
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   50
            Top             =   2100
            Width           =   660
         End
         Begin VB.Label Label1Ret 
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
            Index           =   4
            Left            =   1155
            TabIndex        =   49
            Top             =   1620
            Width           =   465
         End
         Begin VB.Label PaisLabelRet 
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
            Left            =   4290
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   48
            Top             =   1590
            Width           =   495
         End
         Begin VB.Label Label1Ret 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ/CPF:"
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
            Left            =   525
            TabIndex        =   47
            Top             =   3615
            Width           =   990
         End
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   4455
         Left            =   30
         TabIndex        =   57
         Top             =   -15
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   7858
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Local de retirada diferente do emitente"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Local de entrega diferente do destinatário"
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
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4740
      Picture         =   "RetiradaEntregaOcx.ctx":0BFA
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   5370
      Width           =   885
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3750
      Picture         =   "RetiradaEntregaOcx.ctx":0CFC
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   5370
      Width           =   885
   End
   Begin MSComctlLib.TabStrip TabStrip2 
      Height          =   5325
      Left            =   45
      TabIndex        =   58
      Top             =   15
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   9393
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Exportação, Compras e Projeto"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Endereços"
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
Attribute VB_Name = "RetiradaEntregaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private gobjInfoAdic As ClassInfoAdic

Dim giFrameAtual As Integer
Dim giFrameAtual2 As Integer
Dim giCarregando As Integer

Private gobjTela As Object
Private iIndexBrasil As Integer
Private iIndexUF As Integer
Private sCEPEnt As String
Private sCEPRet As String
Private bMudouCEPEnt As Boolean
Private bMudouCEPRet As Boolean

Private iAlterado As Integer
Private iClienteRetAlterado As Integer
Private iClienteEntAlterado As Integer
Private iFornecedorRetAlterado As Integer
Private iFornecedorEntAlterado As Integer

Private sNumDEAnt As String

Private WithEvents objEventoPaisEnt As AdmEvento
Attribute objEventoPaisEnt.VB_VarHelpID = -1
Private WithEvents objEventoPaisRet As AdmEvento
Attribute objEventoPaisRet.VB_VarHelpID = -1
Private WithEvents objEventoCidadeEnt As AdmEvento
Attribute objEventoCidadeEnt.VB_VarHelpID = -1
Private WithEvents objEventoCidadeRet As AdmEvento
Attribute objEventoCidadeRet.VB_VarHelpID = -1
Private WithEvents objEventoClienteEnt As AdmEvento
Attribute objEventoClienteEnt.VB_VarHelpID = -1
Private WithEvents objEventoClienteRet As AdmEvento
Attribute objEventoClienteRet.VB_VarHelpID = -1
Private WithEvents objEventoFornecedorEnt As AdmEvento
Attribute objEventoFornecedorEnt.VB_VarHelpID = -1
Private WithEvents objEventoFornecedorRet As AdmEvento
Attribute objEventoFornecedorRet.VB_VarHelpID = -1
Private WithEvents objEventoNatureza As AdmEvento
Attribute objEventoNatureza.VB_VarHelpID = -1
Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1
Private WithEvents objEventoDE As AdmEvento
Attribute objEventoDE.VB_VarHelpID = -1

Private Sub Form_Unload()

    Call gobjTela.gobjTelaProjetoInfo.Finaliza_InfoAdi
    
    Set objEventoPaisEnt = Nothing
    Set objEventoPaisRet = Nothing
    Set objEventoCidadeEnt = Nothing
    Set objEventoCidadeRet = Nothing
    Set objEventoClienteEnt = Nothing
    Set objEventoClienteRet = Nothing
    Set objEventoFornecedorEnt = Nothing
    Set objEventoFornecedorRet = Nothing
    Set objEventoNatureza = Nothing
    Set objEventoCcl = Nothing
    Set objEventoDE = Nothing
    
    Set gobjInfoAdic = Nothing
    Set gobjTela = Nothing
End Sub

Private Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer
Dim colCodigo As New Collection
Dim vCodigo As Variant, sMascaraCcl As String
Dim objCodigoDescricao As AdmCodigoNome
Dim colCodigoDescricao As New AdmColCodigoNome

On Error GoTo Erro_Form_Load

    giFrameAtual = 1
    giFrameAtual2 = 1
    
    Set objEventoPaisEnt = New AdmEvento
    Set objEventoPaisRet = New AdmEvento
    Set objEventoCidadeEnt = New AdmEvento
    Set objEventoCidadeRet = New AdmEvento
    Set objEventoClienteEnt = New AdmEvento
    Set objEventoClienteRet = New AdmEvento
    Set objEventoFornecedorEnt = New AdmEvento
    Set objEventoFornecedorRet = New AdmEvento
    Set objEventoNatureza = New AdmEvento
    Set objEventoCcl = New AdmEvento
    Set objEventoDE = New AdmEvento
    
    'Inicializa a mascara de Natureza
    lErro = Inicializa_Mascara_Natureza()
    If lErro <> SUCESSO Then gError 207577
    
    'Inicializa Máscara de Ccl
    sMascaraCcl = String(STRING_CCL, 0)

    lErro = MascaraCcl(sMascaraCcl)
    If lErro <> SUCESSO Then gError 207577

    Ccl.Mask = sMascaraCcl
    
    bMudouCEPEnt = False
    bMudouCEPRet = False
    
    lErro = CF("Carrega_Combo", TipoLogradouroEnt, "TiposDeLogradouro", "Sigla", TIPO_STR, "Nome", TIPO_STR)
    If lErro <> SUCESSO Then gError 207577

    lErro = CF("Carrega_Combo", TipoLogradouroRet, "TiposDeLogradouro", "Sigla", TIPO_STR, "Nome", TIPO_STR)
    If lErro <> SUCESSO Then gError 207578

    'Lê cada codigo da tabela Estados
    lErro = CF("Codigos_Le", "Estados", "Sigla", TIPO_STR, colCodigo, STRING_ESTADOS_SIGLA)
    If lErro <> SUCESSO Then gError 207579

    'Preenche as ComboBox Estados com os objetos da colecao colCodigo
    For Each vCodigo In colCodigo
        EstadoEnt.AddItem vCodigo
        EstadoRet.AddItem vCodigo
        UFEmbarque.AddItem vCodigo
    Next
    
    'Lê cada codigo e descricao da tabela Paises
    lErro = CF("Cod_Nomes_Le", "Paises", "Codigo", "Nome", STRING_PAISES_NOME, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 207580
    
    'Preenche cada ComboBox País com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao
        PaisEnt.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        PaisEnt.ItemData(PaisEnt.NewIndex) = objCodigoDescricao.iCodigo
        PaisRet.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        PaisRet.ItemData(PaisRet.NewIndex) = objCodigoDescricao.iCodigo
    Next
    
    LogradouroEnt.MaxLength = STRING_ENDERECO
    LogradouroRet.MaxLength = STRING_ENDERECO
    
    BairroEnt.MaxLength = STRING_BAIRRO
    BairroRet.MaxLength = STRING_BAIRRO
    
    CidadeEnt.MaxLength = STRING_CIDADE
    CidadeRet.MaxLength = STRING_CIDADE
    
    iAlterado = 0
    iClienteRetAlterado = 0
    iClienteEntAlterado = 0
    iFornecedorRetAlterado = 0
    iFornecedorEntAlterado = 0
    
    OptionClienteEnt.Value = True
    OptionClienteRet.Value = True
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case 207577 To 207580
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207581)
    
    End Select
    
    iAlterado = 0
    iClienteRetAlterado = 0
    iClienteEntAlterado = 0
    iFornecedorRetAlterado = 0
    iFornecedorEntAlterado = 0
    
End Sub

Public Function Trata_Parametros(ByVal objInfoAdic As ClassInfoAdic, ByVal objTela As Object) As Long
'Trata os parametros passados para a tela..

Dim lErro As Long
Dim sNaturezaEnxuta As String
Dim sCclMascarado As String

On Error GoTo Erro_Trata_Parametros

    Set gobjInfoAdic = objInfoAdic
    Set gobjTela = objTela
    
    If Not (gobjTela.gobjTelaProjetoInfo Is Nothing) Then
        Set gobjTela.gobjTelaProjetoInfo.objUserControlInfoAdi = Me
    End If
    
    If Len(Trim(gobjInfoAdic.sNatureza)) <> 0 Then
    
        sNaturezaEnxuta = String(STRING_NATMOVCTA_CODIGO, 0)
    
        lErro = Mascara_RetornaItemEnxuto(SEGMENTO_NATMOVCTA, gobjInfoAdic.sNatureza, sNaturezaEnxuta)
        If lErro <> SUCESSO Then gError 207582
    
        Natureza.PromptInclude = False
        Natureza.Text = sNaturezaEnxuta
        Natureza.PromptInclude = True
        
        Call Natureza_Validate(bSGECancelDummy)
        
    End If
    
    lErro = Traz_Endereco_Tela(gobjInfoAdic.objRetEnt)
    If lErro <> SUCESSO Then gError 207582
    
    lErro = Traz_Compra_Tela(gobjInfoAdic.objCompra)
    If lErro <> SUCESSO Then gError 207582
    
    lErro = Traz_Exportacao_Tela(gobjInfoAdic.objExportacao)
    If lErro <> SUCESSO Then gError 207582
    
    If Len(Trim(gobjInfoAdic.sCcl)) <> 0 Then
    
        sCclMascarado = String(STRING_CCL, 0)
        
        lErro = Mascara_RetornaCclEnxuta(gobjInfoAdic.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then gError 207582
    
        Ccl.PromptInclude = False
        Ccl.Text = sCclMascarado
        Ccl.PromptInclude = True
        
    Else
        
        Ccl.PromptInclude = False
        Ccl.Text = ""
        Ccl.PromptInclude = True
    
    End If
    Call Ccl_Validate(bSGECancelDummy)
    
'    If gobjTela.Projeto.ClipText <> "" Then
'        Projeto.PromptInclude = False
'        Projeto.Text = gobjTela.Projeto.Text
'        Projeto.PromptInclude = True
'        Call Projeto_Validate(bSGECancelDummy)
'
'        If gobjTela.Etapa.Text <> "" Then
'            Call CF("SCombo_Seleciona2", Etapa, gobjTela.Etapa.Text)
'        End If
'        Call Projeto_Validate(bSGECancelDummy)
'    End If

    If Not (gobjTela.gobjTelaProjetoInfo Is Nothing) Then
        lErro = gobjTela.gobjTelaProjetoInfo.Inicializa_InfoAdi
        If lErro <> SUCESSO Then gError 207582
    End If
    
    If gobjTela.Name = "PedidoCompras" Then
        FramePRJ.Enabled = True
        BotaoProjetos.Enabled = False
    Else
        FramePRJ.Enabled = False
    End If
    
    iAlterado = 0
    iClienteRetAlterado = 0
    iClienteEntAlterado = 0
    iFornecedorRetAlterado = 0
    iFornecedorEntAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case 207582
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207583)
    
    End Select
    
    If Not (gobjTela.gobjTelaProjetoInfo Is Nothing) Then
        Call gobjTela.gobjTelaProjetoInfo.Finaliza_InfoAdi
    End If
    
    iAlterado = 0
    iClienteRetAlterado = 0
    iClienteEntAlterado = 0
    iFornecedorRetAlterado = 0
    iFornecedorEntAlterado = 0
    
End Function

Public Function Move_Compra_Memoria(ByVal objCompra As ClassInfoAdicCompra) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Compra_Memoria

    objCompra.sContrato = Contrato.Text
    objCompra.sNotaEmpenho = NotaEmpenho.Text
    objCompra.sPedido = Pedido.Text

    Move_Compra_Memoria = SUCESSO

    Exit Function

Erro_Move_Compra_Memoria:

    Move_Compra_Memoria = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207594)

    End Select

    Exit Function
    
End Function

Public Function Move_Exportacao_Memoria(ByVal objExportacao As ClassInfoAdicExportacao) As Long

Dim lErro As Long
Dim objDE As New ClassDEInfo

On Error GoTo Erro_Move_Exportacao_Memoria

    objExportacao.sUFEmbarque = Trim(UFEmbarque.Text)
    objExportacao.sLocalEmbarque = LocalEmbarque.Text
    
    objDE.sNumero = Trim(NumDE.ClipText)
    
    If Len(Trim(objDE.sNumero)) > 0 Then
    
        lErro = CF("DEInfo_Le", objDE)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        objExportacao.lNumIntDE = objDE.lNumIntDoc
        
        objExportacao.sNumRE = Trim(NumRE.Text)
    
    End If

    Move_Exportacao_Memoria = SUCESSO

    Exit Function

Erro_Move_Exportacao_Memoria:

    Move_Exportacao_Memoria = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207594)

    End Select

    Exit Function
    
End Function

Public Function Move_Endereco_Memoria(ByVal objRetEnt As ClassRetiradaEntrega) As Long

Dim lErro As Long
Dim objEndereco As ClassEndereco
Dim objTab As Object
Dim objcliente As New ClassCliente
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Move_Endereco_Memoria

    Set objRetEnt.objEnderecoRet = New ClassEndereco

    objRetEnt.objEnderecoRet.sLogradouro = Trim(LogradouroRet.Text)
    objRetEnt.objEnderecoRet.sTipoLogradouro = SCodigo_Extrai(TipoLogradouroRet.Text)
    objRetEnt.objEnderecoRet.sComplemento = Trim(ComplementoRet.Text)
    objRetEnt.objEnderecoRet.lNumero = StrParaLong(NumeroRet.Text)
    objRetEnt.objEnderecoRet.sBairro = Trim(BairroRet.Text)
    objRetEnt.objEnderecoRet.sCidade = Trim(CidadeRet.Text)
    objRetEnt.objEnderecoRet.sCEP = Trim(CEPRet.Text)
    objRetEnt.objEnderecoRet.iCodigoPais = Codigo_Extrai(PaisRet.Text)
    objRetEnt.objEnderecoRet.sSiglaEstado = Trim(EstadoRet.Text)
    If objRetEnt.objEnderecoRet.iCodigoPais = 0 Then objRetEnt.objEnderecoRet.iCodigoPais = PAIS_BRASIL
    If objRetEnt.objEnderecoRet.iCodigoPais = PAIS_BRASIL And Len(objRetEnt.objEnderecoRet.sLogradouro) > 0 And (EstadoRet.ListIndex = -1 Or Len(Trim(EstadoRet.Text)) = 0) Then gError 207592
    objRetEnt.sCNPJCPFRet = Trim(CNPJCPFRet.Text)
    
    If OptionClienteRet.Value = True Then
        
        If Len(Trim(ClienteRet.ClipText)) > 0 Then
    
            objcliente.sNomeReduzido = ClienteRet.Text
            
            'Lê o Cliente
            lErro = CF("Cliente_Le_NomeReduzido", objcliente)
            If lErro <> SUCESSO And lErro <> 12348 Then gError 207584
            
            'Não encontrou p Cliente --> erro
            If lErro = 12348 Then gError 207585
    
            objRetEnt.lClienteRet = objcliente.lCodigo
            
        Else
            objRetEnt.lClienteRet = 0
        End If
        
        If Len(Trim(FilialRet.Text)) > 0 Then
            objRetEnt.iFilialCliRet = Codigo_Extrai(FilialRet.Text)
        Else
            objRetEnt.iFilialCliRet = 0
        End If

    Else
    
        If Len(Trim(FornecedorRet.Text)) > 0 Then
            
            objFornecedor.sNomeReduzido = FornecedorRet.Text
            
            'Lê o fornecedor
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then gError 207586
            
            If lErro = 6681 Then gError 207587 'Se nao encontrar --> erro
    
            objRetEnt.lFornecedorRet = objFornecedor.lCodigo
    
        Else
            objRetEnt.lFornecedorRet = 0
        End If
        
        If Len(Trim(FilialRet.Text)) > 0 Then
            objRetEnt.iFilialFornRet = Codigo_Extrai(FilialRet.Text)
        Else
            objRetEnt.iFilialFornRet = 0
        End If

    End If

    Set objRetEnt.objEnderecoEnt = New ClassEndereco

    objRetEnt.objEnderecoEnt.sLogradouro = Trim(LogradouroEnt.Text)
    objRetEnt.objEnderecoEnt.sTipoLogradouro = SCodigo_Extrai(TipoLogradouroEnt.Text)
    objRetEnt.objEnderecoEnt.sComplemento = Trim(ComplementoEnt.Text)
    objRetEnt.objEnderecoEnt.lNumero = StrParaLong(NumeroEnt.Text)
    objRetEnt.objEnderecoEnt.sBairro = Trim(BairroEnt.Text)
    objRetEnt.objEnderecoEnt.sCidade = Trim(CidadeEnt.Text)
    objRetEnt.objEnderecoEnt.sCEP = Trim(CEPEnt.Text)
    objRetEnt.objEnderecoEnt.iCodigoPais = Codigo_Extrai(PaisEnt.Text)
    objRetEnt.objEnderecoEnt.sSiglaEstado = Trim(EstadoEnt.Text)
    If objRetEnt.objEnderecoEnt.iCodigoPais = 0 Then objRetEnt.objEnderecoEnt.iCodigoPais = PAIS_BRASIL
    If objRetEnt.objEnderecoEnt.iCodigoPais = PAIS_BRASIL And Len(objRetEnt.objEnderecoEnt.sLogradouro) > 0 And (EstadoEnt.ListIndex = -1 Or Len(Trim(EstadoEnt.Text)) = 0) Then gError 207593
    objRetEnt.sCNPJCPFEnt = Trim(CNPJCPFEnt.Text)
    
    If OptionClienteEnt.Value = True Then
        
        If Len(Trim(ClienteEnt.ClipText)) > 0 Then
    
            objcliente.sNomeReduzido = ClienteEnt.Text
            
            'Lê o Cliente
            lErro = CF("Cliente_Le_NomeReduzido", objcliente)
            If lErro <> SUCESSO And lErro <> 12348 Then gError 207588
            
            'Não encontrou p Cliente --> erro
            If lErro = 12348 Then gError 207589
    
            objRetEnt.lClienteEnt = objcliente.lCodigo
            
        Else
            objRetEnt.lClienteEnt = 0
        End If

        If Len(Trim(FilialEnt.Text)) > 0 Then
            objRetEnt.iFilialCliEnt = Codigo_Extrai(FilialEnt.Text)
        Else
            objRetEnt.iFilialCliEnt = 0
        End If

    Else
    
        If Len(Trim(FornecedorEnt.Text)) > 0 Then
            
            objFornecedor.sNomeReduzido = FornecedorEnt.Text
            
            'Lê o fornecedor
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then gError 207590
            
            If lErro = 6681 Then gError 207591 'Se nao encontrar --> erro
    
            objRetEnt.lFornecedorEnt = objFornecedor.lCodigo
            
        Else
            objRetEnt.lFornecedorEnt = 0
        End If
        
        If Len(Trim(FilialEnt.Text)) > 0 Then
            objRetEnt.iFilialFornEnt = Codigo_Extrai(FilialEnt.Text)
        Else
            objRetEnt.iFilialFornEnt = 0
        End If

    End If
    
    Move_Endereco_Memoria = SUCESSO

    Exit Function

Erro_Move_Endereco_Memoria:

    Move_Endereco_Memoria = gErr

    Select Case gErr
    
        Case 207584, 207586, 207588, 207590
        
        Case 207585
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, ClienteRet.Text)
        
        Case 207587
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, FornecedorRet.Text)
    
        Case 207589
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, ClienteEnt.Text)
        
        Case 207591
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, FornecedorEnt.Text)
    
        Case 207592
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_CADASTRADO", gErr, EstadoRet.Text)

        Case 207593
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_CADASTRADO", gErr, EstadoEnt.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207594)

    End Select

    Exit Function

End Function

Public Function Traz_Compra_Tela(ByVal objCompra As ClassInfoAdicCompra) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_Compra_Tela

    giCarregando = 1

    If Not (objCompra Is Nothing) Then

        Contrato.Text = objCompra.sContrato
        NotaEmpenho.Text = objCompra.sNotaEmpenho
        Pedido.Text = objCompra.sPedido
    
    End If
    
    giCarregando = 0
        
    Traz_Compra_Tela = SUCESSO

    Exit Function

Erro_Traz_Compra_Tela:

    Traz_Compra_Tela = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207595)

    End Select

    giCarregando = 0

    Exit Function

End Function

Public Function Traz_Exportacao_Tela(ByVal objExportacao As ClassInfoAdicExportacao) As Long

Dim lErro As Long
Dim objDE As New ClassDEInfo

On Error GoTo Erro_Traz_Exportacao_Tela

    giCarregando = 1
    
    If Not (objExportacao Is Nothing) Then

        LocalEmbarque.Text = objExportacao.sLocalEmbarque
        
        UFEmbarque.Text = objExportacao.sUFEmbarque
        Call UFEmbarque_Validate(bSGECancelDummy)
        
        objDE.lNumIntDoc = objExportacao.lNumIntDE
        
        If objDE.lNumIntDoc > 0 Then
        
            lErro = CF("DEInfo_Le", objDE)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
            
            If lErro = SUCESSO Then
                NumDE.PromptInclude = False
                NumDE.Text = objDE.sNumero
                NumDE.PromptInclude = True
                Call NumDE_Validate(bSGECancelDummy)
                
                Call CF("SCombo_Seleciona2", NumRE, objExportacao.sNumRE)
            End If
            
        End If
        
    End If
    
    giCarregando = 0
        
    Traz_Exportacao_Tela = SUCESSO

    Exit Function

Erro_Traz_Exportacao_Tela:

    Traz_Exportacao_Tela = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207595)

    End Select

    giCarregando = 0

    Exit Function

End Function

Public Function Traz_Endereco_Tela(ByVal objRetEnt As ClassRetiradaEntrega) As Long

Dim lErro As Long
Dim bCancel As Boolean

On Error GoTo Erro_Traz_Endereco_Tela

    giCarregando = 1

    If Not (objRetEnt Is Nothing) Then
        If Not objRetEnt.objEnderecoRet Is Nothing Then
    
            'Se o cliente ainda tiver um código, ou seja, estiver cadastrado.
            If objRetEnt.lClienteRet <> 0 Then
            
                OptionClienteRet.Value = True
            
                ClienteRet.Text = objRetEnt.lClienteRet
                Call ClienteRet_Validate(bCancel)
        
                FilialRet.Text = objRetEnt.iFilialCliRet
                Call FilialRet_Validate(bCancel)
                
                
            ElseIf objRetEnt.lFornecedorRet <> 0 Then
                
                OptionFornecedorRet.Value = True
            
                
                FornecedorRet.Text = objRetEnt.lFornecedorRet
                Call FornecedorRet_Validate(bCancel)
        
                FilialRet.Text = objRetEnt.iFilialFornRet
                Call FilialRet_Validate(bCancel)
                
            End If
    
            PaisRet.Text = objRetEnt.objEnderecoRet.iCodigoPais
            Call PaisRet_Validate(bSGECancelDummy)
            BairroRet.Text = objRetEnt.objEnderecoRet.sBairro
            CidadeRet.Text = objRetEnt.objEnderecoRet.sCidade
            CEPRet.Text = objRetEnt.objEnderecoRet.sCEP
            EstadoRet.Text = objRetEnt.objEnderecoRet.sSiglaEstado
            Call EstadoRet_Validate(bSGECancelDummy)
            If objRetEnt.objEnderecoRet.iCodigoPais = 0 Then objRetEnt.objEnderecoRet.iCodigoPais = PAIS_BRASIL
        
            TipoLogradouroRet.Text = objRetEnt.objEnderecoRet.sTipoLogradouro
            Call TipoLogradouroRet_Validate(bSGECancelDummy)
            
            LogradouroRet.Text = objRetEnt.objEnderecoRet.sLogradouro
            
            If objRetEnt.objEnderecoRet.lNumero <> 0 Then
                NumeroRet.Text = CStr(objRetEnt.objEnderecoRet.lNumero)
            Else
                NumeroRet.Text = ""
            End If
            ComplementoRet.Text = objRetEnt.objEnderecoRet.sComplemento
        
        End If
        
        CNPJCPFRet.Text = objRetEnt.sCNPJCPFRet
        Call CNPJCPFRet_Validate(bSGECancelDummy)
            
        If Not objRetEnt.objEnderecoEnt Is Nothing Then
            
            'Se o cliente ainda tiver um código, ou seja, estiver cadastrado.
            If objRetEnt.lClienteEnt <> 0 Then
            
                OptionClienteEnt.Value = True
            
                ClienteEnt.Text = objRetEnt.lClienteEnt
                Call ClienteEnt_Validate(bCancel)
        
                FilialEnt.Text = objRetEnt.iFilialCliEnt
                Call FilialEnt_Validate(bCancel)
                
            ElseIf objRetEnt.lFornecedorEnt <> 0 Then
                
                OptionFornecedorEnt.Value = True
            
                FornecedorEnt.Text = objRetEnt.lFornecedorEnt
                Call FornecedorEnt_Validate(bCancel)
        
                FilialEnt.Text = objRetEnt.iFilialFornEnt
                Call FilialEnt_Validate(bCancel)
                
            End If
            
            PaisEnt.Text = objRetEnt.objEnderecoEnt.iCodigoPais
            Call PaisEnt_Validate(bSGECancelDummy)
            BairroEnt.Text = objRetEnt.objEnderecoEnt.sBairro
            CidadeEnt.Text = objRetEnt.objEnderecoEnt.sCidade
            CEPEnt.Text = objRetEnt.objEnderecoEnt.sCEP
            EstadoEnt.Text = objRetEnt.objEnderecoEnt.sSiglaEstado
            Call EstadoEnt_Validate(bSGECancelDummy)
            If objRetEnt.objEnderecoEnt.iCodigoPais = 0 Then objRetEnt.objEnderecoEnt.iCodigoPais = PAIS_BRASIL
        
            TipoLogradouroEnt.Text = objRetEnt.objEnderecoEnt.sTipoLogradouro
            Call TipoLogradouroEnt_Validate(bSGECancelDummy)
            
            LogradouroEnt.Text = objRetEnt.objEnderecoEnt.sLogradouro
            
            If objRetEnt.objEnderecoEnt.lNumero <> 0 Then
                NumeroEnt.Text = CStr(objRetEnt.objEnderecoEnt.lNumero)
            Else
                NumeroEnt.Text = ""
            End If
            ComplementoEnt.Text = objRetEnt.objEnderecoEnt.sComplemento
        
        End If
        
        CNPJCPFEnt.Text = objRetEnt.sCNPJCPFEnt
        Call CNPJCPFEnt_Validate(bSGECancelDummy)
        
    End If
        
    giCarregando = 0
        
    Traz_Endereco_Tela = SUCESSO

    Exit Function

Erro_Traz_Endereco_Tela:

    Traz_Endereco_Tela = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207595)

    End Select

    giCarregando = 0

    Exit Function

End Function

Private Sub BotaoAnexos_Click()
   Call Chama_Tela_Modal("Anexos", gobjInfoAdic.objAnexos)
End Sub

Private Sub BotaoCancela_Click()
    Unload Me
End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long
Dim objRetEnt As New ClassRetiradaEntrega
Dim objCompra As New ClassInfoAdicCompra
Dim objExportacao As New ClassInfoAdicExportacao
Dim sNaturezaFormatada As String
Dim iNaturezaPreenchida As Integer
Dim sCclFormatada As String, iCclPreenchida As Integer

On Error GoTo Erro_BotaoOK_Click

    sNaturezaFormatada = String(STRING_NATMOVCTA_CODIGO, 0)
    
    'Coloca no formato do BD
    lErro = CF("Item_Formata", SEGMENTO_NATMOVCTA, Natureza.Text, sNaturezaFormatada, iNaturezaPreenchida)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    gobjInfoAdic.sNatureza = sNaturezaFormatada
    
    lErro = CF("Ccl_Formata", Ccl.Text, sCclFormatada, iCclPreenchida)
    If lErro <> SUCESSO Then gError 22940

    If iCclPreenchida = CCL_PREENCHIDA Then
        gobjInfoAdic.sCcl = sCclFormatada
    Else
        gobjInfoAdic.sCcl = ""
    End If
    
    lErro = Move_Endereco_Memoria(objRetEnt)
    If lErro <> SUCESSO Then gError 207596
    
    lErro = Move_Compra_Memoria(objCompra)
    If lErro <> SUCESSO Then gError 207596
    
    lErro = Move_Exportacao_Memoria(objExportacao)
    If lErro <> SUCESSO Then gError 207596
    
    Set gobjInfoAdic.objRetEnt = objRetEnt
    
    If Len(Trim(objCompra.sContrato & objCompra.sNotaEmpenho & objCompra.sPedido)) > 0 Then
        Set gobjInfoAdic.objCompra = objCompra
    Else
        Set gobjInfoAdic.objCompra = Nothing
    End If
    
    If Len(Trim(objExportacao.sLocalEmbarque & objExportacao.sUFEmbarque & NumDE.ClipText & NumRE.Text)) > 0 Then
        Set gobjInfoAdic.objExportacao = objExportacao
    Else
        Set gobjInfoAdic.objExportacao = Nothing
    End If
    
    lErro = gobjTela.gobjTelaProjetoInfo.Transfere_InfoAdi
    If lErro <> SUCESSO Then gError 207596
    
    If iAlterado = REGISTRO_ALTERADO Then
        gobjTela.iAlterado = iAlterado
    End If
    
    Unload Me
    
    Exit Sub

Erro_BotaoOK_Click:

    Select Case gErr

        Case 207596, ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207597)

    End Select

    Exit Sub
    
End Sub

Public Sub ClienteRet_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objTipoCliente As New ClassTipoCliente
Dim colTipoFrete As Collection
Dim objTipoFrete As ClassTipoFrete
Dim iIndice As Integer

On Error GoTo Erro_ClienteRet_Validate

    'Verifica se o cliente foi alterado
    If iClienteRetAlterado = 0 Then Exit Sub
    
    'Se op cliente está preenchido
    If Len(Trim(ClienteRet.Text)) > 0 Then

        lErro = TP_Cliente_Le(ClienteRet, objcliente, iCodFilial)
        If lErro <> SUCESSO Then gError 207598

        lErro = CF("FiliaisClientes_Le_Cliente", objcliente, colCodigoNome)
        If lErro <> SUCESSO Then gError 207599

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", FilialRet, colCodigoNome)

        'Seleciona filial na Combo Filial
        If iCodFilial = FILIAL_MATRIZ Then
            FilialRet.ListIndex = 0
        Else
            Call CF("Filial_Seleciona", FilialRet, iCodFilial)
        End If
                    
                
    ElseIf Len(Trim(ClienteRet.Text)) = 0 Then

        FilialRet.Clear

    End If

    iClienteRetAlterado = 0

    Exit Sub

Erro_ClienteRet_Validate:

    Cancel = True

    Select Case gErr

        Case 207598, 207599

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207600)

    End Select

    Exit Sub

End Sub

Public Sub ClienteEnt_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objTipoCliente As New ClassTipoCliente
Dim colTipoFrete As Collection
Dim objTipoFrete As ClassTipoFrete
Dim iIndice As Integer

On Error GoTo Erro_ClienteEnt_Validate

    'Verifica se o cliente foi alterado
    If iClienteEntAlterado = 0 Then Exit Sub
    
    'Se op cliente está preenchido
    If Len(Trim(ClienteEnt.Text)) > 0 Then

        lErro = TP_Cliente_Le(ClienteEnt, objcliente, iCodFilial)
        If lErro <> SUCESSO Then gError 207601

        lErro = CF("FiliaisClientes_Le_Cliente", objcliente, colCodigoNome)
        If lErro <> SUCESSO Then gError 207602

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", FilialEnt, colCodigoNome)

        'Seleciona filial na Combo Filial
        If iCodFilial = FILIAL_MATRIZ Then
            FilialEnt.ListIndex = 0
        Else
            Call CF("Filial_Seleciona", FilialEnt, iCodFilial)
        End If
                    
                
    ElseIf Len(Trim(ClienteEnt.Text)) = 0 Then

        FilialEnt.Clear

    End If

    iClienteEntAlterado = 0

    Exit Sub

Erro_ClienteEnt_Validate:

    Cancel = True

    Select Case gErr

        Case 207601, 207602

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207603)

    End Select

    Exit Sub

End Sub

Public Sub FornecedorRet_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_FornecedorRet_Validate

    If iFornecedorRetAlterado = 1 Then

        If Len(Trim(FornecedorRet.Text)) > 0 Then

            'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
            lErro = TP_Fornecedor_Le3(FornecedorRet, objFornecedor, iCodFilial)
            If lErro <> SUCESSO Then gError 207604

            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then gError 207605

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", FilialRet, colCodigoNome)

            If colCodigoNome.Count = 1 Or iCodFilial <> 0 Then
            
                If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ
                
                'Seleciona filial na Combo Filial
                Call CF("Filial_Seleciona", FilialRet, iCodFilial)
                
            End If
            
            
        ElseIf Len(Trim(FornecedorRet.Text)) = 0 Then

            FilialRet.Clear

        End If

        iFornecedorRetAlterado = 0

    End If

    Exit Sub

Erro_FornecedorRet_Validate:

    Cancel = True

    Select Case gErr

        Case 207604, 207605

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207606)

    End Select

    Exit Sub

End Sub

Public Sub FornecedorEnt_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_FornecedorEnt_Validate

    If iFornecedorEntAlterado = 1 Then

        If Len(Trim(FornecedorEnt.Text)) > 0 Then

            'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
            lErro = TP_Fornecedor_Le3(FornecedorEnt, objFornecedor, iCodFilial)
            If lErro <> SUCESSO Then gError 207607

            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then gError 207608

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", FilialEnt, colCodigoNome)

            If colCodigoNome.Count = 1 Or iCodFilial <> 0 Then
            
                If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ
                
                'Seleciona filial na Combo Filial
                Call CF("Filial_Seleciona", FilialEnt, iCodFilial)
                
            End If
            
            
        ElseIf Len(Trim(FornecedorEnt.Text)) = 0 Then

            FilialEnt.Clear

        End If

        iFornecedorEntAlterado = 0

    End If

    Exit Sub

Erro_FornecedorEnt_Validate:

    Cancel = True

    Select Case gErr

        Case 207607, 207608

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207609)

    End Select

    Exit Sub

End Sub

Public Sub FilialRet_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialCliente As New ClassFilialCliente
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sCliente As String
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_FilialRet_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(FilialRet.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If FilialRet.Text = FilialRet.List(FilialRet.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialRet, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 207610

    If OptionClienteRet.Value = True Then

        'Se nao encontra o item com o código informado
        If lErro = 6730 Then
    
            'Verifica de o Cliente foi digitado
            If Len(Trim(ClienteRet.Text)) = 0 Then gError 207611
    
            sCliente = ClienteRet.Text
    
            objFilialCliente.iCodFilial = iCodigo
    
            'Pesquisa se existe filial com o codigo extraido
            lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
            If lErro <> SUCESSO And lErro <> 17660 Then gError 207612
    
            If lErro = 17660 Then gError 207613
    
            'Coloca na tela
            FilialRet.Text = iCodigo & SEPARADOR & objFilialCliente.sNome
    
            lErro = Trata_FilialClienteRet()
            If lErro <> SUCESSO Then gError 207614
    
        End If

        'Não encontrou valor informado que era STRING
        If lErro = 6731 Then gError 207615

    Else
    
        'Se nao encontra o ítem com o código informado
        If lErro = 6730 Then
    
            'Verifica de o fornecedor foi digitado
            If Len(Trim(FornecedorRet.Text)) = 0 Then gError 207616
    
            sFornecedor = FornecedorRet.Text
    
            objFilialFornecedor.iCodFilial = iCodigo
    
            'Pesquisa se existe filial com o codigo extraido
            lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 18272 Then gError 207617
    
            If lErro = 18272 Then gError 207618
    
            'coloca na tela
            FilialRet.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome
    
            lErro = Trata_FilialFornRet()
            If lErro <> SUCESSO Then gError 207619
    
    
        End If
    
        'Não encontrou valor informado que era STRING
        If lErro = 6731 Then gError 207620
    
    
    
    End If

    Exit Sub

Erro_FilialRet_Validate:

    Cancel = True


    Select Case gErr

        Case 207610, 207612, 207614, 207617, 207619

        Case 207611
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 207613
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, FilialRet.Text)

        Case 207615
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA2", gErr, iCodigo, ClienteRet.Text)
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, ClienteRet.Text)
'            If vbMsgRes = vbYes Then
'                Call Chama_Tela("FiliaisClientes", objFilialCliente)
'            End If

        Case 207616
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 207618
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, iCodigo, FornecedorRet.Text)
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, FornecedorRet.Text)
'
'            If vbMsgRes = vbYes Then
'                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
'            End If

        Case 207620
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", gErr, FilialRet.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207620)

    End Select

    Exit Sub

End Sub

Public Sub FilialEnt_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialCliente As New ClassFilialCliente
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sCliente As String
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_FilialEnt_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(FilialEnt.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If FilialEnt.Text = FilialEnt.List(FilialEnt.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialEnt, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 207621

    If OptionClienteEnt.Value = True Then

        'Se nao encontra o item com o código informado
        If lErro = 6730 Then
    
            'Verifica de o Cliente foi digitado
            If Len(Trim(ClienteEnt.Text)) = 0 Then gError 207622
    
            sCliente = ClienteEnt.Text
    
            objFilialCliente.iCodFilial = iCodigo
    
            'Pesquisa se existe filial com o codigo extraido
            lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
            If lErro <> SUCESSO And lErro <> 17660 Then gError 207623
    
            If lErro = 17660 Then gError 207624
    
            'Coloca na tela
            FilialEnt.Text = iCodigo & SEPARADOR & objFilialCliente.sNome
    
            lErro = Trata_FilialClienteEnt()
            If lErro <> SUCESSO Then gError 207625
    
        End If

        'Não encontrou valor informado que era STRING
        If lErro = 6731 Then gError 207626

    Else
    
        'Se nao encontra o ítem com o código informado
        If lErro = 6730 Then
    
            'Verifica de o fornecedor foi digitado
            If Len(Trim(FornecedorEnt.Text)) = 0 Then gError 207627
    
            sFornecedor = FornecedorEnt.Text
    
            objFilialFornecedor.iCodFilial = iCodigo
    
            'Pesquisa se existe filial com o codigo extraido
            lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 18272 Then gError 207628
    
            If lErro = 18272 Then gError 207629
    
            'coloca na tela
            FilialEnt.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome
    
            lErro = Trata_FilialFornEnt()
            If lErro <> SUCESSO Then gError 207630
    
    
        End If
    
        'Não encontrou valor informado que era STRING
        If lErro = 6731 Then gError 207631
    
    
    
    End If

    Exit Sub

Erro_FilialEnt_Validate:

    Cancel = True


    Select Case gErr

        Case 207621, 207623, 207625, 207628, 20630

        Case 207622
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 207624
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA2", gErr, iCodigo, ClienteRet.Text)
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, ClienteEnt.Text)
'            If vbMsgRes = vbYes Then
'                Call Chama_Tela("FiliaisClientes", objFilialCliente)
'            End If

        Case 207626
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, FilialEnt.Text)

        Case 207627
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 207629
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, iCodigo, FornecedorRet.Text)
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, FornecedorEnt.Text)
'
'            If vbMsgRes = vbYes Then
'                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
'            End If

        Case 207631
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", gErr, FilialEnt.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207632)
    End Select

    Exit Sub

End Sub

Public Sub FilialRet_Click()

Dim lErro As Long

On Error GoTo Erro_FilialRet_Click

    'Verifica se algo foi selecionada
    If FilialRet.ListIndex = -1 Then Exit Sub

    If OptionClienteRet.Value = True Then

    'Faz o tratamento da Filial selecionada
    lErro = Trata_FilialClienteRet()
    If lErro <> SUCESSO Then gError 207633

    Else
    
    'Faz o tratamento da Filial selecionada
    lErro = Trata_FilialFornRet()
    If lErro <> SUCESSO Then gError 207756
    
    End If

    Exit Sub

Erro_FilialRet_Click:

    Select Case gErr

        Case 207633

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207634)

    End Select

    Exit Sub

End Sub

Public Sub FilialEnt_Click()

Dim lErro As Long

On Error GoTo Erro_FilialEnt_Click


    'Verifica se algo foi selecionada
    If FilialEnt.ListIndex = -1 Then Exit Sub


    If OptionClienteEnt.Value = True Then

    'Faz o tratamento da Filial selecionada
    lErro = Trata_FilialClienteEnt()
    If lErro <> SUCESSO Then gError 207635

    Else
    
    'Faz o tratamento da Filial selecionada
    lErro = Trata_FilialFornEnt()
    If lErro <> SUCESSO Then gError 207635
    
    End If
    
    Exit Sub

Erro_FilialEnt_Click:

    Select Case gErr

        Case 207635

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207636)

    End Select

    Exit Sub

End Sub


Private Function Trata_FilialClienteRet() As Long

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente, objcliente As New ClassCliente
Dim objEndereco As New ClassEndereco
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Trata_FilialClienteRet

    If giCarregando = 0 Then

        objFilialCliente.iCodFilial = Codigo_Extrai(FilialRet.Text)
        'Lê a FilialCliente
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", ClienteRet.Text, objFilialCliente)
        If lErro <> SUCESSO Then gError 207637
    
        lErro = ERRO_LEITURA_SEM_DADOS
    
        objEndereco.lCodigo = objFilialCliente.lEnderecoEntrega
    
        lErro = CF("Endereco_Le", objEndereco)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 207752
            
        If Len(objEndereco.sLogradouro) = 0 Then
        
            objEndereco.lCodigo = objFilialCliente.lEndereco
        
            lErro = CF("Endereco_Le", objEndereco)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 207755
        
        End If
    
        If lErro = SUCESSO Then
            
            vbMsgRes = vbYes
            If Len(Trim(LogradouroRet.Text)) <> 0 Then
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_SUBSTITUIR_ENDERECO_ATUAL1")
            End If
            
            If vbMsgRes = vbYes Then
            
                BairroRet.Text = objEndereco.sBairro
                CidadeRet.Text = objEndereco.sCidade
                EstadoRet.Text = objEndereco.sSiglaEstado
                Call EstadoRet_Validate(bSGECancelDummy)
                PaisRet.Text = PAIS_BRASIL
                PaisRet_Validate (bSGECancelDummy)
                TipoLogradouroRet.Text = objEndereco.sTipoLogradouro
                Call TipoLogradouroRet_Validate(bSGECancelDummy)
                LogradouroRet.Text = objEndereco.sLogradouro
                If objEndereco.lNumero <> 0 Then
                    NumeroRet.Text = objEndereco.lNumero
                Else
                    NumeroRet.Text = ""
                End If
                ComplementoRet.Text = objEndereco.sComplemento
                CEPRet.Text = objEndereco.sCEP
                
                bMudouCEPRet = True
            
            End If
        
        End If
    
    
        CNPJCPFRet.Text = objFilialCliente.sCgc
        Call CNPJCPFRet_Validate(bSGECancelDummy)

    End If
    
    Trata_FilialClienteRet = SUCESSO

    Exit Function

Erro_Trata_FilialClienteRet:

    Trata_FilialClienteRet = gErr

    Select Case gErr

        Case 207637, 207752, 207755

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207638)
            
    End Select

    Exit Function

End Function

Private Function Trata_FilialClienteEnt() As Long

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente, objcliente As New ClassCliente
Dim objEndereco As New ClassEndereco
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Trata_FilialClienteEnt

    If giCarregando = 0 Then

        objFilialCliente.iCodFilial = Codigo_Extrai(FilialEnt.Text)
        'Lê a FilialCliente
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", ClienteEnt.Text, objFilialCliente)
        If lErro <> SUCESSO Then gError 207639
    
        lErro = ERRO_LEITURA_SEM_DADOS
    
        objEndereco.lCodigo = objFilialCliente.lEnderecoEntrega
    
        lErro = CF("Endereco_Le", objEndereco)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 207753
            
        If Len(objEndereco.sLogradouro) = 0 Then
        
            objEndereco.lCodigo = objFilialCliente.lEndereco
        
            lErro = CF("Endereco_Le", objEndereco)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 207754
        
        End If
        
        If lErro = SUCESSO Then
            
            vbMsgRes = vbYes
            If Len(Trim(LogradouroEnt.Text)) <> 0 Then
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_SUBSTITUIR_ENDERECO_ATUAL1")
            End If
            
            If vbMsgRes = vbYes Then
            
                BairroEnt.Text = objEndereco.sBairro
                CidadeEnt.Text = objEndereco.sCidade
                EstadoEnt.Text = objEndereco.sSiglaEstado
                Call EstadoEnt_Validate(bSGECancelDummy)
                PaisEnt.Text = PAIS_BRASIL
                PaisEnt_Validate (bSGECancelDummy)
                TipoLogradouroEnt.Text = objEndereco.sTipoLogradouro
                Call TipoLogradouroEnt_Validate(bSGECancelDummy)
                LogradouroEnt.Text = objEndereco.sLogradouro
                If objEndereco.lNumero <> 0 Then
                    NumeroEnt.Text = objEndereco.lNumero
                Else
                    NumeroEnt.Text = ""
                End If
                ComplementoEnt.Text = objEndereco.sComplemento
                CEPEnt.Text = objEndereco.sCEP
                
                bMudouCEPEnt = True
            
            End If
        
        End If
    
    
        CNPJCPFEnt.Text = objFilialCliente.sCgc
        Call CNPJCPFEnt_Validate(bSGECancelDummy)

    End If
    
    Trata_FilialClienteEnt = SUCESSO

    Exit Function

Erro_Trata_FilialClienteEnt:

    Trata_FilialClienteEnt = gErr

    Select Case gErr

        Case 207639, 207753, 207754

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207640)
            
    End Select

    Exit Function

End Function

Private Function Trata_FilialFornRet() As Long

Dim lErro As Long
Dim objFilialForn As New ClassFilialFornecedor
Dim objEndereco As New ClassEndereco
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Trata_FilialFornRet
    
    If giCarregando = 0 Then
    
        objFilialForn.iCodFilial = Codigo_Extrai(FilialRet.Text)
        
        If objFilialForn.iCodFilial <> 0 Then
        
            'Lê a Filial
            lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", FornecedorRet.Text, objFilialForn)
            If lErro <> SUCESSO Then gError 207641
            
            lErro = ERRO_LEITURA_SEM_DADOS
        
            If objFilialForn.lEndereco <> 0 Then
            
                objEndereco.lCodigo = objFilialForn.lEndereco
            
                lErro = CF("Endereco_Le", objEndereco)
                If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 207752
            
            End If
            
            If lErro = SUCESSO Then
                
                vbMsgRes = vbYes
                If Len(Trim(LogradouroRet.Text)) <> 0 Then
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_SUBSTITUIR_ENDERECO_ATUAL1")
                End If
                
                If vbMsgRes = vbYes Then
                
                    BairroRet.Text = objEndereco.sBairro
                    CidadeRet.Text = objEndereco.sCidade
                    EstadoRet.Text = objEndereco.sSiglaEstado
                    Call EstadoRet_Validate(bSGECancelDummy)
                    PaisRet.Text = PAIS_BRASIL
                    PaisRet_Validate (bSGECancelDummy)
                    TipoLogradouroRet.Text = objEndereco.sTipoLogradouro
                    Call TipoLogradouroRet_Validate(bSGECancelDummy)
                    LogradouroRet.Text = objEndereco.sLogradouro
                    If objEndereco.lNumero <> 0 Then
                        NumeroRet.Text = objEndereco.lNumero
                    Else
                        NumeroRet.Text = ""
                    End If
                    ComplementoRet.Text = objEndereco.sComplemento
                    CEPRet.Text = objEndereco.sCEP
                    
                    bMudouCEPRet = True
                
                End If
            
            End If
            
            
            
            CNPJCPFRet.Text = objFilialForn.sCgc
            Call CNPJCPFRet_Validate(bSGECancelDummy)
            
            
        End If
    
    End If
    
    Trata_FilialFornRet = SUCESSO
     
    Exit Function
    
Erro_Trata_FilialFornRet:

    Trata_FilialFornRet = gErr
     
    Select Case gErr
          
        Case 207641
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207642)
     
    End Select
     
    Exit Function

End Function

Private Function Trata_FilialFornEnt() As Long

Dim lErro As Long
Dim objFilialForn As New ClassFilialFornecedor
Dim objEndereco As New ClassEndereco
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Trata_FilialFornEnt
    
    If giCarregando = 0 Then
    
        objFilialForn.iCodFilial = Codigo_Extrai(FilialEnt.Text)
        
        If objFilialForn.iCodFilial <> 0 Then
        
            'Lê a Filial
            lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", FornecedorEnt.Text, objFilialForn)
            If lErro <> SUCESSO Then gError 207643
            
            lErro = ERRO_LEITURA_SEM_DADOS
        
            If objFilialForn.lEndereco <> 0 Then
            
                objEndereco.lCodigo = objFilialForn.lEndereco
            
                lErro = CF("Endereco_Le", objEndereco)
                If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 207753
                
            End If
            
            If lErro = SUCESSO Then
                
                vbMsgRes = vbYes
                If Len(Trim(LogradouroEnt.Text)) <> 0 Then
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_SUBSTITUIR_ENDERECO_ATUAL1")
                End If
                
                If vbMsgRes = vbYes Then
                
                    BairroEnt.Text = objEndereco.sBairro
                    CidadeEnt.Text = objEndereco.sCidade
                    EstadoEnt.Text = objEndereco.sSiglaEstado
                    Call EstadoEnt_Validate(bSGECancelDummy)
                    PaisEnt.Text = PAIS_BRASIL
                    PaisEnt_Validate (bSGECancelDummy)
                    TipoLogradouroEnt.Text = objEndereco.sTipoLogradouro
                    Call TipoLogradouroEnt_Validate(bSGECancelDummy)
                    LogradouroEnt.Text = objEndereco.sLogradouro
                    If objEndereco.lNumero <> 0 Then
                        NumeroEnt.Text = objEndereco.lNumero
                    Else
                        NumeroEnt.Text = ""
                    End If
                    ComplementoEnt.Text = objEndereco.sComplemento
                    CEPEnt.Text = objEndereco.sCEP
                    
                    bMudouCEPEnt = True
                
                End If
            
            End If
            
        
            CNPJCPFEnt.Text = objFilialForn.sCgc
            Call CNPJCPFEnt_Validate(bSGECancelDummy)
            
            
        End If
    
    End If
    
    Trata_FilialFornEnt = SUCESSO
     
    Exit Function
    
Erro_Trata_FilialFornEnt:

    Trata_FilialFornEnt = gErr
     
    Select Case gErr
          
        Case 207643
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207644)
     
    End Select
     
    Exit Function

End Function


Public Sub CNPJCPFRet_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CNPJCPFRet_Validate
    
    'Se CGC/CPF não foi preenchido -- Exit Sub
    If Len(Trim(CNPJCPFRet.Text)) = 0 Then Exit Sub
    
    Select Case Len(Trim(CNPJCPFRet.Text))

        Case STRING_CPF 'CPF
            
            'Critica Cpf
            lErro = Cpf_Critica(CNPJCPFRet.Text)
            If lErro <> SUCESSO Then Error 12316
            
            'Formata e coloca na Tela
            CNPJCPFRet.Format = "000\.000\.000-00; ; ; "
            CNPJCPFRet.Text = CNPJCPFRet.Text

        Case STRING_CGC 'CGC
            
            'Critica CGC
            lErro = Cgc_Critica(CNPJCPFRet.Text)
            If lErro <> SUCESSO Then Error 12317
            
            'Formata e Coloca na Tela
            CNPJCPFRet.Format = "00\.000\.000\/0000-00; ; ; "
            CNPJCPFRet.Text = CNPJCPFRet.Text

        Case Else
                
            gError 207645

    End Select

    Exit Sub

Erro_CNPJCPFRet_Validate:

    Cancel = True


    Select Case gErr

        Case 207645
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC_CPF", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207646)

    End Select

    Exit Sub

End Sub

Public Sub CNPJCPFEnt_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CNPJCPFEnt_Validate
    
    'Se CGC/CPF não foi preenchido -- Exit Sub
    If Len(Trim(CNPJCPFEnt.Text)) = 0 Then Exit Sub
    
    Select Case Len(Trim(CNPJCPFEnt.Text))

        Case STRING_CPF 'CPF
            
            'Critica Cpf
            lErro = Cpf_Critica(CNPJCPFEnt.Text)
            If lErro <> SUCESSO Then Error 12316
            
            'Formata e coloca na Tela
            CNPJCPFEnt.Format = "000\.000\.000-00; ; ; "
            CNPJCPFEnt.Text = CNPJCPFEnt.Text

        Case STRING_CGC 'CGC
            
            'Critica CGC
            lErro = Cgc_Critica(CNPJCPFEnt.Text)
            If lErro <> SUCESSO Then Error 12317
            
            'Formata e Coloca na Tela
            CNPJCPFEnt.Format = "00\.000\.000\/0000-00; ; ; "
            CNPJCPFEnt.Text = CNPJCPFEnt.Text

        Case Else
                
            gError 207647
            
    End Select

    Exit Sub

Erro_CNPJCPFEnt_Validate:

    Cancel = True

    Select Case gErr

        Case 207647
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC_CPF", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207648)

    End Select


    Exit Sub

End Sub

Private Sub OptionClienteEnt_Click()
    ClienteLabelEnt.Visible = True
    ClienteEnt.Visible = True
    FornecedorLabelEnt.Visible = False
    FornecedorEnt.Visible = False
    
    iClienteEntAlterado = REGISTRO_ALTERADO
    Call ClienteEnt_Validate(bSGECancelDummy)
    
End Sub

Private Sub OptionClienteRet_Click()
    ClienteLabelRet.Visible = True
    ClienteRet.Visible = True
    FornecedorLabelRet.Visible = False
    FornecedorRet.Visible = False
    
    iClienteRetAlterado = REGISTRO_ALTERADO
    Call ClienteRet_Validate(bSGECancelDummy)
End Sub

Private Sub OptionFornecedorEnt_Click()
    ClienteLabelEnt.Visible = False
    ClienteEnt.Visible = False
    FornecedorLabelEnt.Visible = True
    FornecedorEnt.Visible = True
    
    iFornecedorEntAlterado = REGISTRO_ALTERADO
    Call FornecedorEnt_Validate(bSGECancelDummy)
    
End Sub

Private Sub OptionFornecedorRet_Click()
    ClienteLabelRet.Visible = False
    ClienteRet.Visible = False
    FornecedorLabelRet.Visible = True
    FornecedorRet.Visible = True
    
    iFornecedorRetAlterado = REGISTRO_ALTERADO
    Call FornecedorRet_Validate(bSGECancelDummy)
    
End Sub

Public Sub PaisRet_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objPais As New ClassPais

On Error GoTo Erro_PaisRet_Validate

    'Verifica se foi preenchida a Combo Pais
    If Len(Trim(PaisRet.Text)) <> 0 Then

    '    'Verifica se está preenchida com o ítem selecionado na ComboBox Pais
    '    If Controle("Pais", iIndice).Text = Controle("Pais", iIndice).List(Controle("Pais", iIndice).ListIndex) Then Exit Sub
    
        'Verifica se existe o ítem na List da Combo. Se existir seleciona.
        lErro = Combo_Seleciona(PaisRet, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 207649
    
        'Nao existe o item com o CODIGO na List da ComboBox
        If lErro = 6730 Then
    
            objPais.iCodigo = iCodigo
    
            'Tenta ler Pais com esse codigo no BD
            lErro = CF("Paises_Le", objPais)
            If lErro <> SUCESSO And lErro <> 47876 Then gError 207650
            
            If lErro <> SUCESSO Then gError 207651
    
            PaisRet.Text = CStr(iCodigo) & SEPARADOR & objPais.sNome
    
        End If
    
        'Nao existe o item com a STRING na List da ComboBox
        If lErro = 6731 Then gError 207652
        
        If Codigo_Extrai(PaisRet.Text) = PAIS_BRASIL Then
            EstadoRet.Enabled = True
            If EstadoRet.Text = "EX" Then EstadoRet.ListIndex = iIndexUF
        Else
            EstadoRet.Enabled = False
            EstadoRet.Text = "EX"
        End If
        
    End If
    
    Exit Sub

Erro_PaisRet_Validate:

    Cancel = True

    Select Case gErr
    
        Case 207649, 207650

        Case 207651
            Call Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_CADASTRADO1", gErr, objPais.iCodigo)
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PAIS", objPais.iCodigo)
'            If vbMsgRes = vbYes Then Call Chama_Tela("Paises", objPais)

        Case 207652
            Call Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_CADASTRADO1", gErr, Trim(PaisRet.Text))

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207653)

    End Select

    Exit Sub

End Sub

Public Sub PaisEnt_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objPais As New ClassPais

On Error GoTo Erro_PaisEnt_Validate

    'Verifica se foi preenchida a Combo Pais
    If Len(Trim(PaisEnt.Text)) <> 0 Then

    '    'Verifica se está preenchida com o ítem selecionado na ComboBox Pais
    '    If Controle("Pais", iIndice).Text = Controle("Pais", iIndice).List(Controle("Pais", iIndice).ListIndex) Then Exit Sub
    
        'Verifica se existe o ítem na List da Combo. Se existir seleciona.
        lErro = Combo_Seleciona(PaisEnt, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 207654
    
        'Nao existe o item com o CODIGO na List da ComboBox
        If lErro = 6730 Then
    
            objPais.iCodigo = iCodigo
    
            'Tenta ler Pais com esse codigo no BD
            lErro = CF("Paises_Le", objPais)
            If lErro <> SUCESSO And lErro <> 47876 Then gError 207655
            If lErro <> SUCESSO Then gError 207656
    
            PaisEnt.Text = CStr(iCodigo) & SEPARADOR & objPais.sNome
    
        End If
    
        'Nao existe o item com a STRING na List da ComboBox
        If lErro = 6731 Then gError 207657
        
        If Codigo_Extrai(PaisEnt.Text) = PAIS_BRASIL Then
            EstadoEnt.Enabled = True
            If EstadoEnt.Text = "EX" Then EstadoEnt.ListIndex = iIndexUF
        Else
            EstadoEnt.Enabled = False
            EstadoEnt.Text = "EX"
        End If
        
    End If
    
    Exit Sub

Erro_PaisEnt_Validate:

    Cancel = True

    Select Case gErr
    
        Case 207654, 207655

        Case 207656
            Call Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_CADASTRADO1", gErr, objPais.iCodigo)
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PAIS", objPais.iCodigo)
'            If vbMsgRes = vbYes Then Call Chama_Tela("Paises", objPais)

        Case 207657
            Call Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_CADASTRADO1", gErr, Trim(PaisEnt.Text))

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207658)

    End Select

    Exit Sub

End Sub

Public Sub EstadoRet_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EstadoRet_Validate

    'Verifica se foi preenchido o Estado
    If Len(Trim(EstadoRet.Text)) <> 0 Then

        'Verifica se está preenchida com o ítem selecionado na ComboBox Estado
        If EstadoRet.Text = EstadoRet.List(EstadoRet.ListIndex) Then Exit Sub
    
        'Verifica se existe o ítem na Combo Estado, se existir seleciona o item
        lErro = Combo_Item_Igual_CI(EstadoRet)
        If lErro <> SUCESSO And lErro <> 58583 Then gError 207659
    
        'Não existe o ítem na ComboBox Estado
        If lErro = 58583 And Codigo_Extrai(PaisRet.Text) = PAIS_BRASIL Then gError 207660
    
    End If
    
    Exit Sub

Erro_EstadoRet_Validate:

    Cancel = True

    Select Case gErr

        Case 207659

        Case 207660
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_CADASTRADO", gErr, EstadoRet.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207661)

    End Select

    Exit Sub

End Sub

Public Sub EstadoEnt_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EstadoEnt_Validate

    'Verifica se foi preenchido o Estado
    If Len(Trim(EstadoEnt.Text)) <> 0 Then

        'Verifica se está preenchida com o ítem selecionado na ComboBox Estado
        If EstadoEnt.Text = EstadoEnt.List(EstadoEnt.ListIndex) Then Exit Sub
    
        'Verifica se existe o ítem na Combo Estado, se existir seleciona o item
        lErro = Combo_Item_Igual_CI(EstadoEnt)
        If lErro <> SUCESSO And lErro <> 58583 Then gError 207662
    
        'Não existe o ítem na ComboBox Estado
        If lErro = 58583 And Codigo_Extrai(PaisEnt.Text) = PAIS_BRASIL Then gError 207663
    
    End If
    
    Exit Sub

Erro_EstadoEnt_Validate:

    Cancel = True

    Select Case gErr

        Case 207662

        Case 207663
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_CADASTRADO", gErr, EstadoEnt.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207664)

    End Select

    Exit Sub

End Sub

Public Sub TipoLogradouroRet_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndiceAux As Integer
Dim iIndice1 As Integer
Dim iIndice2 As Integer
Dim iIndice3 As Integer
Dim iIndice4 As Integer
Dim iTam As Integer
Dim iPos As Integer
Dim sValor As String
Dim sValorRed As String
Dim sValorLin As String

On Error GoTo Erro_TipoLogradouroRet_Validate

    'Verifica se foi preenchido o TipoLogradouro
    If Len(Trim(TipoLogradouroRet.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox TipoLogradouro
    If UCase(SCodigo_Extrai(TipoLogradouroRet.Text)) = UCase(SCodigo_Extrai(TipoLogradouroRet.List(TipoLogradouroRet.ListIndex))) Then Exit Sub

    'Verifica se existe o ítem na Combo TipoLogradouro, se existir seleciona o item
    lErro = Combo_Item_Igual_CI(TipoLogradouroRet)
    If lErro <> SUCESSO And lErro <> 58583 Then gError 207665
    
    sValor = UCase(SCodigo_Extrai(TipoLogradouroRet.Text))
    iTam = Len(sValor)
    iPos = InStr(1, sValor, " ")
    If iPos <> 0 Then
        sValorRed = left(sValor, iPos - 1)
    Else
        sValorRed = ""
    End If

    'Não existe o ítem na ComboBox TipoLogradouro
    If lErro = 58583 Then 'gError 202966
    
        iIndice1 = -1
        iIndice2 = -1
        iIndice3 = -1
        iIndice4 = -1
        For iIndiceAux = 0 To TipoLogradouroRet.ListCount - 1
        
            sValorLin = TipoLogradouroRet.List(iIndiceAux)
        
            If sValor = UCase(SCodigo_Extrai(sValorLin)) Then
                    iIndice1 = iIndiceAux
                Exit For
            End If
        
            If sValor = left(UCase(SCodigo_Extrai(sValorLin)), iTam) Then
                    If iIndice2 = -1 Then iIndice2 = iIndiceAux
            End If
            
            iPos = InStr(1, sValorLin, SEPARADOR)
            
            If sValor = Mid(UCase(sValorLin), iPos + 1, iTam) Then
                    If iIndice3 = -1 Then iIndice3 = iIndiceAux
            End If

            If sValorRed = Mid(UCase(sValorLin), iPos + 1, iTam) Then
                    If iIndice4 = -1 Then iIndice4 = iIndiceAux
            End If
            
        Next
        
        If iIndice1 = -1 And iIndice2 = -1 And iIndice3 = -1 And iIndice4 = -1 Then
            gError 207666
        End If
    
    
        If iIndice1 <> -1 Then
            TipoLogradouroRet.ListIndex = iIndice1
        ElseIf iIndice2 <> -1 Then
            TipoLogradouroRet.ListIndex = iIndice2
        ElseIf iIndice3 <> -1 Then
            TipoLogradouroRet.ListIndex = iIndice3
        Else
            TipoLogradouroRet.ListIndex = iIndice4
        End If
    
    End If
    
    Exit Sub

Erro_TipoLogradouroRet_Validate:

    Cancel = True

    Select Case gErr

        Case 207665

        Case 207666
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOLOGRADOURO_NAO_CADASTRADO", gErr, TipoLogradouroRet.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207667)

    End Select

    Exit Sub

End Sub


Public Sub TipoLogradouroEnt_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndiceAux As Integer
Dim iIndice1 As Integer
Dim iIndice2 As Integer
Dim iIndice3 As Integer
Dim iIndice4 As Integer
Dim iTam As Integer
Dim iPos As Integer
Dim sValor As String
Dim sValorRed As String
Dim sValorLin As String

On Error GoTo Erro_TipoLogradouroEnt_Validate

    'Verifica se foi preenchido o TipoLogradouro
    If Len(Trim(TipoLogradouroEnt.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox TipoLogradouro
    If UCase(SCodigo_Extrai(TipoLogradouroEnt.Text)) = UCase(SCodigo_Extrai(TipoLogradouroEnt.List(TipoLogradouroEnt.ListIndex))) Then Exit Sub

    'Verifica se existe o ítem na Combo TipoLogradouro, se existir seleciona o item
    lErro = Combo_Item_Igual_CI(TipoLogradouroEnt)
    If lErro <> SUCESSO And lErro <> 58583 Then gError 207668
    
    sValor = UCase(SCodigo_Extrai(TipoLogradouroEnt.Text))
    iTam = Len(sValor)
    iPos = InStr(1, sValor, " ")
    If iPos <> 0 Then
        sValorRed = left(sValor, iPos - 1)
    Else
        sValorRed = ""
    End If

    'Não existe o ítem na ComboBox TipoLogradouro
    If lErro = 58583 Then 'gError 202966
    
        iIndice1 = -1
        iIndice2 = -1
        iIndice3 = -1
        iIndice4 = -1
        For iIndiceAux = 0 To TipoLogradouroEnt.ListCount - 1
        
            sValorLin = TipoLogradouroEnt.List(iIndiceAux)
        
            If sValor = UCase(SCodigo_Extrai(sValorLin)) Then
                    iIndice1 = iIndiceAux
                Exit For
            End If
        
            If sValor = left(UCase(SCodigo_Extrai(sValorLin)), iTam) Then
                    If iIndice2 = -1 Then iIndice2 = iIndiceAux
            End If
            
            iPos = InStr(1, sValorLin, SEPARADOR)
            
            If sValor = Mid(UCase(sValorLin), iPos + 1, iTam) Then
                    If iIndice3 = -1 Then iIndice3 = iIndiceAux
            End If

            If sValorRed = Mid(UCase(sValorLin), iPos + 1, iTam) Then
                    If iIndice4 = -1 Then iIndice4 = iIndiceAux
            End If
            
        Next
        
        If iIndice1 = -1 And iIndice2 = -1 And iIndice3 = -1 And iIndice4 = -1 Then
            gError 207669
        End If
    
    
        If iIndice1 <> -1 Then
            TipoLogradouroEnt.ListIndex = iIndice1
        ElseIf iIndice2 <> -1 Then
            TipoLogradouroEnt.ListIndex = iIndice2
        ElseIf iIndice3 <> -1 Then
            TipoLogradouroEnt.ListIndex = iIndice3
        Else
            TipoLogradouroEnt.ListIndex = iIndice4
        End If
    
    End If
    
    Exit Sub

Erro_TipoLogradouroEnt_Validate:

    Cancel = True

    Select Case gErr

        Case 207668

        Case 207669
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOLOGRADOURO_NAO_CADASTRADO", gErr, TipoLogradouroEnt.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207670)

    End Select

    Exit Sub

End Sub

Private Function FilialEmpresa_SetaEstados(objControle As Object)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim objEndereco As New ClassEndereco
Dim iSigla As Integer
Dim iIndice As Integer
Dim objTab As Object

On Error GoTo Erro_FilialEmpresa_SetaEstados

    If giFilialEmpresa <> EMPRESA_TODA And iIndexUF = 0 Then

        objFilialEmpresa.iCodFilial = giFilialEmpresa
        'Lê a Filial Empresa
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 207671

        'Se não encontrou a Filial da Empresa --> Erro
        If lErro <> SUCESSO Then gError 207672

        Set objEndereco = objFilialEmpresa.objEndereco

        If objEndereco.sSiglaEstado <> "" Then

            iSigla = Len(Trim(objEndereco.sSiglaEstado))
            'Seleciona o Estado "default" p/ a Filial se existir
            For iIndice = 0 To objControle.ListCount - 1
                If UCase(right(objControle.List(iIndice), iSigla)) = UCase(objEndereco.sSiglaEstado) Then
                    iIndexUF = iIndice
                    Exit For
                End If
            Next

        End If

    End If
    
    objControle.ListIndex = iIndexUF

    FilialEmpresa_SetaEstados = SUCESSO

    Exit Function

Erro_FilialEmpresa_SetaEstados:

    FilialEmpresa_SetaEstados = gErr

    Select Case gErr

        Case 207671

        Case 207672
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207673)

    End Select

    Exit Function

End Function

Public Sub BotaoLimpaEnt_Click()

Dim objTab As Object
Dim iIndice As Integer
Dim objControle As Object

On Error GoTo Erro_BotaoLimpaEnt_Click
    
    CEPEnt.Text = ""
    FornecedorEnt.Text = ""
    ClienteEnt.Text = ""
    CidadeEnt.Text = ""
    BairroEnt.Text = ""
    LogradouroEnt.Text = ""
    ComplementoEnt.Text = ""
    NumeroEnt.Text = ""
    CNPJCPFEnt.Text = ""
    
    
    
    'Seleciona Brasil nas Combos País
    FilialEnt.ListIndex = -1
    PaisEnt.ListIndex = iIndexBrasil
    
    TipoLogradouroEnt.ListIndex = -1
    
    Set objControle = EstadoEnt
    
    Call FilialEmpresa_SetaEstados(objControle)
    sCEPEnt = ""
    bMudouCEPEnt = False
    
    Exit Sub

Erro_BotaoLimpaEnt_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207674)

    End Select

    Exit Sub
End Sub

Public Sub BotaoLimpaRet_Click()

Dim objTab As Object
Dim iIndice As Integer
Dim objControle As Object

On Error GoTo Erro_BotaoLimpaRet_Click
    
    CEPRet.Text = ""
    FornecedorRet.Text = ""
    ClienteRet.Text = ""
    CidadeRet.Text = ""
    BairroRet.Text = ""
    LogradouroRet.Text = ""
    ComplementoRet.Text = ""
    NumeroRet.Text = ""
    CNPJCPFRet.Text = ""
    
    
    
    'Seleciona Brasil nas Combos País
    FilialRet.ListIndex = -1
    PaisRet.ListIndex = iIndexBrasil
    
    TipoLogradouroRet.ListIndex = -1
    
    Set objControle = EstadoRet
    
    Call FilialEmpresa_SetaEstados(objControle)
    sCEPRet = ""
    bMudouCEPRet = False
    
    Exit Sub

Erro_BotaoLimpaRet_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207675)

    End Select

    Exit Sub
End Sub

Public Sub BairroEnt_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub BairroRet_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CEPEnt_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CEPRet_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CEPEnt_GotFocus()
    Call MaskEdBox_TrataGotFocus(CEPEnt, iAlterado)
End Sub

Public Sub CEPRet_GotFocus()
    Call MaskEdBox_TrataGotFocus(CEPRet, iAlterado)
End Sub

Public Sub CidadeEnt_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CidadeRet_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub LogradouroEnt_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub LogradouroRet_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TipoLogradouroEnt_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TipoLogradouroRet_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TipoLogradouroEnt_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TipoLogradouroRet_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NumeroEnt_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NumeroRet_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NumeroEnt_GotFocus()
    Call MaskEdBox_TrataGotFocus(NumeroEnt, iAlterado)
End Sub

Public Sub NumeroRet_GotFocus()
    Call MaskEdBox_TrataGotFocus(NumeroRet, iAlterado)
End Sub

Public Sub ComplementoEnt_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ComplementoRet_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub EstadoEnt_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub EstadoRet_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub EstadoEnt_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub EstadoRet_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoPaisEnt_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPais As New ClassPais

On Error GoTo Erro_objEventoPaisEnt_evSelecao

    Set objPais = obj1

    PaisEnt.Text = CStr(objPais.iCodigo)
    PaisEnt_Validate (bSGECancelDummy)

    Me.Show

    Exit Sub

Erro_objEventoPaisEnt_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207682)

    End Select

    Exit Sub

End Sub

Private Sub objEventoPaisRet_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPais As New ClassPais

On Error GoTo Erro_objEventoPaisRet_evSelecao

    Set objPais = obj1

    PaisRet.Text = CStr(objPais.iCodigo)
    PaisRet_Validate (bSGECancelDummy)

    Me.Show

    Exit Sub

Erro_objEventoPaisRet_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207683)

    End Select

    Exit Sub

End Sub

Public Sub PaisEnt_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PaisRet_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PaisEnt_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PaisRet_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PaisLabelEnt_Click()

Dim objPais As New ClassPais
Dim colSelecao As Collection

    objPais.iCodigo = Codigo_Extrai(PaisEnt.Text)

    'Chama a Tela de PaisesLista
    Call Chama_Tela_Modal("PaisesLista", colSelecao, objPais, objEventoPaisEnt)
        

End Sub

Public Sub PaisLabelRet_Click()

Dim objPais As New ClassPais
Dim colSelecao As Collection

    objPais.iCodigo = Codigo_Extrai(PaisRet.Text)

    'Chama a Tela de PaisesLista
    Call Chama_Tela_Modal("PaisesLista", colSelecao, objPais, objEventoPaisRet)
        

End Sub

Public Sub LabelCidadeEnt_Click()

Dim objCidade As New ClassCidades
Dim colSelecao As Collection

    objCidade.sDescricao = CidadeEnt.Text

    'Chama a Tela de browse
    Call Chama_Tela_Modal("CidadeLista", colSelecao, objCidade, objEventoCidadeEnt)
        
End Sub

Public Sub LabelCidadeRet_Click()

Dim objCidade As New ClassCidades
Dim colSelecao As Collection

    objCidade.sDescricao = CidadeRet.Text

    'Chama a Tela de browse
    Call Chama_Tela_Modal("CidadeLista", colSelecao, objCidade, objEventoCidadeRet)
        
End Sub

Private Sub objEventoCidadeEnt_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCidade As ClassCidades

On Error GoTo Erro_objEventoCidadeEnt_evSelecao

    Set objCidade = obj1

    CidadeEnt.Text = objCidade.sDescricao

    Me.Show

    Exit Sub

Erro_objEventoCidadeEnt_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207694)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCidadeRet_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCidade As ClassCidades

On Error GoTo Erro_objEventoCidadeRet_evSelecao

    Set objCidade = obj1

    CidadeRet.Text = objCidade.sDescricao

    Me.Show

    Exit Sub

Erro_objEventoCidadeRet_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207695)

    End Select

    Exit Sub

End Sub

Public Sub CidadeEnt_Validate(Cancel As Boolean)

Dim lErro As Long, objCidade As New ClassCidades
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_CidadeEnt_Validate

    If Len(Trim(CidadeEnt.Text)) = 0 Then Exit Sub

    objCidade.sDescricao = CidadeEnt.Text
    
    lErro = CF("Cidade_Le_Nome", objCidade)
    If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 207696

    If lErro <> SUCESSO Then gError 207697

    Exit Sub

Erro_CidadeEnt_Validate:

    Cancel = True

    Select Case gErr

        Case 207696

        Case 207697
            Call Rotina_Erro(vbOKOnly, "ERRO_CIDADE_NAO_CADASTRADA2", gErr, CidadeEnt.Text)
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CIDADE")
'            If vbMsgRes = vbYes Then
'                Call Chama_Tela("CidadeCadastro", objCidade)
'            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207698)

    End Select

    Exit Sub

End Sub

Public Sub CidadeRet_Validate(Cancel As Boolean)

Dim lErro As Long, objCidade As New ClassCidades
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_CidadeRet_Validate

    If Len(Trim(CidadeRet.Text)) = 0 Then Exit Sub

    objCidade.sDescricao = CidadeRet.Text
    
    lErro = CF("Cidade_Le_Nome", objCidade)
    If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 207699

    If lErro <> SUCESSO Then gError 207700

    Exit Sub

Erro_CidadeRet_Validate:

    Cancel = True

    Select Case gErr

        Case 207699

        Case 207700
            Call Rotina_Erro(vbOKOnly, "ERRO_CIDADE_NAO_CADASTRADA2", gErr, CidadeRet.Text)
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CIDADE")
'            If vbMsgRes = vbYes Then
'                Call Chama_Tela("CidadeCadastro", objCidade)
'            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207701)

    End Select

    Exit Sub

End Sub

Public Sub CEPEnt_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objEndereco As New ClassEndereco
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_CEPEnt_Validate

    If Len(Trim(CEPEnt.Text)) = 0 Then Exit Sub

    objEndereco.sCEP = CEPEnt.Text
    
    lErro = CF("Endereco_Le_CEP", objEndereco)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 207702
    
    If lErro = SUCESSO Then
        
        vbMsgRes = vbYes
        If Len(Trim(LogradouroEnt.Text)) <> 0 And sCEPEnt <> CEPEnt.Text Then
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_SUBSTITUIR_ENDERECO_ATUAL")
        End If
        
        If vbMsgRes = vbYes And sCEPEnt <> CEPEnt.Text Then
        
            BairroEnt.Text = objEndereco.sBairro
            CidadeEnt.Text = objEndereco.sCidade
            EstadoEnt.Text = objEndereco.sSiglaEstado
            Call EstadoEnt_Validate(bSGECancelDummy)
            PaisEnt.Text = PAIS_BRASIL
            PaisEnt_Validate (bSGECancelDummy)
            TipoLogradouroEnt.Text = objEndereco.sTipoLogradouro
            Call TipoLogradouroEnt_Validate(bSGECancelDummy)
            LogradouroEnt.Text = objEndereco.sLogradouro
            
            bMudouCEPEnt = True
        
        End If
    
    End If
    
    sCEPEnt = CEPEnt.Text

    Exit Sub

Erro_CEPEnt_Validate:

    Cancel = True

    Select Case gErr
    
        Case 207702
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207703)

    End Select

    Exit Sub

End Sub

Public Sub CEPRet_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objEndereco As New ClassEndereco
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_CEPRet_Validate

    If Len(Trim(CEPRet.Text)) = 0 Then Exit Sub

    objEndereco.sCEP = CEPRet.Text
    
    lErro = CF("Endereco_Le_CEP", objEndereco)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 207704
    
    If lErro = SUCESSO Then
        
        vbMsgRes = vbYes
        If Len(Trim(LogradouroRet.Text)) <> 0 And sCEPRet <> CEPRet.Text Then
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_SUBSTITUIR_ENDERECO_ATUAL")
        End If
        
        If vbMsgRes = vbYes And sCEPRet <> CEPRet.Text Then
        
            BairroRet.Text = objEndereco.sBairro
            CidadeRet.Text = objEndereco.sCidade
            EstadoRet.Text = objEndereco.sSiglaEstado
            Call EstadoRet_Validate(bSGECancelDummy)
            PaisRet.Text = PAIS_BRASIL
            PaisRet_Validate (bSGECancelDummy)
            TipoLogradouroRet.Text = objEndereco.sTipoLogradouro
            Call TipoLogradouroRet_Validate(bSGECancelDummy)
            LogradouroRet.Text = objEndereco.sLogradouro
            
            bMudouCEPRet = True
        
        End If
    
    End If
    
    sCEPRet = CEPRet.Text

    Exit Sub

Erro_CEPRet_Validate:

    Cancel = True

    Select Case gErr
    
        Case 207704
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207705)

    End Select

    Exit Sub

End Sub

Public Sub CEPEnt_LostFocus()
    If bMudouCEPEnt Then Call CidadeEnt.SetFocus
End Sub

Public Sub CEPRet_LostFocus()
    If bMudouCEPRet Then Call CidadeRet.SetFocus
End Sub

Private Sub TabStrip1_Click()

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual
    If TabStrip1.SelectedItem.Index <> giFrameAtual Then

        If TabStrip_PodeTrocarTab(giFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Esconde o frame atual, mostra o novo
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        Frame1(giFrameAtual).Visible = False

        'Armazena novo valor de giFrameAtual
        giFrameAtual = TabStrip1.SelectedItem.Index
       
    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207706)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip2_Click()

On Error GoTo Erro_TabStrip2_Click

    'Se frame selecionado não for o atual
    If TabStrip2.SelectedItem.Index <> giFrameAtual2 Then

        If TabStrip_PodeTrocarTab(giFrameAtual2, TabStrip2, Me) <> SUCESSO Then Exit Sub

        'Esconde o frame atual, mostra o novo
        Frame2(TabStrip2.SelectedItem.Index).Visible = True
        Frame2(giFrameAtual2).Visible = False

        'Armazena novo valor de giFrameAtual
        giFrameAtual2 = TabStrip2.SelectedItem.Index
       
    End If

    Exit Sub

Erro_TabStrip2_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207706)

    End Select

    Exit Sub

End Sub

Public Sub ClienteEnt_Change()


    iAlterado = REGISTRO_ALTERADO
    iClienteEntAlterado = REGISTRO_ALTERADO

    Call ClienteEnt_Preenche

End Sub

Public Sub ClienteRet_Change()


    iAlterado = REGISTRO_ALTERADO
    iClienteRetAlterado = REGISTRO_ALTERADO

    Call ClienteRet_Preenche

End Sub

Private Sub ClienteEnt_Preenche()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objcliente As Object
    
On Error GoTo Erro_ClienteEnt_Preenche
    
    Set objcliente = ClienteEnt
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objcliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 207707

    Exit Sub

Erro_ClienteEnt_Preenche:

    Select Case gErr

        Case 207707

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207708)

    End Select
    
    Exit Sub

End Sub

Private Sub ClienteRet_Preenche()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objcliente As Object
    
On Error GoTo Erro_ClienteRet_Preenche
    
    Set objcliente = ClienteRet
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objcliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 207709

    Exit Sub

Erro_ClienteRet_Preenche:

    Select Case gErr

        Case 207709

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207710)

    End Select
    
    Exit Sub

End Sub

Public Sub FornecedorEnt_Change()

    iAlterado = REGISTRO_ALTERADO
    iFornecedorEntAlterado = REGISTRO_ALTERADO

    Call FornecedorEnt_Preenche

End Sub

Public Sub FornecedorRet_Change()

    iAlterado = REGISTRO_ALTERADO
    iFornecedorRetAlterado = REGISTRO_ALTERADO

    Call FornecedorRet_Preenche

End Sub

Private Sub FornecedorEnt_Preenche()
'por Jorge Specian - Para localizar pela parte digitada do Nome
'Reduzido do Fornecedor através da CF Fornecedor_Pesquisa_NomeReduzido em RotinasCPR.ClassCPRSelect'

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objFornecedor As Object
    
On Error GoTo Erro_FornecedorEnt_Preenche
    
    Set objFornecedor = FornecedorEnt
    
    lErro = CF("Fornecedor_Pesquisa_NomeReduzido", objFornecedor, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 207711

    Exit Sub

Erro_FornecedorEnt_Preenche:

    Select Case gErr

        Case 207711

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207712)

    End Select
    
    Exit Sub

End Sub

Private Sub FornecedorRet_Preenche()
'por Jorge Specian - Para localizar pela parte digitada do Nome
'Reduzido do Fornecedor através da CF Fornecedor_Pesquisa_NomeReduzido em RotinasCPR.ClassCPRSelect'

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objFornecedor As Object
    
On Error GoTo Erro_FornecedorRet_Preenche
    
    Set objFornecedor = FornecedorRet
    
    lErro = CF("Fornecedor_Pesquisa_NomeReduzido", objFornecedor, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 207713

    Exit Sub

Erro_FornecedorRet_Preenche:

    Select Case gErr

        Case 207713

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207714)

    End Select
    
    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Informações Adicionais"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RetiradaEntrega"
    
End Function

Public Sub Show()
'    Parent.Show
'    Parent.SetFocus
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

Public Sub ClienteLabelRet_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As Collection

    'Preenche NomeReduzido com o cliente da tela
    objcliente.sNomeReduzido = ClienteRet.Text

    'Chama Tela ClienteLista
    Call Chama_Tela_Modal("ClientesLista", colSelecao, objcliente, objEventoClienteRet)

End Sub

Private Sub objEventoClienteRet_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente
Dim bCancel As Boolean

    Set objcliente = obj1

    'Preenche campo Cliente
    ClienteRet.Text = objcliente.sNomeReduzido

    'Executa o Validate
    Call ClienteRet_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Private Sub ClienteLabelEnt_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As Collection

    'Preenche NomeReduzido com o cliente da tela
    objcliente.sNomeReduzido = ClienteEnt.Text

    'Chama Tela ClienteLista
    Call Chama_Tela_Modal("ClientesLista", colSelecao, objcliente, objEventoClienteEnt)

End Sub

Private Sub objEventoClienteEnt_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente
Dim bCancel As Boolean

    Set objcliente = obj1

    'Preenche campo Cliente
    ClienteEnt.Text = objcliente.sNomeReduzido

    'Executa o Validate
    Call ClienteEnt_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Public Sub FornecedorLabelEnt_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'recolhe o Nome Reduzido da tela
    objFornecedor.sNomeReduzido = FornecedorEnt.Text

    'Chama a Tela de browse Fornecedores
    Call Chama_Tela_Modal("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedorEnt)

    Exit Sub

End Sub

Public Sub objEventoFornecedorEnt_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor
Dim bCancel As Boolean

    Set objFornecedor = obj1

    'Coloca o Fornecedor na tela
    FornecedorEnt.Text = objFornecedor.lCodigo
    Call FornecedorEnt_Validate(bCancel)

    Me.Show

End Sub

Public Sub FornecedorLabelRet_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'recolhe o Nome Reduzido da tela
    objFornecedor.sNomeReduzido = FornecedorRet.Text

    'Chama a Tela de browse Fornecedores
    Call Chama_Tela_Modal("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedorRet)

    Exit Sub

End Sub

Public Sub objEventoFornecedorRet_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor
Dim bCancel As Boolean

    Set objFornecedor = obj1

    'Coloca o Fornecedor na tela
    FornecedorRet.Text = objFornecedor.lCodigo
    Call FornecedorRet_Validate(bCancel)

    Me.Show

End Sub

Public Sub UFEmbarque_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_UFEmbarque_Validate

    'Verifica se foi preenchido o Estado
    If Len(Trim(UFEmbarque.Text)) <> 0 Then

        'Verifica se está preenchida com o ítem selecionado na ComboBox Estado
        If UFEmbarque.Text = UFEmbarque.List(UFEmbarque.ListIndex) Then Exit Sub
    
        'Verifica se existe o ítem na Combo Estado, se existir seleciona o item
        lErro = Combo_Item_Igual_CI(UFEmbarque)
        If lErro <> SUCESSO And lErro <> 58583 Then gError 207659
    
        'Não existe o ítem na ComboBox Estado
        If lErro = 58583 Then gError 207660
    
    End If
    
    Exit Sub

Erro_UFEmbarque_Validate:

    Cancel = True

    Select Case gErr

        Case 207659

        Case 207660
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_CADASTRADO", gErr, UFEmbarque.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207661)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNatureza_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNatMovCta As ClassNatMovCta
Dim sNaturezaEnxuta As String

On Error GoTo Erro_objEventoNatureza_evSelecao

    Set objNatMovCta = obj1

    If objNatMovCta.sCodigo = "" Then
        
        Natureza.PromptInclude = False
        Natureza.Text = ""
        Natureza.PromptInclude = True
    
    Else

        sNaturezaEnxuta = String(STRING_NATMOVCTA_CODIGO, 0)
    
        lErro = Mascara_RetornaItemEnxuto(SEGMENTO_NATMOVCTA, objNatMovCta.sCodigo, sNaturezaEnxuta)
        If lErro <> SUCESSO Then gError 122833

        Natureza.PromptInclude = False
        Natureza.Text = sNaturezaEnxuta
        Natureza.PromptInclude = True
    
    End If

    Call Natureza_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

Erro_objEventoNatureza_evSelecao:

    Select Case gErr

        Case 122833

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub Natureza_Validate(Cancel As Boolean)
     
Dim lErro As Long
Dim sNaturezaFormatada As String
Dim iNaturezaPreenchida As Integer
Dim objNatMovCta As New ClassNatMovCta

On Error GoTo Erro_Natureza_Validate

    If Len(Natureza.ClipText) > 0 Then

        sNaturezaFormatada = String(STRING_NATMOVCTA_CODIGO, 0)

        'critica o formato da Natureza
        lErro = CF("Item_Formata", SEGMENTO_NATMOVCTA, Natureza.Text, sNaturezaFormatada, iNaturezaPreenchida)
        If lErro <> SUCESSO Then gError 122826
        
        'Obj recebe código
        objNatMovCta.sCodigo = sNaturezaFormatada
        
        'Verifica se a Natureza é analítica e se seu Tipo Corresponde a um pagamento
        lErro = CF("Natureza_Critica", objNatMovCta, NATUREZA_TIPO_INDEFINIDA)
        If lErro <> SUCESSO Then gError 122843
        
        'Coloca a Descrição da Natureza na Tela
        LabelNaturezaDesc.Caption = objNatMovCta.sDescricao
        
    Else
    
        LabelNaturezaDesc.Caption = ""
    
    End If
    
    Exit Sub
    
Erro_Natureza_Validate:

    Cancel = True

    Select Case gErr
    
        Case 122826, 122843
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
        
    End Select

    Exit Sub
    
End Sub


Private Sub LabelNatureza_Click()

    Dim objNatMovCta As New ClassNatMovCta
    Dim colSelecao As New Collection

    objNatMovCta.sCodigo = Natureza.ClipText
    
    'colSelecao.Add NATUREZA_TIPO_PAGAMENTO
    
    'Call Chama_Tela_Modal("NatMovCtaLista", colSelecao, objNatMovCta, objEventoNatureza, "Tipo = ?")
    
    Call Chama_Tela_Modal("NatMovCtaLista", colSelecao, objNatMovCta, objEventoNatureza, "")

End Sub

Private Sub Natureza_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Function Inicializa_Mascara_Natureza() As Long
'inicializa a mascara da Natureza

Dim sMascaraNatureza As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascara_Natureza

    'Inicializa a máscara da Natureza
    sMascaraNatureza = String(STRING_NATMOVCTA_CODIGO, 0)
    
    'Armazena em sMascaraNatureza a mascara a ser a ser exibida no campo Natureza
    lErro = MascaraItem(SEGMENTO_NATMOVCTA, sMascaraNatureza)
    If lErro <> SUCESSO Then gError 122836
    
    'coloca a mascara na tela.
    Natureza.Mask = sMascaraNatureza
    
    Inicializa_Mascara_Natureza = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascara_Natureza:

    Inicializa_Mascara_Natureza = gErr
    
    Select Case gErr
    
        Case 122836
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARAITEM", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
        
    End Select

    Exit Function

End Function

Private Sub BotaoProjetos_Click()
    Call gobjTela.gobjTelaProjetoInfo.BotaoProjetos_Click
End Sub

Private Sub LabelProjeto_Click()
    Call gobjTela.gobjTelaProjetoInfo.LabelProjetoInfoAdi_Click
End Sub

Private Sub Projeto_GotFocus()
    Call MaskEdBox_TrataGotFocus(Projeto, gobjTela.iAlterado)
End Sub

Private Sub Projeto_Change()
    gobjTela.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Projeto_Validate(Cancel As Boolean)
    Call gobjTela.gobjTelaProjetoInfo.ProjetoInfoAdi_Validate(Cancel)
End Sub

Sub Etapa_Change()
     gobjTela.iAlterado = REGISTRO_ALTERADO
End Sub

Sub Etapa_Click()
    gobjTela.iAlterado = REGISTRO_ALTERADO
End Sub

Sub Etapa_Validate(Cancel As Boolean)
    Call gobjTela.gobjTelaProjetoInfo.ProjetoInfoAdi_Validate(Cancel)
End Sub

Public Sub CclLabel_Click()

Dim objCcl As New ClassCcl
Dim colSelecao As New Collection

    Call Chama_Tela_Modal("CclLista", colSelecao, objCcl, objEventoCcl)

End Sub

Private Sub objEventoCcl_evSelecao(obj1 As Object)
'Preenche Ccl

Dim objCcl As New ClassCcl
Dim sCclFormatada As String
Dim sCclMascarado As String
Dim lErro As Long

On Error GoTo Erro_objEventoCcl_evSelecao

    Set objCcl = obj1

    sCclMascarado = String(STRING_CCL, 0)

    lErro = Mascara_RetornaCclEnxuta(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then gError 22930

    Ccl.PromptInclude = False
    Ccl.Text = sCclMascarado
    Ccl.PromptInclude = True
    
    DescCcl.Caption = objCcl.sDescCcl

    Me.Show

    Exit Sub

Erro_objEventoCcl_evSelecao:

    Select Case gErr

        Case 22930
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", gErr, objCcl.sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175259)

    End Select

    Exit Sub

End Sub

Private Sub Ccl_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Ccl_Validate(Cancel As Boolean)
'verifica existência da Ccl informada

Dim lErro As Long, sCclFormatada As String
Dim objCcl As New ClassCcl

On Error GoTo Erro_Ccl_Validate

    'se Ccl não estiver preenchida sai da rotina
    If Len(Trim(Ccl.Text)) <> 0 Then

        lErro = CF("Ccl_Critica", Ccl.Text, sCclFormatada, objCcl)
        If lErro <> SUCESSO And lErro <> 5703 Then gError 31558
    
        If lErro = 5703 Then gError 31559
        
        DescCcl.Caption = objCcl.sDescCcl
        
    Else
        DescCcl.Caption = ""
    End If

    Exit Sub

Erro_Ccl_Validate:

    Cancel = True

    Select Case gErr

        Case 31558

        Case 31559
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, Ccl.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175260)

    End Select

    Exit Sub

End Sub

Public Sub NumDE_GotFocus()
    Call MaskEdBox_TrataGotFocus(NumDE, iAlterado)
End Sub

Public Sub NumDE_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NumRE_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub LabelDE_Click()

Dim objDE As New ClassDEInfo
Dim colSelecao As Collection

    objDE.sNumero = NumDE.Text

    'Chama a Tela de PaisesLista
    Call Chama_Tela_Modal("DEInfoLista", colSelecao, objDE, objEventoDE)

End Sub

Private Sub objEventoDE_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objDE As ClassDEInfo

On Error GoTo Erro_objEventoDE_evSelecao

    Set objDE = obj1

    NumDE.PromptInclude = False
    NumDE.Text = objDE.sNumero
    NumDE.PromptInclude = True
    Call NumDE_Validate(bSGECancelDummy)
    
    Call CF("SCombo_Seleciona2", NumRE, objDE.sNumRegistro)

    Me.Show

    Exit Sub

Erro_objEventoDE_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 216042)

    End Select

    Exit Sub

End Sub

Private Sub NumDE_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objDE As New ClassDEInfo
Dim objRE As ClassDERegistro

On Error GoTo Erro_NumDE_Validate

    If sNumDEAnt <> NumDE.ClipText Then

        If Len(NumDE.ClipText) > 0 Then
    
            objDE.sNumero = NumDE.ClipText
            
            lErro = CF("DEInfo_Le", objDE)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
            
            If lErro = ERRO_LEITURA_SEM_DADOS Then gError 216040
            
            NumRE.Clear
            NumRE.AddItem ""
            For Each objRE In objDE.colRE
                NumRE.AddItem objRE.sNumRegistro
            Next
            
            If objDE.colRE.Count = 1 Then NumRE.ListIndex = 1
            
            If objDE.sLocalEmbarque <> "" Then LocalEmbarque.Text = objDE.sLocalEmbarque
            If objDE.sUFEmbarque <> "" Then
                UFEmbarque.Text = objDE.sUFEmbarque
                Call UFEmbarque_Validate(bSGECancelDummy)
            End If
        
        Else
        
            NumRE.Clear
        
        End If
        
        sNumDEAnt = NumDE.ClipText
        
    End If
    
    Exit Sub
    
Erro_NumDE_Validate:

    Cancel = True

    Select Case gErr
    
        Case 216040
            Call Rotina_Erro(vbOKOnly, "ERRO_DEINFO_NAO_CADASTRADO", gErr, objDE.sNumero)
    
        Case ERRO_SEM_MENSAGEM
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 216041)
        
    End Select

    Exit Sub
    
End Sub
