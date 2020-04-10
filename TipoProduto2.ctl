VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl TipoProduto 
   ClientHeight    =   5700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9090
   KeyPreview      =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   9090
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4905
      Index           =   4
      Left            =   150
      TabIndex        =   81
      Top             =   690
      Visible         =   0   'False
      Width           =   8775
      Begin VB.ComboBox OrigemMercadoria 
         Height          =   315
         ItemData        =   "TipoProduto2.ctx":0000
         Left            =   4200
         List            =   "TipoProduto2.ctx":001F
         Style           =   2  'Dropdown List
         TabIndex        =   115
         Top             =   705
         Width           =   4515
      End
      Begin VB.Frame Frame28 
         Caption         =   "Demais Informações"
         Height          =   1815
         Left            =   90
         TabIndex        =   94
         Top             =   3000
         Width           =   8670
         Begin VB.ComboBox ProdutoEspecifico 
            Height          =   315
            ItemData        =   "TipoProduto2.ctx":0386
            Left            =   2190
            List            =   "TipoProduto2.ctx":039A
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   78
            Top             =   255
            Width           =   2925
         End
         Begin MSMask.MaskEdBox Genero 
            Height          =   315
            Left            =   2190
            TabIndex        =   79
            Top             =   585
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   2
            Format          =   "00"
            Mask            =   "##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ISSQN 
            Height          =   315
            Left            =   2190
            TabIndex        =   80
            Top             =   1095
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Format          =   "0000"
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Produto específico:"
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
            Index           =   3
            Left            =   390
            TabIndex        =   101
            Top             =   300
            Width           =   1680
         End
         Begin VB.Label DescISSQN 
            BorderStyle     =   1  'Fixed Single
            Height          =   630
            Left            =   2850
            TabIndex        =   100
            Top             =   1095
            Width           =   5790
         End
         Begin VB.Label LabelISSQN 
            AutoSize        =   -1  'True
            Caption         =   "ISSQN:"
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
            Left            =   1425
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   99
            Top             =   1140
            Width           =   645
         End
         Begin VB.Label LabelGenero 
            AutoSize        =   -1  'True
            Caption         =   "Gênero:"
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
            Left            =   1395
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   98
            Top             =   645
            Width           =   675
         End
         Begin VB.Label DescGenero 
            BorderStyle     =   1  'Fixed Single
            Height          =   495
            Left            =   2580
            TabIndex        =   97
            Top             =   585
            Width           =   6045
         End
         Begin VB.Label CodServNFe 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2190
            TabIndex        =   96
            Top             =   1410
            Width           =   660
         End
         Begin VB.Label LabelCodServNFe 
            AutoSize        =   -1  'True
            Caption         =   "Código Serviço NFSe:"
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
            Left            =   180
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   95
            Top             =   1470
            Width           =   1875
         End
      End
      Begin VB.Frame Frame29 
         Caption         =   "NCM"
         Height          =   915
         Left            =   315
         TabIndex        =   86
         Top             =   645
         Width           =   3465
         Begin MSMask.MaskEdBox ClasFiscIPI 
            Height          =   300
            Left            =   1935
            TabIndex        =   73
            Top             =   165
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Format          =   "0000\.00\.00"
            Mask            =   "##########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CEST 
            Height          =   300
            Left            =   1935
            TabIndex        =   74
            Top             =   540
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   7
            Format          =   "00\.000\.00"
            Mask            =   "#######"
            PromptChar      =   " "
         End
         Begin VB.Label LabelCEST 
            AutoSize        =   -1  'True
            Caption         =   "CEST:"
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
            TabIndex        =   116
            Top             =   600
            Width           =   555
         End
         Begin VB.Label LabelClassificacaoFiscal 
            AutoSize        =   -1  'True
            Caption         =   "Classificação Fiscal:"
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
            Left            =   135
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   87
            Top             =   225
            Width           =   1755
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Contabilidade"
         Height          =   660
         Left            =   75
         TabIndex        =   82
         Top             =   2235
         Width           =   8670
         Begin MSMask.MaskEdBox ContaContabil 
            Height          =   315
            Left            =   2190
            TabIndex        =   76
            Top             =   240
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ContaProducao 
            Height          =   315
            Left            =   5925
            TabIndex        =   77
            Top             =   240
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label LabelContaProducao 
            AutoSize        =   -1  'True
            Caption         =   "Conta de Produção:"
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
            Left            =   4170
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   84
            ToolTipText     =   "Conta Contábil de Produção"
            Top             =   255
            Width           =   1725
         End
         Begin VB.Label ContaContabilLabel 
            AutoSize        =   -1  'True
            Caption         =   "Conta de Aplicação:"
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
            Left            =   390
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   83
            ToolTipText     =   "Conta Contábil de Aplicação"
            Top             =   270
            Width           =   1755
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "IPI"
         Height          =   1965
         Left            =   75
         TabIndex        =   88
         Top             =   210
         Width           =   4005
         Begin VB.CheckBox IncideIPI 
            Caption         =   "Incide"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Left            =   255
            TabIndex        =   71
            Top             =   195
            Value           =   1  'Checked
            Width           =   915
         End
         Begin VB.Frame Frame30 
            Caption         =   "Exceção a TIPI"
            Height          =   510
            Left            =   240
            TabIndex        =   89
            Top             =   1365
            Width           =   3465
            Begin MSMask.MaskEdBox ExTIPI 
               Height          =   300
               Left            =   1935
               TabIndex        =   75
               Top             =   150
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   3
               Mask            =   "###"
               PromptChar      =   " "
            End
            Begin VB.Label Label22 
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
               Left            =   1215
               TabIndex        =   90
               Top             =   195
               Width           =   660
            End
         End
         Begin MSMask.MaskEdBox AliquotaIPI 
            Height          =   285
            Left            =   2190
            TabIndex        =   72
            Top             =   180
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodigoIPI 
            Height          =   300
            Left            =   3870
            TabIndex        =   91
            Top             =   240
            Visible         =   0   'False
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   3
            PromptChar      =   "_"
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Alíquota:"
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
            Left            =   1320
            TabIndex        =   93
            Top             =   195
            Width           =   795
         End
         Begin VB.Label Label27 
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
            Left            =   3150
            TabIndex        =   92
            Top             =   300
            Visible         =   0   'False
            Width           =   660
         End
      End
      Begin VB.Label LabelOrigemMerc 
         Caption         =   "Origem da Mercadoria:"
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
         Left            =   4155
         TabIndex        =   114
         Top             =   420
         Width           =   2520
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4770
      Index           =   1
      Left            =   105
      TabIndex        =   20
      Top             =   675
      Width           =   8805
      Begin VB.Frame Frame19 
         Caption         =   "Controle de Reserva/Estoque"
         Height          =   735
         Left            =   120
         TabIndex        =   110
         Top             =   4035
         Width           =   4290
         Begin VB.ComboBox ControleEstoque 
            Height          =   315
            ItemData        =   "TipoProduto2.ctx":03DF
            Left            =   1290
            List            =   "TipoProduto2.ctx":03EC
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   285
            Width           =   2895
         End
         Begin VB.Label Label4 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   795
            TabIndex        =   111
            Top             =   315
            Width           =   450
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "Custo"
         Height          =   675
         Left            =   105
         TabIndex        =   107
         Top             =   3285
         Width           =   4305
         Begin VB.ComboBox ApropriacaoComp 
            Height          =   315
            ItemData        =   "TipoProduto2.ctx":041D
            Left            =   1305
            List            =   "TipoProduto2.ctx":0424
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   225
            Width           =   2610
         End
         Begin VB.ComboBox ApropriacaoProd 
            Height          =   315
            ItemData        =   "TipoProduto2.ctx":0435
            Left            =   1305
            List            =   "TipoProduto2.ctx":043F
            Style           =   2  'Dropdown List
            TabIndex        =   108
            Top             =   225
            Visible         =   0   'False
            Width           =   2610
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Apropriação:"
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
            Left            =   120
            TabIndex        =   109
            Top             =   255
            Width           =   1095
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Identificação"
         Height          =   1050
         Left            =   105
         TabIndex        =   102
         Top             =   150
         Width           =   8610
         Begin VB.ComboBox NaturezaProduto 
            Height          =   315
            ItemData        =   "TipoProduto2.ctx":046B
            Left            =   5460
            List            =   "TipoProduto2.ctx":049A
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   210
            Width           =   3075
         End
         Begin MSMask.MaskEdBox Descricao 
            Height          =   315
            Left            =   1305
            TabIndex        =   3
            Top             =   630
            Width           =   7170
            _ExtentX        =   12647
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   315
            Left            =   1305
            TabIndex        =   0
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
         Begin MSMask.MaskEdBox Sigla 
            Height          =   315
            Left            =   3270
            TabIndex        =   1
            Top             =   225
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   5
            PromptChar      =   " "
         End
         Begin VB.Label Label2 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   1
            Left            =   4530
            TabIndex        =   106
            Top             =   240
            Width           =   840
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Sigla:"
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
            Left            =   2715
            TabIndex        =   105
            Top             =   255
            Width           =   495
         End
         Begin VB.Label Label2 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   0
            Left            =   315
            TabIndex        =   104
            Top             =   675
            Width           =   930
         End
         Begin VB.Label LblTipo 
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
            Left            =   795
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   103
            Top             =   255
            Width           =   450
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Categorias"
         Height          =   3495
         Left            =   4515
         TabIndex        =   22
         Top             =   1275
         Width           =   4185
         Begin VB.ComboBox ComboCategoriaProdutoItem 
            Height          =   315
            Left            =   2175
            TabIndex        =   24
            Top             =   720
            Width           =   1632
         End
         Begin VB.ComboBox ComboCategoriaProduto 
            Height          =   315
            Left            =   570
            TabIndex        =   23
            Top             =   750
            Width           =   1548
         End
         Begin MSFlexGridLib.MSFlexGrid GridCategoria 
            Height          =   1875
            Left            =   150
            TabIndex        =   9
            Top             =   315
            Width           =   3435
            _ExtentX        =   6059
            _ExtentY        =   3307
            _Version        =   393216
            Rows            =   6
            Cols            =   3
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Características"
         Height          =   1860
         Left            =   120
         TabIndex        =   21
         Top             =   1275
         Width           =   4290
         Begin VB.ListBox ListaCaracteristicas 
            Height          =   1185
            ItemData        =   "TipoProduto2.ctx":058D
            Left            =   240
            List            =   "TipoProduto2.ctx":05A0
            Style           =   1  'Checkbox
            TabIndex        =   4
            Top             =   255
            Width           =   3855
         End
         Begin VB.OptionButton Comprado 
            Caption         =   "Comprado"
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
            Left            =   2610
            TabIndex        =   6
            Top             =   1530
            Value           =   -1  'True
            Width           =   1245
         End
         Begin VB.OptionButton Produzido 
            Caption         =   "Produzido"
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
            Left            =   405
            TabIndex        =   5
            Top             =   1500
            Width           =   1395
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4800
      Index           =   2
      Left            =   120
      TabIndex        =   25
      Top             =   690
      Visible         =   0   'False
      Width           =   8775
      Begin VB.Frame Frame16 
         Caption         =   "Unidade de Medida"
         Height          =   3675
         Left            =   1350
         TabIndex        =   26
         Top             =   420
         Width           =   6420
         Begin VB.Frame Frame17 
            Caption         =   "Unidade Padrão"
            Height          =   2400
            Left            =   615
            TabIndex        =   29
            Top             =   915
            Width           =   5235
            Begin VB.ComboBox SiglaUMTrib 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   19
               Top             =   1875
               Width           =   915
            End
            Begin VB.ComboBox SiglaUMVenda 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   18
               Top             =   1365
               Width           =   915
            End
            Begin VB.ComboBox SiglaUMCompra 
               Height          =   315
               Left            =   1335
               Style           =   2  'Dropdown List
               TabIndex        =   17
               Top             =   855
               Width           =   915
            End
            Begin VB.ComboBox SiglaUMEstoque 
               Height          =   315
               Left            =   1335
               Style           =   2  'Dropdown List
               TabIndex        =   16
               Top             =   360
               Width           =   915
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Tributável:"
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
               Left            =   345
               TabIndex        =   113
               Top             =   1920
               Width           =   915
            End
            Begin VB.Label NomeUMTrib 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   2460
               TabIndex        =   112
               Top             =   1875
               Width           =   2280
            End
            Begin VB.Label NomeUMVenda 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   2460
               TabIndex        =   35
               Top             =   1365
               Width           =   2280
            End
            Begin VB.Label LblUMVenda 
               AutoSize        =   -1  'True
               Caption         =   "Vendas:"
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
               Left            =   570
               TabIndex        =   34
               Top             =   1410
               Width           =   705
            End
            Begin VB.Label NomeUMCompra 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   2460
               TabIndex        =   33
               Top             =   855
               Width           =   2280
            End
            Begin VB.Label LblUMCompra 
               AutoSize        =   -1  'True
               Caption         =   "Compras:"
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
               Left            =   480
               TabIndex        =   32
               Top             =   900
               Width           =   795
            End
            Begin VB.Label NomeUMEstoque 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   2460
               TabIndex        =   31
               Top             =   360
               Width           =   2280
            End
            Begin VB.Label LblUMEstoque 
               AutoSize        =   -1  'True
               Caption         =   "Estoque:"
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
               Left            =   510
               TabIndex        =   30
               Top             =   360
               Width           =   765
            End
         End
         Begin MSMask.MaskEdBox ClasseUM 
            Height          =   315
            Left            =   1350
            TabIndex        =   15
            Top             =   405
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "9999"
            PromptChar      =   " "
         End
         Begin VB.Label LblClasseUM 
            AutoSize        =   -1  'True
            Caption         =   "Classe:"
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
            Left            =   660
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   27
            Top             =   435
            Width           =   630
         End
         Begin VB.Label DescricaoClasseUM 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1965
            TabIndex        =   28
            Top             =   405
            Width           =   3885
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4785
      Index           =   3
      Left            =   120
      TabIndex        =   36
      Top             =   690
      Visible         =   0   'False
      Width           =   8730
      Begin VB.Frame Frame12 
         Caption         =   "Recebimento"
         Height          =   2760
         Left            =   45
         TabIndex        =   46
         Top             =   1890
         Width           =   4290
         Begin VB.Frame Frame14 
            Caption         =   "Faixa de recebimento"
            Height          =   1095
            Left            =   315
            TabIndex        =   48
            Top             =   525
            Width           =   3585
            Begin MSMask.MaskEdBox PercentMaisReceb 
               Height          =   315
               Left            =   2355
               TabIndex        =   50
               Top             =   240
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox PercentMenosReceb 
               Height          =   315
               Left            =   2355
               TabIndex        =   52
               Top             =   660
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Porcentagem a menos:"
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
               Left            =   330
               TabIndex        =   51
               Top             =   720
               Width           =   1950
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Porcentagem a mais:"
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
               Left            =   510
               TabIndex        =   49
               Top             =   300
               Width           =   1785
            End
         End
         Begin VB.Frame Frame13 
            Caption         =   "Recebimento fora da faixa"
            Height          =   960
            Left            =   315
            TabIndex        =   53
            Top             =   1680
            Width           =   3585
            Begin VB.OptionButton RecebForaFaixa 
               Caption         =   "Não aceita recebimento"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   420
               TabIndex        =   54
               Top             =   315
               Value           =   -1  'True
               Width           =   2415
            End
            Begin VB.OptionButton RecebForaFaixa 
               Caption         =   "Avisa e aceita recebimento"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   1
               Left            =   420
               TabIndex        =   55
               Top             =   600
               Width           =   2655
            End
         End
         Begin VB.CheckBox NaoTemFaixaReceb 
            Caption         =   "Aceita qualquer quantidade sem aviso"
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
            Left            =   330
            TabIndex        =   47
            Top             =   255
            Width           =   3585
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Cotações Anteriores"
         Height          =   2070
         Left            =   4455
         TabIndex        =   63
         Top             =   1890
         Width           =   4095
         Begin VB.CheckBox NaoConsideraQuantCotacaoAnterior 
            Caption         =   "Usa independente de quantidade"
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
            Left            =   345
            TabIndex        =   64
            Top             =   360
            Width           =   3165
         End
         Begin VB.Frame Frame11 
            Caption         =   "Limites percentuais de quantidade para uso"
            Height          =   1185
            Index           =   0
            Left            =   270
            TabIndex        =   65
            Top             =   750
            Width           =   3525
            Begin MSMask.MaskEdBox PercentMaisQuantCotacaoAnterior 
               Height          =   315
               Left            =   2310
               TabIndex        =   67
               Top             =   300
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox PercentMenosQuantCotacaoAnterior 
               Height          =   315
               Left            =   2310
               TabIndex        =   69
               Top             =   735
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Percentagem a mais:"
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
               TabIndex        =   66
               Top             =   360
               Width           =   1785
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Percentagem a menos:"
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
               TabIndex        =   68
               Top             =   780
               Width           =   1950
            End
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Estoque de Segurança"
         Height          =   1080
         Left            =   45
         TabIndex        =   41
         Top             =   795
         Width           =   4290
         Begin MSMask.MaskEdBox ConsumoMedioMax 
            Height          =   315
            Left            =   3300
            TabIndex        =   43
            Top             =   240
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Format          =   "0\%"
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TempoRessupMax 
            Height          =   315
            Left            =   3300
            TabIndex        =   45
            Top             =   660
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Format          =   "0\%"
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "% a Mais Max de Consumo Médio:"
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
            TabIndex        =   42
            Top             =   285
            Width           =   2895
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "% a Mais Max de Tempo de Ressup:"
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
            TabIndex        =   44
            Top             =   735
            Width           =   3090
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ressuprimento"
         Height          =   765
         Left            =   4455
         TabIndex        =   56
         Top             =   -15
         Width           =   4095
         Begin MSMask.MaskEdBox IntRessup 
            Height          =   315
            Left            =   2910
            TabIndex        =   58
            Top             =   330
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Intervalo de Ressuprimento:"
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
            TabIndex        =   57
            Top             =   375
            Width           =   2400
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "dias"
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
            Left            =   3540
            TabIndex        =   59
            Top             =   375
            Width           =   360
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Pedido/Requisição de Compras"
         Height          =   780
         Left            =   4455
         TabIndex        =   60
         Top             =   795
         Width           =   4095
         Begin MSMask.MaskEdBox Residuo 
            Height          =   315
            Left            =   2415
            TabIndex        =   62
            Top             =   330
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   2
            Format          =   "0\%"
            Mask            =   "##"
            PromptChar      =   " "
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Resíduo (%):"
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
            Left            =   1200
            TabIndex        =   61
            Top             =   360
            Width           =   1110
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Consumo"
         Height          =   765
         Left            =   30
         TabIndex        =   37
         Top             =   -15
         Width           =   4305
         Begin MSMask.MaskEdBox MesesConsumoMedio 
            Height          =   315
            Left            =   2985
            TabIndex        =   39
            Top             =   330
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   556
            _Version        =   393216
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
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "meses"
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
            Left            =   3555
            TabIndex        =   40
            Top             =   390
            Width           =   540
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Faixa cálculo Consumo Médio:"
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
            Left            =   285
            TabIndex        =   38
            Top             =   405
            Width           =   2595
         End
      End
      Begin VB.CommandButton BotaoComprasConfig 
         Caption         =   "Traz dados default"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   5265
         TabIndex        =   70
         Top             =   4095
         Width           =   2700
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6840
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   70
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "TipoProduto2.ctx":0636
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1095
         Picture         =   "TipoProduto2.ctx":07B4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "TipoProduto2.ctx":0CE6
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "TipoProduto2.ctx":0E70
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5265
      Left            =   90
      TabIndex        =   10
      Top             =   360
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   9287
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Principais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Unidades de Medida"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Estoque/Compras"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
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
Attribute VB_Name = "TipoProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iFrameAtual As Integer
Dim iAlterado As Integer
Dim iClasseUMAlterada As Integer
Dim iClasseUMAnterior As Integer

Dim objGridCategoria As AdmGrid
Dim iGrid_Categoria_Col As Integer
Dim iGrid_Valor_Col As Integer

Private WithEvents objEventoClasseUM As AdmEvento
Attribute objEventoClasseUM.VB_VarHelpID = -1
Private WithEvents objEventoTipo As AdmEvento
Attribute objEventoTipo.VB_VarHelpID = -1
Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1
Private WithEvents objEventoContaProducao As AdmEvento
Attribute objEventoContaProducao.VB_VarHelpID = -1
Private WithEvents objEventoClasFiscIPI As AdmEvento
Attribute objEventoClasFiscIPI.VB_VarHelpID = -1
Private WithEvents objEventoGenero As AdmEvento
Attribute objEventoGenero.VB_VarHelpID = -1
Private WithEvents objEventoISSQN As AdmEvento
Attribute objEventoISSQN.VB_VarHelpID = -1
Private WithEvents objEventoCEST As AdmEvento
Attribute objEventoCEST.VB_VarHelpID = -1

Private Sub AliquotaIPI_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ApropriacaoComp_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ApropriacaoProd_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaContabilLabel_Click()
'Chama o browser de plano de contas

Dim lErro As Long
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim iContaPreenchida As Integer
Dim sConta As String

On Error GoTo Erro_ContaContabilLabel_Click

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaContabil.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then gError 66081

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    Call Chama_Tela("PlanoContaESTLista", colSelecao, objPlanoConta, objEventoContaContabil)

    Exit Sub

Erro_ContaContabilLabel_Click:

    Select Case gErr

        Case 66081
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174903)

    End Select

    Exit Sub
    
End Sub

Private Sub LabelContaProducao_Click()
'Chama o browser de plano de contas

Dim lErro As Long
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim iContaPreenchida As Integer
Dim sConta As String

On Error GoTo Erro_LabelContaProducao_Click

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaProducao.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then gError 66093

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    Call Chama_Tela("PlanoContaESTLista", colSelecao, objPlanoConta, objEventoContaProducao)

    Exit Sub

Erro_LabelContaProducao_Click:

    Select Case gErr

        Case 66093
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174904)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoComprasConfig_Click()
'Traz dados default de Compras Config

Dim lErro As Long

On Error GoTo Erro_BotaoComprasConfig_Click

    lErro = Traz_ComprasConfig_Tela()
    If lErro <> SUCESSO Then Error 25625

    Exit Sub

Erro_BotaoComprasConfig_Click:

    Select Case Err
       
        Case 25625 'Tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174905)

    End Select

    Exit Sub

End Sub

Private Sub ClasFiscIPI_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ClasFiscIPI_GotFocus()
   
   Call MaskEdBox_TrataGotFocus(ClasFiscIPI, iAlterado)
   
End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    If Len(Trim(Codigo.Text)) > 0 Then
        
        lErro = Valor_Positivo_Critica(Codigo.Text)
        If lErro <> SUCESSO Then Error 57771
    
    End If

    Exit Sub
    
Erro_Codigo_Validate:

    Cancel = True
    
    Select Case Err

        Case 57771 'Erro tratado na rotina chamada
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174906)

    End Select

    Exit Sub

End Sub

Private Sub CodigoIPI_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub AliquotaIPI_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AliquotaIPI_Validate

    'Verifica se está preenchida
    If Len(Trim(AliquotaIPI.Text)) = 0 Then Exit Sub

    'Critica se é percentagem
    lErro = Porcentagem_Critica(AliquotaIPI.Text)
    If lErro <> SUCESSO Then Error 54130

    'Testa se percentagem é 100
    If StrParaDbl(AliquotaIPI.Text) = 100# Then Error 54131

    'Coloca na tela o valor formatado
    AliquotaIPI.Text = Format(StrParaDbl(AliquotaIPI.Text), "Fixed")

    Exit Sub

Erro_AliquotaIPI_Validate:

    Cancel = True


    Select Case Err

        Case 54130 'Erro tratado na rotina de chamada
        
        Case 54131
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_ALIQUOTA_INVALIDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174907)

    End Select

    Exit Sub

End Sub

Private Sub Comprado_Click()

    iAlterado = REGISTRO_ALTERADO

    If Comprado.Value = True Then
        ApropriacaoComp.Visible = True
        ApropriacaoComp.ListIndex = -1
        ApropriacaoProd.Visible = False
    End If

End Sub

Private Sub NaoConsideraQuantCotacaoAnterior_Click()

    If NaoConsideraQuantCotacaoAnterior.Value = vbUnchecked Then
        PercentMaisQuantCotacaoAnterior.Enabled = True
        PercentMenosQuantCotacaoAnterior.Enabled = True
    Else
        PercentMaisQuantCotacaoAnterior.Enabled = False
        PercentMenosQuantCotacaoAnterior.Enabled = False
    End If

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConsumoMedioMax_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ControleEstoque_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer
Dim sMascaraConta As String

On Error GoTo Erro_Form_Load

    iFrameAtual = 1

    Set objGridCategoria = New AdmGrid
    Set objEventoClasseUM = New AdmEvento
    Set objEventoTipo = New AdmEvento
    Set objEventoContaContabil = New AdmEvento
    Set objEventoContaProducao = New AdmEvento
    Set objEventoClasFiscIPI = New AdmEvento
    Set objEventoGenero = New AdmEvento
    Set objEventoISSQN = New AdmEvento
    Set objEventoCEST = New AdmEvento

    'Carrega a ComboBox CategoriaProduto com os códigos
    lErro = Carrega_ComboCategoriaProduto()
    If lErro <> SUCESSO Then Error 54077

    'Inicializa as propriedades das Maskaras da Conta
    lErro = MascaraConta(sMascaraConta)
    If lErro <> SUCESSO Then gError 66097
    
    ContaContabil.Mask = sMascaraConta
    ContaProducao.Mask = sMascaraConta

    'Inicializa o Grid
    lErro = Inicializa_Grid_Categoria(objGridCategoria)
    If lErro <> SUCESSO Then Error 54078

    'Por default todas as características devem ficar como marcadas
    For iIndice = 0 To ListaCaracteristicas.ListCount - 1
        ListaCaracteristicas.Selected(iIndice) = True
    Next

    If gcolModulo.Ativo(MODULO_COMPRAS) <> MODULO_ATIVO Then

        'Desabilita o TabCompras
        Frame1(3).Enabled = False
        
    End If
    
    OrigemMercadoria.ListIndex = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 66097, 54077, 54078

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174908)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Inicializa_Grid_Categoria(objGridInt As AdmGrid) As Long

    'Tela em questão
    Set objGridInt.objForm = Me

    'Títulos do Grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Categoria")
    objGridInt.colColuna.Add ("Valor")

    'Campos de edição do Grid
    objGridInt.colCampo.Add (ComboCategoriaProduto.Name)
    objGridInt.colCampo.Add (ComboCategoriaProdutoItem.Name)

    objGridInt.objGrid = GridCategoria

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 21

    'Linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 8

    'Inicializa as colunas do Grid
    iGrid_Categoria_Col = 1
    iGrid_Valor_Col = 2

    GridCategoria.ColWidth(0) = 400

    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Categoria = SUCESSO

    Exit Function

End Function

Private Function Carrega_ComboCategoriaProduto() As Long
'Carrega as Categorias na Combobox

Dim lErro As Long
Dim colCategorias As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto

On Error GoTo Erro_Carrega_ComboCategoriaProduto

    'Lê o código e a descrição de todas as categorias
    lErro = CF("CategoriasProduto_Le_Todas", colCategorias)
    If lErro <> SUCESSO And lErro <> 22542 Then Error 54079

    For Each objCategoriaProduto In colCategorias

        'Insere na combo CategoriaProduto
        ComboCategoriaProduto.AddItem objCategoriaProduto.sCategoria

    Next

    Carrega_ComboCategoriaProduto = SUCESSO

    Exit Function

Erro_Carrega_ComboCategoriaProduto:

    Carrega_ComboCategoriaProduto = Err

    Select Case Err

        Case 54079

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174909)

    End Select

    Exit Function

End Function

Private Sub objEventoContaContabil_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaContabil_evSelecao

    Set objPlanoConta = obj1

    If objPlanoConta.sConta = "" Then
    
        ContaContabil.Text = ""
        
    Else
    
        ContaContabil.PromptInclude = False

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then gError 66082

        ContaContabil.Text = sContaEnxuta
        ContaContabil.PromptInclude = True
        
    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case gErr

        Case 66082
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174910)

    End Select

    Exit Sub

End Sub

Private Sub objEventoContaProducao_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaProducao_evSelecao

    Set objPlanoConta = obj1

    If objPlanoConta.sConta = "" Then
    
        ContaProducao.Text = ""
        
    Else
    
        ContaProducao.PromptInclude = False

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then gError 66094

        ContaProducao.Text = sContaEnxuta
        ContaProducao.PromptInclude = True
        
    End If

    Me.Show

    Exit Sub

Erro_objEventoContaProducao_evSelecao:

    Select Case gErr

        Case 66094
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174911)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()
'Chama TipoDeProduto_Exclui

Dim lErro As Long
Dim objTipoDeProduto As New ClassTipoDeProduto
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    Parent.MousePointer = vbHourglass
    
    'Verifica se o Tipo de Produto foi preenchido
    If Len(Codigo.ClipText) = 0 Then Error 54095

    'Passa para objTipoDeProduto
    objTipoDeProduto.iTipo = StrParaInt(Codigo.ClipText)

    'Lê os dados do Tipo de Produto a ser excluído
    lErro = CF("TipoDeProduto_Le", objTipoDeProduto)
    If lErro <> SUCESSO And lErro <> 22531 Then Error 54093

    'Verifica se Tipo De Produto não está cadastrado
    If lErro <> SUCESSO Then Error 54096

    'Envia aviso perguntando se realmente deseja excluir Tipo De Produto
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_TIPOPRODUTO", objTipoDeProduto.iTipo)

    If vbMsgRes = vbNo Then Exit Sub

    'Exclui o Tipo de Produto
    lErro = CF("TipoDeProduto_Exclui", objTipoDeProduto)
    If lErro <> SUCESSO Then Error 54094

    'Limpa a Tela
    Call Limpa_Tela_TipoProduto
    
    'Fecha comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Parent.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    Parent.MousePointer = vbDefault
    
    Select Case Err

        Case 54093, 54094

        Case 54095
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TIPOPRODUTO_NAO_PREENCHIDO", Err)

        Case 54096
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_INEXISTENTE", Err, objTipoDeProduto.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174912)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_BotaoGravar_Click

    'Chama a função de gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 54119

    'Limpa a tela
    Call Limpa_Tela_TipoProduto

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 54119
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174913)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long
'Chama TipoDeProduto_Grava

Dim lErro As Long
Dim objTipoDeProduto As New ClassTipoDeProduto
Dim objTipoDeProdutoCategoria As New ClassTipoDeProdutoCategoria

On Error GoTo Erro_Gravar_Registro

    Parent.MousePointer = vbHourglass
    
    'Verifica se o Código do Tipo de Produto está preenchido
    If Len(Trim(Codigo.ClipText)) = 0 Then Error 54120

    'Verifica se a Descrição está preenchida
    If Len(Trim(Descricao.Text)) = 0 Then Error 54121

    'Verifica se a Sigla está preenchida
    If Len(Trim(Sigla.Text)) = 0 Then Error 54122

    'Verifica se foi informada a natureza
    If Len(Trim(NaturezaProduto.Text)) = 0 Then Error 54145
    
    'Verifica se o tipo é produzido
    If Produzido.Value = True Then
        objTipoDeProduto.iCompras = PRODUTO_PRODUZIVEL
        'Verifica se foi informada a apropriação
        If ApropriacaoProd.ListIndex = -1 Then Error 54125
    Else
        objTipoDeProduto.iCompras = PRODUTO_COMPRAVEL
        'Verifica se foi informada a apropriação
        If ApropriacaoComp.ListIndex = -1 Then Error 54126
    End If

    'Chama Move_Tela_Memoria para passar os dados da tela para o objeto
    lErro = Move_Tela_Memoria(objTipoDeProduto)
    If lErro <> SUCESSO Then Error 54123

    lErro = Trata_Alteracao(objTipoDeProduto, objTipoDeProduto.iTipo)
    If lErro <> SUCESSO Then Error 32312

    'Chama a função de gravação
    lErro = CF("TipoDeProduto_Grava", objTipoDeProduto)
    If lErro <> SUCESSO Then Error 54124

    Parent.MousePointer = vbDefault
    
    Exit Function

Erro_Gravar_Registro:

    Parent.MousePointer = vbDefault
    
    Gravar_Registro = Err

    Select Case Err

        Case 32312

        Case 54120
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TIPOPRODUTO_NAO_PREENCHIDO", Err)

        Case 54121
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_TIPOPRODUTO_NAO_PREENCHIDA", Err)

        Case 54122
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SIGLA_TIPOPRODUTO_NAO_PREENCHIDA", Err)

        Case 54123, 54124
          
        Case 54125, 54126
            lErro = Rotina_Erro(vbOKOnly, "ERRO_APROPRIACAO_NAO_DEFINIDA", Err)

        Case 54138
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERCENT_MAIS_QUANTCOTACAO_ANTERIOR_NAO_PREENCHIDA", Err)
        
        Case 54139
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERCENT_MENOS_QUANTCOTACAO_ANTERIOR_NAO_PREENCHIDA", Err)
        
        Case 54140
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERCENT_MAIS_RECEB_NAO_PREENCHIDA", Err)
            
        Case 54141
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERCENT_MENOS_RECEB_NAO_PREENCHIDA", Err)
            
        Case 54142
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALIQUOTA_IPI_NAO_PREENCHIDA", Err)
            
        Case 54145
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZA_PRODUTO_NAO_PREENCHIDA", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174914)

    End Select

    Exit Function

End Function

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub ComboCategoriaProduto_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComboCategoriaProdutoItem_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ControleEstoque_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescricaoClasseUM_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub IntRessup_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LblClasseUM_Click()
'Abre browse de classes de Unidades de Medida

Dim objClasseUM As New ClassClasseUM
Dim colSelecao As New Collection

    'Se a classe estiver preenchida
    If Len(Trim(ClasseUM.Text)) > 0 Then
        'Usa a funcao codigo extrai
        objClasseUM.iClasse = Codigo_Extrai(ClasseUM.Text)
    End If

    'Chama tela
    Call Chama_Tela("ClasseUMLista", colSelecao, objClasseUM, objEventoClasseUM)

    Exit Sub

End Sub

Private Sub ListaCaracteristicas_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MesesConsumoMedio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MesesConsumoMedio_Validate(Cancel As Boolean)
Dim lErro As Long

On Error GoTo Erro_MesesConsumoMedio_Validate

    'Verifica se está preenchida
    If Len(Trim(MesesConsumoMedio.Text)) = 0 Then Exit Sub

    'Critica se o valor é positivo
    lErro = Valor_Positivo_Critica(MesesConsumoMedio.Text)
    If lErro <> SUCESSO Then Error 54137

    Exit Sub

Erro_MesesConsumoMedio_Validate:

    Cancel = True


    Select Case Err

        Case 54137 'Erro criticado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174915)

    End Select

    Exit Sub

End Sub


Private Sub NaoTemFaixaReceb_Click()

    If NaoTemFaixaReceb.Value = vbUnchecked Then
        
        PercentMaisReceb.Enabled = True
        PercentMenosReceb.Enabled = True
        RecebForaFaixa(0).Enabled = True
        RecebForaFaixa(1).Enabled = True
    Else
        PercentMaisReceb.Enabled = False
        PercentMenosReceb.Enabled = False
        RecebForaFaixa(0).Enabled = False
        RecebForaFaixa(1).Enabled = False

    End If

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NaturezaProduto_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub NaturezaProduto_Click()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub NomeUMCompra_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NomeUMEstoque_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NomeUMVenda_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoClasseUM_evSelecao(obj1 As Object)

Dim objClasseUM As New ClassClasseUM

    Set objClasseUM = obj1

    'Preenche Text da ClasseUM
    ClasseUM.Text = CStr(objClasseUM.iClasse)
    Call ClasseUM_Validate(bSGECancelDummy)

    Me.Show

End Sub

Private Sub LblTipo_Click()
'Abrir browse de Tipos de Produto

Dim lErro As Long
Dim objTipoDeProduto As New ClassTipoDeProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LblTipo_Click

    'Usa Move_Tela_Memoria
    lErro = Move_Tela_Memoria(objTipoDeProduto)
    If lErro <> SUCESSO Then Error 54086

    'Chama tela
    Call Chama_Tela("TipoProdutoLista", colSelecao, objTipoDeProduto, objEventoTipo)

    Exit Sub

Erro_LblTipo_Click:

    Select Case Err

        Case 54086

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174916)

    End Select

    Exit Sub

End Sub

Private Sub objEventoTipo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTipoDeProduto As New ClassTipoDeProduto

On Error GoTo Erro_objEventoTipo_evSelecao

    Set objTipoDeProduto = obj1

    'Chama Traz_TipoDeProduto_Tela
    lErro = Traz_TipoDeProduto_Tela(objTipoDeProduto)
    If lErro <> SUCESSO And lErro <> 54082 Then Error 54089

    'se o tipo não estiver cadastrado ==> erro
    If lErro = 54082 Then Error 55863

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoTipo_evSelecao:

    Select Case Err

        Case 54089

        Case 55863
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_INEXISTENTE", Err, objTipoDeProduto.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174917)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objTipoDeProduto As ClassTipoDeProduto) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se algum Tipo de Produto foi passado por parametro
    If Not (objTipoDeProduto Is Nothing) Then

        'Traz os dados do Tipo para a tela
        lErro = Traz_TipoDeProduto_Tela(objTipoDeProduto)
        If lErro <> SUCESSO And lErro <> 54082 Then Error 54087

        'Se for um Tipo novo
        If lErro = 54082 Then
            
            'Coloca o código na tela
            Codigo.Text = objTipoDeProduto.iTipo
            
        End If
    
    End If

    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 54087

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174918)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Traz_TipoDeProduto_Tela(objTipoDeProduto As ClassTipoDeProduto) As Long
'Mostra os dados do Tipo de Produto na tela

Dim lErro As Long
Dim iIndice As Integer
Dim objTipoDeProdutoCategoria As New ClassTipoDeProdutoCategoria

On Error GoTo Erro_Traz_TipoDeProduto_Tela

    'Lê a tabela TiposDeProduto a partir do Tipo de Produto
    lErro = CF("TipoDeProduto_Le", objTipoDeProduto)
    If lErro <> SUCESSO And lErro <> 22531 Then Error 54085
    
    If lErro = 22531 Then Error 54082

    'Limpa a Tela
    Call Limpa_Tela_TipoProduto

    'Lê a tabela TipoDeProdutoCategorias a partir da Tipo De Produto
    lErro = CF("TipoDeProduto_Le_Categorias", objTipoDeProduto, objTipoDeProduto.colCategoriaItem)
    If lErro <> SUCESSO Then Error 54083

    'Exibe os dados de objTipoDeProduto na tela
    Codigo.Text = objTipoDeProduto.iTipo
    Descricao.Text = objTipoDeProduto.sDescricao
    Sigla.Text = objTipoDeProduto.sSigla

    'Move o que está no objeto para a Lista de Características
    If objTipoDeProduto.iFaturamento = 0 Then ListaCaracteristicas.Selected(0) = False
    If objTipoDeProduto.iPCP = 0 Then ListaCaracteristicas.Selected(1) = False
    If objTipoDeProduto.iKitBasico = 0 Then ListaCaracteristicas.Selected(2) = False
    If objTipoDeProduto.iKitInt = 0 Then ListaCaracteristicas.Selected(3) = False
    If objTipoDeProduto.iKitVendaComp = 0 Then ListaCaracteristicas.Selected(4) = False

    'Exibe os dados da coleção na tela
    For iIndice = 1 To objTipoDeProduto.colCategoriaItem.Count
        'Insere no Grid Categoria
        GridCategoria.TextMatrix(iIndice, iGrid_Categoria_Col) = objTipoDeProduto.colCategoriaItem.Item(iIndice).sCategoria
        GridCategoria.TextMatrix(iIndice, iGrid_Valor_Col) = objTipoDeProduto.colCategoriaItem.Item(iIndice).sItem
    Next

    objGridCategoria.iLinhasExistentes = objTipoDeProduto.colCategoriaItem.Count

    'Preenche a Combo ControleEstoque
    For iIndice = 0 To ControleEstoque.ListCount - 1
        If ControleEstoque.ItemData(iIndice) = objTipoDeProduto.iControleEstoque Then
            ControleEstoque.ListIndex = iIndice
            Exit For
        End If
    Next

    If objTipoDeProduto.iApropriacaoCusto = APROPR_CUSTO_MEDIO Then
        Comprado.Value = True
        'Preenche a Combo Apropriação
        For iIndice = 0 To ApropriacaoComp.ListCount
            If ApropriacaoComp.ItemData(iIndice) = objTipoDeProduto.iApropriacaoCusto Then
                ApropriacaoComp.ListIndex = iIndice
                Exit For
            End If
        Next
    Else
        
        Produzido.Value = True
        'Preenche a Combo Apropriação
        For iIndice = 0 To ApropriacaoProd.ListCount - 1

            If ApropriacaoProd.ItemData(iIndice) = objTipoDeProduto.iApropriacaoCusto Then
                ApropriacaoProd.ListIndex = iIndice
                Exit For
            End If
        Next

    End If

    'Exibe a Natureza
    For iIndice = 0 To NaturezaProduto.ListCount - 1
        If NaturezaProduto.ItemData(iIndice) = objTipoDeProduto.iNatureza Then
            NaturezaProduto.ListIndex = iIndice
            Exit For
        End If
    Next
    
    Call Traz_TabUnidadesMedida_Tela(objTipoDeProduto)

    'Traz os dados do frame Estoque/Compras para tela
    MesesConsumoMedio.Text = CStr(objTipoDeProduto.iMesesConsumoMedio)
    IntRessup.Text = CStr(objTipoDeProduto.iIntRessup)
    TempoRessupMax.Text = (objTipoDeProduto.dTempoRessupMax) * 100
    ConsumoMedioMax.Text = (objTipoDeProduto.dConsumoMedioMax) * 100
    If objTipoDeProduto.dResiduo <> -1 Then
        Residuo.Text = (objTipoDeProduto.dResiduo) * 100
    Else
        Residuo.Text = ""
    End If
    PercentMaisQuantCotacaoAnterior.Text = Format(objTipoDeProduto.dPercentMaisQuantCotAnt * 100, "Fixed")
    PercentMenosQuantCotacaoAnterior.Text = Format(objTipoDeProduto.dPercentMenosQuantCotAnt * 100, "Fixed")
    PercentMaisReceb.Text = Format(objTipoDeProduto.dPercentMaisReceb * 100, "Fixed")
    PercentMenosReceb.Text = Format(objTipoDeProduto.dPercentMenosReceb * 100, "Fixed")
    
    'NaoConsideraQuantCotacaoAnterior
    If objTipoDeProduto.iConsideraQuantCotAnt = 0 Then
        NaoConsideraQuantCotacaoAnterior.Value = vbUnchecked
    Else
        NaoConsideraQuantCotacaoAnterior.Value = vbChecked
    End If
    
    'Se NaoTemFaixaReceb estiver marcado
    If objTipoDeProduto.iTemFaixaReceb = 0 Then
        NaoTemFaixaReceb.Value = vbUnchecked
    Else
        NaoTemFaixaReceb.Value = vbChecked
    End If
    
    'Traz o valor de ObjTipoDeProduto.iRecebForaFaixa para tela
    RecebForaFaixa(objTipoDeProduto.iRecebForaFaixa) = True
    
    lErro = Traz_TabTributacao_Tela(objTipoDeProduto)
    If lErro <> SUCESSO Then Error 54084

    iAlterado = 0

    Traz_TipoDeProduto_Tela = SUCESSO

    Exit Function

Erro_Traz_TipoDeProduto_Tela:

    Traz_TipoDeProduto_Tela = Err

    Select Case Err

        Case 54082, 54083, 54084, 54085

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174919)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

    Set objGridCategoria = Nothing
    Set objEventoClasseUM = Nothing
    Set objEventoTipo = Nothing
    Set objEventoContaContabil = Nothing
    Set objEventoContaProducao = Nothing
    Set objEventoClasFiscIPI = Nothing
    Set objEventoGenero = Nothing
    Set objEventoISSQN = Nothing
    Set objEventoCEST = Nothing

End Sub

Private Sub OrigemMercadoria_Click()

    iAlterado = REGISTRO_ALTERADO
    OrigemMercadoria.ToolTipText = OrigemMercadoria.Text

End Sub

Private Sub PercentMaisQuantCotacaoAnterior_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PercentMaisQuantCotacaoAnterior_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercentMaisQuantCotacaoAnterior_Validate

    'Verifica se esta preenchida
    If Len(Trim(PercentMaisQuantCotacaoAnterior.Text)) = 0 Then Exit Sub

    'Critica se é percentagem
    lErro = Porcentagem_Critica(PercentMaisQuantCotacaoAnterior.Text)
    If lErro <> SUCESSO Then Error 54106
    
    'Testa se percentagem é 100
    If StrParaDbl(PercentMaisQuantCotacaoAnterior.Text) = 100# Then Error 54107

    'Coloca na tela
    PercentMaisQuantCotacaoAnterior.Text = Format(PercentMaisQuantCotacaoAnterior.Text, "Fixed")

    Exit Sub

Erro_PercentMaisQuantCotacaoAnterior_Validate:

    Cancel = True


    Select Case Err

        Case 54106 'Erro criticado na rotina de chamada
                
        Case 54107
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERCENTUAL_IGUAL_100", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174920)

    End Select

    Exit Sub

End Sub

Private Sub PercentMaisReceb_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PercentMaisReceb_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercentMaisReceb_Validate

    'Verifica se está preenchida
    If Len(Trim(PercentMaisReceb.Text)) = 0 Then Exit Sub

    'Critica se é percentagem
    lErro = Porcentagem_Critica(PercentMaisReceb.Text)
    If lErro <> SUCESSO Then Error 54108

    'Testa se percentagem é 100
    If StrParaDbl(PercentMaisReceb.Text) = 100# Then Error 54109

    'Coloca na tela
    PercentMaisReceb.Text = Format(PercentMaisReceb.Text, "Fixed")

    Exit Sub

Erro_PercentMaisReceb_Validate:

    Cancel = True


    Select Case Err

        Case 54108  'Erro criticado na rotina chamada
            
        Case 54109
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERCENTUAL_IGUAL_100", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174921)

    End Select

    Exit Sub

End Sub

Private Sub PercentMenosQuantCotacaoAnterior_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PercentMenosQuantCotacaoAnterior_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercentMenosQuantCotacaoAnterior_Validate

    'Verifica se está preenchida
    If Len(Trim(PercentMenosQuantCotacaoAnterior.Text)) = 0 Then Exit Sub

    'Critica se é percentagem
    lErro = Porcentagem_Critica(PercentMenosQuantCotacaoAnterior.Text)
    If lErro <> SUCESSO Then Error 54102
      
    'Testa se percentagem é 100
    If StrParaDbl(PercentMenosQuantCotacaoAnterior.Text) = 100# Then Error 54103

    'Coloca na tela
    PercentMenosQuantCotacaoAnterior.Text = Format(PercentMenosQuantCotacaoAnterior.Text, "Fixed")

    Exit Sub

Erro_PercentMenosQuantCotacaoAnterior_Validate:

    Cancel = True


    Select Case Err

        Case 54102 'Erro criticado na rotina de chamada
            
        Case 54103
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERCENTUAL_IGUAL_100", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174922)

    End Select

    Exit Sub

End Sub

Private Sub PercentMenosReceb_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PercentMenosReceb_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercentMenosReceb_Validate

    'Verifica se está preenchido
    If Len(Trim(PercentMenosReceb.Text)) = 0 Then Exit Sub

    'Critica se é percentagem
    lErro = Porcentagem_Critica(PercentMenosReceb.Text)
    If lErro <> SUCESSO Then Error 54110

    'Testa se percentagem é 100
    If StrParaDbl(PercentMenosReceb.Text) = 100# Then Error 54111

    'Coloca na tela
    PercentMenosReceb.Text = Format(PercentMenosReceb.Text, "Fixed")

    Exit Sub

Erro_PercentMenosReceb_Validate:

    Cancel = True


    Select Case Err

        Case 54110 'Erro criticado na rotina de chamada
            
        Case 54111
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERCENTUAL_IGUAL_100", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174923)

    End Select

    Exit Sub

End Sub

Private Sub Produzido_Click()

    iAlterado = REGISTRO_ALTERADO

    If Produzido.Value = True Then
        ApropriacaoProd.ListIndex = -1
        ApropriacaoProd.Visible = True
        ApropriacaoComp.Visible = False
    End If

End Sub

Private Sub RecebForaFaixa_Click(Index As Integer)

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Residuo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Sigla_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TabStrip1_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        If gcolModulo.Ativo(MODULO_COMPRAS) <> MODULO_ATIVO Then

            If TabStrip1.SelectedItem.Index = 3 Then

                Frame1(3).Enabled = False
                Frame1(iFrameAtual).Visible = True
                TabStrip1.Tabs.Item(iFrameAtual).Selected = True
                Exit Sub
            End If
        End If

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index

    End If

End Sub

Private Sub Trata_ComboCategoriaProdutoItem()

Dim lErro As Long
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim iIndice As Integer

On Error GoTo Erro_Trata_ComboCategoriaProdutoItem
        
    ComboCategoriaProdutoItem.Clear

    'Se alguém estiver selecionado
    If Len(GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Categoria_Col)) > 0 Then

        'Preencher a Combo de Itens desta Categoria
        objCategoriaProduto.sCategoria = GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Categoria_Col)

        lErro = Carrega_ComboCategoriaProdutoItem(objCategoriaProduto)
        If lErro <> SUCESSO Then Error 54104

        For iIndice = 0 To ComboCategoriaProdutoItem.ListCount - 1
            If ComboCategoriaProdutoItem.List(iIndice) = ComboCategoriaProdutoItem.Text Then
                ComboCategoriaProdutoItem.ListIndex = iIndice
                Exit For
            End If
        Next

    End If
    
    Exit Sub

Erro_Trata_ComboCategoriaProdutoItem:

    Select Case Err

        Case 54104

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174924)

    End Select

End Sub

Private Function Carrega_ComboCategoriaProdutoItem(objCategoriaProduto As ClassCategoriaProduto) As Long
'Carrega a Combobox de Itens(valores possiveis) da Categoria

Dim lErro As Long
Dim colItensCategoria As New Collection
Dim objCategoriaProdutoItem As ClassCategoriaProdutoItem

On Error GoTo Erro_Carrega_ComboCategoriaProdutoItem

    'Lê a tabela CategoriaProdutoItem a partir da Categoria
    lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colItensCategoria)
    If lErro <> SUCESSO Then Error 54105

    'Insere na combo CategoriaProdutoItem
    For Each objCategoriaProdutoItem In colItensCategoria

        'Insere na combo CategoriaProduto
        ComboCategoriaProdutoItem.AddItem objCategoriaProdutoItem.sItem

    Next

    Carrega_ComboCategoriaProdutoItem = SUCESSO

    Exit Function

Erro_Carrega_ComboCategoriaProdutoItem:

    Carrega_ComboCategoriaProdutoItem = Err

    Select Case Err

        Case 54105

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174925)

    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        Select Case GridCategoria.Col

            Case iGrid_Categoria_Col

                lErro = Saida_Celula_Categoria(objGridInt)
                If lErro <> SUCESSO Then Error 54134

            Case iGrid_Valor_Col

                lErro = Saida_Celula_Valor(objGridInt)
                If lErro <> SUCESSO Then Error 54133

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 54132

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 54132
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 54133, 54134

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174926)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Categoria(objGridInt As AdmGrid) As Long
'Faz a critica da celula conta do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objCategoriaProduto As New ClassCategoriaProduto

On Error GoTo Erro_Saida_Celula_Categoria

    Set objGridInt.objControle = ComboCategoriaProduto

    'Se a ComboCategoria estiver preenchida
    If Len(Trim(ComboCategoriaProduto.Text)) > 0 Then

        'Tenta selecionar na combo
        lErro = Combo_Item_Igual(ComboCategoriaProduto)
        If lErro <> SUCESSO Then

            'Preenche o objeto com a Categoria
             objCategoriaProduto.sCategoria = ComboCategoriaProduto.Text

             'Lê Categoria De Produto no BD
             lErro = CF("CategoriaProduto_Le", objCategoriaProduto)
             If lErro <> SUCESSO And lErro <> 22540 Then Error 54115
             If lErro <> SUCESSO Then Error 54114  'Categoria não está cadastrada

        End If

        If ComboCategoriaProduto.Text <> GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Categoria_Col) Then
            GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Valor_Col) = ""
        End If
        'Verifica se já existe a categoria no Grid
        For iIndice = 1 To objGridCategoria.iLinhasExistentes
            If iIndice <> GridCategoria.Row Then If GridCategoria.TextMatrix(iIndice, iGrid_Categoria_Col) = ComboCategoriaProduto.Text Then Error 54113
        Next

        If GridCategoria.Row - GridCategoria.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    Else
        GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Valor_Col) = ""
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 54112

    Saida_Celula_Categoria = SUCESSO

    Exit Function

Erro_Saida_Celula_Categoria:

    Saida_Celula_Categoria = Err

    Select Case Err

        Case 54112
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 54113
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_JA_SELECIONADA", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 54114  'Categoria não está cadastrada

            'Se não for perguntar se deseja criar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CATEGORIAPRODUTO", ComboCategoriaProduto.Text)

            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                'Chama a Tela "CategoriaProduto"
                Call Chama_Tela("CategoriaProduto", objCategoriaProduto)

            Else

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 54115
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174927)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Valor(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Item do grid que está deixando de ser a corrente

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As New ClassCategoriaProdutoItem
Dim colItens As New Collection

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridCategoria.objControle = ComboCategoriaProdutoItem

    'Se a ComboCategoriaProdutoItem estiver preenchida
    If Len(Trim(ComboCategoriaProdutoItem.Text)) > 0 Then

        'se o campo de categoria estiver vazio ==> erro
        If Len(GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Categoria_Col)) = 0 Then Error 55905

        'Tenta selecionar na combo
        lErro = Combo_Item_Igual(ComboCategoriaProdutoItem)
        If lErro <> SUCESSO Then

            'Preenche o objeto com a Categoria
            objCategoriaProdutoItem.sCategoria = GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Categoria_Col)
            objCategoriaProdutoItem.sItem = ComboCategoriaProdutoItem.Text

            'Lê Categoria De Produto no BD
            lErro = CF("CategoriaProduto_Le_Item", objCategoriaProdutoItem)
            If lErro <> SUCESSO And lErro <> 22603 Then Error 54118
            If lErro <> SUCESSO Then Error 54117  'Item da Categoria não está cadastrado

        End If

        If GridCategoria.Row - GridCategoria.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 54116

    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = Err

    Select Case Err

        Case 54116
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 54117  'Item da Categoria não está cadastrado

            'Se não for perguntar se deseja criar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CATEGORIAPRODUTOITEM", ComboCategoriaProdutoItem.Text)

            If vbMsgRes = vbYes Then

                'Preenche o objeto com a Categoria
                objCategoriaProduto.sCategoria = ComboCategoriaProduto.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                'Chama a Tela "CategoriaProduto"
                Call Chama_Tela("CategoriaProduto", objCategoriaProduto)

            Else

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 54118
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 55905
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_CATEGORIA_NAO_PREENCHIDA", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174928)

    End Select

    Exit Function

End Function

Private Sub GridCategoria_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridCategoria, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCategoria, iAlterado)
    End If

End Sub

Private Sub GridCategoria_GotFocus()

    Call Grid_Recebe_Foco(objGridCategoria)

End Sub

Private Sub GridCategoria_EnterCell()

    Call Grid_Entrada_Celula(objGridCategoria, iAlterado)

End Sub

Private Sub GridCategoria_LeaveCell()

    Call Saida_Celula(objGridCategoria)

End Sub

Private Sub GridCategoria_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridCategoria)

End Sub

Private Sub GridCategoria_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCategoria, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCategoria, iAlterado)
    End If

End Sub

Private Sub GridCategoria_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridCategoria)

End Sub

Private Sub GridCategoria_RowColChange()

    Call Grid_RowColChange(objGridCategoria)

End Sub

Private Sub GridCategoria_Scroll()

    Call Grid_Scroll(objGridCategoria)

End Sub

Private Sub ComboCategoriaProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComboCategoriaProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCategoria)

End Sub

Private Sub ComboCategoriaProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCategoria)

End Sub

Private Sub ComboCategoriaProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCategoria.objControle = ComboCategoriaProduto
    lErro = Grid_Campo_Libera_Foco(objGridCategoria)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub ComboCategoriaProdutoItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComboCategoriaProdutoItem_GotFocus()

Dim lErro As Long
Dim sCategoriaProduto As String
    
    sCategoriaProduto = ComboCategoriaProdutoItem.Text
    
    'Preenche com os ítens relacionados a Categoria correspondente
    Call Trata_ComboCategoriaProdutoItem

    Call Grid_Campo_Recebe_Foco(objGridCategoria)

    'Se havia algo preenchido antes de limpar a combo, preenche novamente com o mesmo valor
    If Len(Trim(sCategoriaProduto)) > 0 Then
        ComboCategoriaProdutoItem.Text = sCategoriaProduto
    End If

    Exit Sub

Erro_ComboCategoriaProdutoItem_GotFocus:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174929)

    End Select

    Exit Sub

End Sub

Private Sub ComboCategoriaProdutoItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCategoria)

End Sub

Private Sub ComboCategoriaProdutoItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCategoria.objControle = ComboCategoriaProdutoItem
    lErro = Grid_Campo_Libera_Foco(objGridCategoria)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA
'""""""""""""""""""""""""""""""""""""""""""""""

'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objTipoDeProduto As New ClassTipoDeProduto

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TiposDeProduto"

    If Len(Trim(Codigo.Text)) > 0 Then objTipoDeProduto.iTipo = Codigo.Text
    If Len(Trim(Descricao.Text)) > 0 Then objTipoDeProduto.sDescricao = Descricao.Text
    If Len(Trim(Sigla.Text)) > 0 Then objTipoDeProduto.sSigla = Sigla.Text

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "TipoDeProduto", objTipoDeProduto.iTipo, 0, "TipoDeProduto"
    colCampoValor.Add "Descricao", objTipoDeProduto.sDescricao, STRING_TIPODEPRODUTO_DESCRICAO, "Descricao"
    colCampoValor.Add "Sigla", objTipoDeProduto.sSigla, STRING_TIPODEPRODUTO_SIGLA, "Sigla"
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174930)

    End Select

    Exit Sub

End Sub

Function Move_Tela_Memoria(objTipoDeProduto As ClassTipoDeProduto) As Long
'Move os dados da tela para memória

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim objTipoDeProdutoCategoria As New ClassTipoDeProdutoCategoria

On Error GoTo Erro_Move_Tela_Memoria

    'Move os dados da tela para objTipoDeProduto
    If Len(Trim(Codigo.Text)) > 0 Then objTipoDeProduto.iTipo = Codigo.Text
    If Len(Trim(Descricao.Text)) > 0 Then objTipoDeProduto.sDescricao = Descricao.Text
    If Len(Trim(Sigla.Text)) > 0 Then objTipoDeProduto.sSigla = Sigla.Text

    'Move o que esta selecionado na Combobox ControleEstoque para o objeto
    If ControleEstoque.ListIndex <> -1 Then objTipoDeProduto.iControleEstoque = ControleEstoque.ItemData(ControleEstoque.ListIndex)

    'Move o que está selecionado em natureza
    If Len(Trim(NaturezaProduto.List(NaturezaProduto.ListIndex))) > 0 Then objTipoDeProduto.iNatureza = NaturezaProduto.ItemData(NaturezaProduto.ListIndex)
    
    'Move o que está em Lista de Características para o objeto
    If ListaCaracteristicas.Selected(0) = True Then objTipoDeProduto.iFaturamento = 1
    If ListaCaracteristicas.Selected(1) = True Then objTipoDeProduto.iPCP = 1
    If ListaCaracteristicas.Selected(2) = True Then objTipoDeProduto.iKitBasico = 1
    If ListaCaracteristicas.Selected(3) = True Then objTipoDeProduto.iKitInt = 1
    If ListaCaracteristicas.Selected(4) = True Then objTipoDeProduto.iKitVendaComp = 1

    If Produzido.Value = True Then
        'Move o que esta selecionado na Combobox Apropriação para o objeto
        If ApropriacaoProd.ListIndex <> -1 Then objTipoDeProduto.iApropriacaoCusto = ApropriacaoProd.ItemData(ApropriacaoProd.ListIndex)
    Else
        If ApropriacaoComp.ListIndex <> -1 Then objTipoDeProduto.iApropriacaoCusto = ApropriacaoComp.ItemData(ApropriacaoComp.ListIndex)
    End If

    'Não está na tela
    objTipoDeProduto.iPrazoValidade = 0

    'Pega o que está no frame Unidades de Medidas
    If Len(Trim(ClasseUM.Text)) > 0 Then
        objTipoDeProduto.iClasseUM = CInt(ClasseUM.Text)
    Else
        objTipoDeProduto.iClasseUM = 0
    End If

    objTipoDeProduto.sSiglaUMCompra = SiglaUMCompra.Text
    objTipoDeProduto.sSiglaUMEstoque = SiglaUMEstoque.Text
    objTipoDeProduto.sSiglaUMVenda = SiglaUMVenda.Text
    objTipoDeProduto.sSiglaUMTrib = SiglaUMTrib.Text

    'Ir preenchendo uma coleção com todas as linhas "existentes" do grid
    For iIndice = 1 To objGridCategoria.iLinhasExistentes

        'Verifica se a Categoria foi preenchida
        If Len(Trim(GridCategoria.TextMatrix(iIndice, iGrid_Categoria_Col))) <> 0 Then

            Set objTipoDeProdutoCategoria = New ClassTipoDeProdutoCategoria

            If Len(Codigo.Text) > 0 Then objTipoDeProdutoCategoria.iTipoDeProduto = CInt(Codigo.Text)
            objTipoDeProdutoCategoria.sCategoria = GridCategoria.TextMatrix(iIndice, iGrid_Categoria_Col)
            objTipoDeProdutoCategoria.sItem = GridCategoria.TextMatrix(iIndice, iGrid_Valor_Col)

            objTipoDeProduto.colCategoriaItem.Add objTipoDeProdutoCategoria

        End If

    Next

    'Passa os dados do frame Estoque/Compras para o objeto
    objTipoDeProduto.iMesesConsumoMedio = StrParaInt(MesesConsumoMedio.Text)
    objTipoDeProduto.dConsumoMedioMax = PercentParaDbl(ConsumoMedioMax.Text & "%")
    objTipoDeProduto.dTempoRessupMax = PercentParaDbl(TempoRessupMax.Text & "%")
    objTipoDeProduto.iIntRessup = StrParaInt(IntRessup.Text)
    objTipoDeProduto.dPercentMaisQuantCotAnt = PercentParaDbl(PercentMaisQuantCotacaoAnterior.Text & "%")
    objTipoDeProduto.dPercentMenosQuantCotAnt = PercentParaDbl(PercentMenosQuantCotacaoAnterior.Text & "%")
    objTipoDeProduto.dPercentMaisReceb = PercentParaDbl(PercentMaisReceb.Text & "%")
    objTipoDeProduto.dPercentMenosReceb = PercentParaDbl(PercentMenosReceb.Text & "%")
    If Len(Trim(Residuo.ClipText)) > 0 Then
        objTipoDeProduto.dResiduo = PercentParaDbl(Residuo.Text & "%")
    Else
        objTipoDeProduto.dResiduo = -1
    End If
    
    If NaoConsideraQuantCotacaoAnterior.Value = vbChecked Then
        objTipoDeProduto.iConsideraQuantCotAnt = 1
    Else
        objTipoDeProduto.iConsideraQuantCotAnt = 0
    End If
    
    If NaoTemFaixaReceb.Value = vbChecked Then
        objTipoDeProduto.iTemFaixaReceb = 1
    Else
        objTipoDeProduto.iTemFaixaReceb = 0
    End If
        
    If RecebForaFaixa(0).Value = True Then
        objTipoDeProduto.iRecebForaFaixa = 0
    Else
        objTipoDeProduto.iRecebForaFaixa = 1
    End If
    
    lErro = Move_TabTributacao_Memoria(objTipoDeProduto)
    If lErro <> SUCESSO Then Error 54092

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case 54092

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174931)

    End Select

    Exit Function

End Function

'Preenche os campos da tela com os correspondentes do BD
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objTipoDeProduto As New ClassTipoDeProduto

On Error GoTo Erro_Tela_Preenche

    objTipoDeProduto.iTipo = colCampoValor.Item("TipoDeProduto").vValor
    
    'Traz dados da TipoDeProduto para a Tela
    lErro = Traz_TipoDeProduto_Tela(objTipoDeProduto)
    If lErro <> SUCESSO Then Error 54081

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 54081

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174932)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 54088

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case Err

        Case 54088

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174933)

    End Select

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 54136

    'Limpa a Tela
    Call Limpa_Tela_TipoProduto
    
    'Fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 54136

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174934)

    End Select

End Sub

Private Sub Limpa_Tela_TipoProduto()

Dim lErro As Long
Dim iIndice As Integer

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Limpa TextBox e MaskedEditBox
    Call Limpa_Tela(Me)

    Sigla.Text = ""

    'Limpa os textos das Combos
    ComboCategoriaProduto.Text = ""
    ComboCategoriaProdutoItem.Text = ""
    ControleEstoque.ListIndex = -1
    NaturezaProduto.ListIndex = -1
    SiglaUMEstoque.Clear
    SiglaUMCompra.Clear
    SiglaUMVenda.Clear
    ApropriacaoProd.ListIndex = -1
    ApropriacaoComp.ListIndex = -1

    'Limpa os textos dos Labels
    DescricaoClasseUM.Caption = ""
    NomeUMEstoque.Caption = ""
    NomeUMCompra.Caption = ""
    NomeUMVenda.Caption = ""
    
    DescISSQN.Caption = ""
    DescGenero.Caption = ""
    CodServNFe.Caption = ""
    NomeUMTrib.Caption = ""
    SiglaUMTrib.Clear
    ProdutoEspecifico.ListIndex = -1

    'Limpa GridCategoria
    Call Grid_Limpa(objGridCategoria)

    'Limpa o frame Estoque/Compras
    MesesConsumoMedio.Text = ""
    ConsumoMedioMax.Text = ""
    TempoRessupMax.Text = ""
    IntRessup.Text = ""
    PercentMaisQuantCotacaoAnterior.Text = ""
    PercentMenosQuantCotacaoAnterior.Text = ""
    PercentMaisReceb.Text = ""
    PercentMenosReceb.Text = ""
    Residuo.Text = ""

    'Por default todas as características devem ficar como marcadas
    For iIndice = 0 To ListaCaracteristicas.ListCount - 1
        ListaCaracteristicas.Selected(iIndice) = True
    Next

    OrigemMercadoria.ListIndex = 0
    
    iClasseUMAnterior = 0
    iClasseUMAlterada = 0
        
    Exit Sub

End Sub

Private Sub IncideIPI_Click()

    'Verifica se está selecionado
    If IncideIPI.Value = vbUnchecked Then
        'Senão estiver desativa e limpa os campos ligados ao IPI
        AliquotaIPI.Text = ""
        AliquotaIPI.Enabled = False
        CodigoIPI.Text = ""
        CodigoIPI.Enabled = False
    'Se estiver selecionado
    Else
        'Ativa os campos relacionados ao IPI
        AliquotaIPI.Enabled = True
        CodigoIPI.Enabled = True
    End If

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ClasseUM_Change()

    iAlterado = REGISTRO_ALTERADO
    iClasseUMAlterada = REGISTRO_ALTERADO

End Sub

Private Sub ClasseUM_GotFocus()

Dim iClasseUMAux As Integer

    iClasseUMAux = iClasseUMAlterada
    Call MaskEdBox_TrataGotFocus(ClasseUM, iAlterado)
    iClasseUMAlterada = iClasseUMAux

End Sub

Private Sub ClasseUM_Validate(Cancel As Boolean)

Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long
Dim objClasseUM As New ClassClasseUM
Dim colSiglas As New Collection

On Error GoTo Erro_ClasseUM_Validate

    'Verifica se a ClasseUM foi alterada
    If iClasseUMAlterada <> REGISTRO_ALTERADO Then Exit Sub

    'Se estiver Preenchida
    If Len(Trim(ClasseUM.Text)) > 0 Then

        'Critica o valor
        lErro = Inteiro_Critica(ClasseUM.Text)
        If lErro <> SUCESSO Then Error 54090

        objClasseUM.iClasse = StrParaInt(ClasseUM.Text)

    End If

    'Limpa o conteúdo das Combos e a Descrição da Classe
    SiglaUMEstoque.Clear
    SiglaUMCompra.Clear
    SiglaUMVenda.Clear
    SiglaUMTrib.Clear
    NomeUMEstoque.Caption = ""
    NomeUMVenda.Caption = ""
    NomeUMCompra.Caption = ""
    NomeUMTrib.Caption = ""
    DescricaoClasseUM.Caption = ""

    If objClasseUM.iClasse = 0 Then
        iClasseUMAlterada = 0
        iClasseUMAnterior = 0
        Exit Sub
    End If
    
    'Verificar se é uma classe cadastrada em ClasseUM
    lErro = CF("ClasseUM_Le", objClasseUM)
    If lErro <> SUCESSO And lErro <> 22537 Then Error 54091

    If lErro = 22537 Then Error 54135

    'Coloca na Tela a Descrição
    DescricaoClasseUM.Caption = objClasseUM.sDescricao

    lErro = Carrega_CombosUM(objClasseUM)
    If lErro <> SUCESSO Then Error 55160

    iClasseUMAlterada = 0
    iClasseUMAnterior = objClasseUM.iClasse
    
    Exit Sub

Exit Sub

Erro_ClasseUM_Validate:

    Cancel = True


    Select Case Err

        Case 54090, 54091, 55160

        Case 54135
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLASSEUM", objClasseUM.iClasse)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("ClasseUM", objClasseUM)
            Else
            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174935)

    End Select

    Exit Sub

End Sub

Private Sub SiglaUMCompra_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Function Carrega_CombosUM(objClasseUM As ClassClasseUM) As Long
'Carrega as combos de Unidades de Medida de acordo com a ClasseUM passada
'selecionando a UM de estoque
Dim lErro As Long
Dim colSiglas As New Collection
Dim iIndice As Integer
Dim iIndice2 As Integer

On Error GoTo Erro_Carrega_CombosUM

    'Lê as U.M. da Classe passada
    lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
    If lErro <> SUCESSO And lErro <> 22539 Then Error 54128

    'Carrega as combos
    If lErro = SUCESSO Then

        For iIndice = 1 To colSiglas.Count
            SiglaUMCompra.AddItem colSiglas.Item(iIndice).sSigla
            SiglaUMEstoque.AddItem colSiglas.Item(iIndice).sSigla
            SiglaUMVenda.AddItem colSiglas.Item(iIndice).sSigla
            SiglaUMTrib.AddItem colSiglas.Item(iIndice).sSigla
            If colSiglas.Item(iIndice).sSigla = objClasseUM.sSiglaUMBase Then iIndice2 = iIndice
        Next

        SiglaUMCompra.ListIndex = iIndice2 - 1
        SiglaUMEstoque.ListIndex = iIndice2 - 1
        SiglaUMVenda.ListIndex = iIndice2 - 1
        SiglaUMTrib.ListIndex = iIndice2 - 1

    End If

    Carrega_CombosUM = SUCESSO

    Exit Function

Erro_Carrega_CombosUM:

    Carrega_CombosUM = Err

    Select Case Err

        Case 54128

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174936)

    End Select

    Exit Function

End Function

Private Sub SiglaUMCompra_Click()

Dim lErro As Long
Dim sSiglaUMCompra As String

On Error GoTo Erro_SiglaUMCompra_Click

    iAlterado = REGISTRO_ALTERADO

    'Se não selecionou nada --> Sai
    If SiglaUMCompra.ListIndex = -1 Then Exit Sub

    sSiglaUMCompra = SiglaUMCompra.Text

    'Verifica se Existe e Exibe De acordo com o Parametro
    lErro = SiglaUM_Exibe(sSiglaUMCompra, "NomeUMCompra")
    If lErro <> SUCESSO Then Error 54098

    Exit Sub

Erro_SiglaUMCompra_Click:

    Select Case Err

        Case 54098

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174937)

    End Select

    Exit Sub

End Sub

Private Sub SiglaUMEstoque_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub SiglaUMVenda_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub SiglaUMVenda_Click()

Dim lErro As Long
Dim sSiglaUMVenda As String

On Error GoTo Erro_SiglaUMVenda_Click

    iAlterado = REGISTRO_ALTERADO

    'Se não selecionou nada --> Sai
    If SiglaUMVenda.ListIndex = -1 Then Exit Sub

    sSiglaUMVenda = SiglaUMVenda.Text

    'Verifica se Existe e Exibe De acordo com o Parametro
    lErro = SiglaUM_Exibe(sSiglaUMVenda, "NomeUMVenda")
    If lErro <> SUCESSO Then Error 54099

    Exit Sub

Erro_SiglaUMVenda_Click:

    Select Case Err

        Case 54099

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174938)

    End Select

    Exit Sub

End Sub

Private Sub SiglaUMEstoque_Click()

Dim lErro As Long
Dim sSiglaUMEstoque As String

On Error GoTo Erro_SiglaUMEstoque_Click

    iAlterado = REGISTRO_ALTERADO

    'Se não tiver nada selecionado --> Sai
    If SiglaUMEstoque.ListIndex = -1 Then Exit Sub

    sSiglaUMEstoque = SiglaUMEstoque.Text

    'Verifica se Existe e Exibe De acordo com o Parametro
    lErro = SiglaUM_Exibe(sSiglaUMEstoque, "NomeUMEstoque")
    If lErro <> SUCESSO Then Error 54097

    Exit Sub

Erro_SiglaUMEstoque_Click:

    Select Case Err

        Case 54097

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174939)

    End Select

    Exit Sub

End Sub

Private Function Traz_TabUnidadesMedida_Tela(objTipoProduto As ClassTipoDeProduto) As Long

Dim lErro As Long
Dim objClasseUM As New ClassClasseUM
Dim iIndice As Integer

    'Se tiver Classe
    If objTipoProduto.iClasseUM <> 0 Then
        ClasseUM.Text = CStr(objTipoProduto.iClasseUM)
        'Teste se Existe No BD e Traz para Tela
        Call ClasseUM_Validate(bSGECancelDummy)
    End If

    'Seleciona nas combos as U.M. contidas no objTipoProduto
    If SiglaUMCompra.ListCount > 0 Then
        'U.M. de Estoque
        For iIndice = 0 To SiglaUMEstoque.ListCount - 1
            If SiglaUMEstoque.List(iIndice) = objTipoProduto.sSiglaUMEstoque Then
                SiglaUMEstoque.ListIndex = iIndice
                Exit For
            End If
        Next

        'U.M. de Compra
        For iIndice = 0 To SiglaUMCompra.ListCount - 1
            If SiglaUMCompra.List(iIndice) = objTipoProduto.sSiglaUMCompra Then
                SiglaUMCompra.ListIndex = iIndice
                Exit For
            End If
        Next

        'U.M. de Venda
        For iIndice = 0 To SiglaUMVenda.ListCount - 1
            If SiglaUMVenda.List(iIndice) = objTipoProduto.sSiglaUMVenda Then
                SiglaUMVenda.ListIndex = iIndice
                Exit For
            End If
        Next
        
        'U.M. de Trib
        For iIndice = 0 To SiglaUMTrib.ListCount - 1
            If SiglaUMTrib.List(iIndice) = objTipoProduto.sSiglaUMTrib Then
                SiglaUMTrib.ListIndex = iIndice
                Exit For
            End If
        Next

    End If

    Traz_TabUnidadesMedida_Tela = SUCESSO

    Exit Function

End Function

Private Function Move_TabTributacao_Memoria(objTipoProduto As ClassTipoDeProduto) As Long

Dim lErro As Long
Dim sContaFormatada As String
Dim iContaPreenchida As Integer
Dim iIndice As Integer

On Error GoTo Erro_Move_TabTributacao_Memoria

    objTipoProduto.sIPICodigo = ClasFiscIPI.Text
    objTipoProduto.sCEST = CEST.Text

    'Verifica se o IncideIPI esta selecionado
    If IncideIPI.Value = vbChecked Then

        'Recolhe os dados relacionadosao IPI
        If Len(Trim(AliquotaIPI.Text)) Then objTipoProduto.dIPIAliquota = CDbl(AliquotaIPI / 100)
        objTipoProduto.sIPICodDIPI = CodigoIPI.Text

    End If

    'Verifica se a Conta Contábil foi informada
    If Len(Trim(ContaContabil.ClipText)) > 0 Then
    
        'Guarda a conta corrente
        lErro = CF("Conta_Formata", ContaContabil.Text, sContaFormatada, iContaPreenchida)
        If lErro <> SUCESSO Then gError 66083
        
        objTipoProduto.sContaContabil = sContaFormatada
        
    End If
    
    'Verifica se a Conta Producao foi informada
    If Len(Trim(ContaProducao.ClipText)) > 0 Then
    
        'Guarda a conta corrente
        lErro = CF("Conta_Formata", ContaProducao.Text, sContaFormatada, iContaPreenchida)
        If lErro <> SUCESSO Then gError 66084
        
        objTipoProduto.sContaProducao = sContaFormatada
    
    End If
    
    objTipoProduto.iProdutoEspecifico = Codigo_Extrai(ProdutoEspecifico.Text)
    objTipoProduto.sISSQN = ISSQN.Text
    objTipoProduto.sGenero = Genero.Text
    objTipoProduto.iExTIPI = StrParaInt(ExTIPI.ClipText)
    
    objTipoProduto.iOrigem = OrigemMercadoria.ItemData(OrigemMercadoria.ListIndex)
    
    Move_TabTributacao_Memoria = SUCESSO

    Exit Function

Erro_Move_TabTributacao_Memoria:

    Move_TabTributacao_Memoria = Err

    Select Case Err

        Case 66083, 66084
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174940)

    End Select

    Exit Function

End Function

Private Sub ContaContabil_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ContaContabil_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaFormatada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim sContaMascarada As String

On Error GoTo Erro_ContaContabil_Validate

    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica_Modulo", ContaContabil.Text, ContaContabil.ClipText, objPlanoConta, MODULO_ESTOQUE)
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 66085

    If lErro = SUCESSO Then

        sContaFormatada = objPlanoConta.sConta

        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then gError 66086

        ContaContabil.PromptInclude = False
        ContaContabil.Text = sContaMascarada
        ContaContabil.PromptInclude = True

    'Se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then

        'Critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
        lErro = CF("Conta_Critica", ContaContabil.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
        If lErro <> SUCESSO And lErro <> 5700 Then gError 66087

        'Conta não cadastrada
        If lErro = 5700 Then gError 66088

    End If

    Exit Sub

Erro_ContaContabil_Validate:

    Cancel = True


    Select Case gErr

        Case 66085, 66087
    
        Case 66086
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            
        Case 66088
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", gErr, ContaContabil.Text)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174941)
    
    End Select

    Exit Sub
    
End Sub

Private Sub ContaProducao_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ContaProducao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaFormatada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim sContaMascarada As String

On Error GoTo Erro_ContaProducao_Validate

    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica_Modulo", ContaProducao.Text, ContaProducao.ClipText, objPlanoConta, MODULO_ESTOQUE)
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 66089

    If lErro = SUCESSO Then

        sContaFormatada = objPlanoConta.sConta

        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then gError 66090

        ContaProducao.PromptInclude = False
        ContaProducao.Text = sContaMascarada
        ContaProducao.PromptInclude = True

    'Se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then

        'Critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
        lErro = CF("Conta_Critica", ContaProducao.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
        If lErro <> SUCESSO And lErro <> 5700 Then gError 66091

        'Conta não cadastrada
        If lErro = 5700 Then gError 66092

    End If

    Exit Sub

Erro_ContaProducao_Validate:

    Cancel = True


    Select Case gErr

        Case 66089, 66091
    
        Case 66090
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            
        Case 66092
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", gErr, ContaProducao.Text)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174942)
    
    End Select

    Exit Sub
    
End Sub

Private Function Traz_TabTributacao_Tela(objTipoDeProduto As ClassTipoDeProduto) As Long
'Preenche o Tab Tributação com os dados do BD

Dim lErro As Long
Dim sContaMascarada As String

On Error GoTo Erro_Traz_TabTributacao_Tela

    ClasFiscIPI.PromptInclude = False
    ClasFiscIPI.Text = objTipoDeProduto.sIPICodigo
    ClasFiscIPI.PromptInclude = True

    CEST.PromptInclude = False
    CEST.Text = objTipoDeProduto.sCEST
    CEST.PromptInclude = True

    'Relacionados ao IPI
    If objTipoDeProduto.dIPIAliquota > 0 Or Len(Trim(objTipoDeProduto.sIPICodDIPI)) > 0 Then
        IncideIPI.Value = vbChecked
        AliquotaIPI.Enabled = True
        AliquotaIPI.Text = Format(objTipoDeProduto.dIPIAliquota * 100, "standard")
        CodigoIPI.Enabled = True
        CodigoIPI.Text = objTipoDeProduto.sIPICodDIPI
    Else
        IncideIPI = vbUnchecked
        AliquotaIPI.Text = ""
        AliquotaIPI.Enabled = False
        CodigoIPI.Text = ""
        CodigoIPI.Enabled = False
    End If

    If objTipoDeProduto.sContaContabil <> "" Then
    
        sContaMascarada = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(objTipoDeProduto.sContaContabil, sContaMascarada)
        If lErro <> SUCESSO Then gError 66095
        
    Else
    
        sContaMascarada = ""
        
    End If

    ContaContabil.PromptInclude = False
    ContaContabil.Text = sContaMascarada
    ContaContabil.PromptInclude = True

    If objTipoDeProduto.sContaProducao <> "" Then
        sContaMascarada = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(objTipoDeProduto.sContaProducao, sContaMascarada)
        If lErro <> SUCESSO Then gError 66096
    Else
        sContaMascarada = ""
    End If

    ContaProducao.PromptInclude = False
    ContaProducao.Text = sContaMascarada
    ContaProducao.PromptInclude = True
    
    If objTipoDeProduto.iExTIPI <> 0 Then
        ExTIPI.PromptInclude = False
        ExTIPI.Text = CStr(objTipoDeProduto.iExTIPI)
        ExTIPI.PromptInclude = True
    End If
    Genero.PromptInclude = False
    Genero.Text = Format(objTipoDeProduto.sGenero, Genero.Format)
    Genero.PromptInclude = True
    Call Genero_Validate(bSGECancelDummy)
    ISSQN.PromptInclude = False
    ISSQN.Text = Format(objTipoDeProduto.sISSQN, ISSQN.Format)
    ISSQN.PromptInclude = True
    Call ISSQN_Validate(bSGECancelDummy)
    ProdutoEspecifico.ListIndex = objTipoDeProduto.iProdutoEspecifico

    'Preenche a Origem na Tela
    Call Combo_Seleciona_ItemData(OrigemMercadoria, objTipoDeProduto.iOrigem)

    Traz_TabTributacao_Tela = SUCESSO

    Exit Function

Erro_Traz_TabTributacao_Tela:

    Traz_TabTributacao_Tela = gErr

    Select Case gErr

        Case 28586

        Case 66095
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objTipoDeProduto.sContaContabil)

        Case 66096
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objTipoDeProduto.sContaProducao)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174943)

    End Select

    Exit Function

End Function

Private Function Traz_ClasseUM_Tela(objClasseUM As ClassClasseUM)
'Traz Valores da Classe de UM para a Tela

Dim lErro As Long

On Error GoTo Erro_Traz_ClasseUM_Tela

    'Limpa o conteúdo das Combos e a Descrição da Classe
    SiglaUMEstoque.Clear
    SiglaUMCompra.Clear
    SiglaUMVenda.Clear
    NomeUMEstoque.Caption = ""
    NomeUMVenda.Caption = ""
    NomeUMCompra.Caption = ""
    DescricaoClasseUM.Caption = ""

    If objClasseUM.iClasse = 0 Then
        iClasseUMAlterada = 0
        iClasseUMAnterior = 0
        Exit Function
    End If

    'Coloca na Tela a Descrição
    DescricaoClasseUM.Caption = objClasseUM.sDescricao

    'Carrega as Combos que tem Pendencia Com a Classe: Estoque, Venda, Compra
    lErro = Carrega_CombosUM(objClasseUM)
    If lErro <> SUCESSO Then Error 54129

    Traz_ClasseUM_Tela = SUCESSO

    Exit Function

Erro_Traz_ClasseUM_Tela:

    Traz_ClasseUM_Tela = Err

    Select Case Err

        Case 54129

        Case 54127

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174944)

    End Select

    Exit Function

End Function

Private Function SiglaUM_Exibe(sSiglaUM As String, sNome As String) As Long

Dim lErro As Long
Dim objUM As New ClassUnidadeDeMedida

On Error GoTo Erro_SiglaUM_Exibe

    'Se não selecionou nada --> Sai
    If SiglaUMCompra.ListIndex = -1 Then Exit Function

    objUM.iClasse = CInt(ClasseUM.Text)
    objUM.sSigla = sSiglaUM

    'Lê a Sigla de Unidade de Medida
    lErro = CF("UM_Le", objUM)
    If lErro <> SUCESSO And lErro <> 23775 Then Error 54100

    'Se não encontrar --> erro
    If lErro = 23775 Then Error 54101

    If sNome = "NomeUMEstoque" Then
        NomeUMEstoque.Caption = objUM.sNome
    ElseIf sNome = "NomeUMVenda" Then
        NomeUMVenda.Caption = objUM.sNome
    ElseIf sNome = "NomeUMTrib" Then
        NomeUMTrib.Caption = objUM.sNome
    Else
        NomeUMCompra.Caption = objUM.sNome
    End If

    SiglaUM_Exibe = SUCESSO

    Exit Function

Erro_SiglaUM_Exibe:

    SiglaUM_Exibe = Err

    Select Case Err

        Case 54100

        Case 54101
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UNIDADE_MEDIDA_NAO_CADASTRADA", Err, objUM.iClasse, objUM.sSiglaUMBase)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174945)

    End Select

    Exit Function

End Function

Private Sub TempoRessupMax_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

'Traz dados default da tabela ComprasConfig para tela
Private Function Traz_ComprasConfig_Tela()

Dim lErro As Long
Dim objConfiguraCOM As New ClassConfiguraCOM

On Error GoTo Erro_Traz_ComprasConfig_Tela
    
    'Lê dados de ComprasConfig
    lErro = CF("ComprasConfig_Le", objConfiguraCOM)
    If lErro <> SUCESSO Then Error 54144
    
    'Traz dados de Consumo para tela
    MesesConsumoMedio.Text = objConfiguraCOM.iMesesConsumoMedio
    
    'Traz dados de Estoque de Segurança para tela
    ConsumoMedioMax.Text = (objConfiguraCOM.dConsumoMedioMax) * 100
    TempoRessupMax.Text = objConfiguraCOM.dTempoRessupMax * 100

    'Traz dados de Resíduo para tela
    Residuo.Text = (objConfiguraCOM.dResiduo) * 100

    'Traz dados de Recebimento
    NaoTemFaixaReceb.Value = objConfiguraCOM.iTemFaixaReceb
    
    If NaoTemFaixaReceb.Value = vbChecked Then
        PercentMaisReceb.Text = ""
        PercentMenosReceb.Text = ""
    Else
        PercentMaisReceb.Text = Format((objConfiguraCOM.dPercentMaisReceb) * 100, "Fixed")
        PercentMenosReceb.Text = Format((objConfiguraCOM.dPercentMenosReceb) * 100, "Fixed")
    End If
    
    RecebForaFaixa(objConfiguraCOM.iRecebForaFaixa).Value = True
    
    'Traz dados de Cotações Anteriores
    NaoConsideraQuantCotacaoAnterior.Value = objConfiguraCOM.iConsideraQuantCotacaoAnterior
    
    If NaoConsideraQuantCotacaoAnterior.Value = vbChecked Then
        PercentMaisQuantCotacaoAnterior.Text = ""
        PercentMenosQuantCotacaoAnterior.Text = ""
    Else
        PercentMaisQuantCotacaoAnterior.Text = Format((objConfiguraCOM.dPercentMaisQuantCotacaoAnterior) * 100, "Fixed")
        PercentMenosQuantCotacaoAnterior.Text = Format((objConfiguraCOM.dPercentMenosQuantCotacaoAnterior) * 100, "Fixed")
    End If
    
    Traz_ComprasConfig_Tela = SUCESSO
    
    Exit Function

Erro_Traz_ComprasConfig_Tela:

    Traz_ComprasConfig_Tela = Err
    
    Select Case Err

        Case 54144  'Tratado na rotina chamada
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174946)

    End Select

    Exit Function
    
End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Tipos de Produto"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "TipoProduto"
    
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


Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
        
    If KeyCode = KEYCODE_BROWSER Then
    
        If Me.ActiveControl Is ContaContabil Then
            Call ContaContabilLabel_Click
        End If
        
        If Me.ActiveControl Is ContaProducao Then
            Call LabelContaProducao_Click
        End If
        
        If Me.ActiveControl Is ClasseUM Then
            Call LblClasseUM_Click
        End If
        
        If Me.ActiveControl Is Codigo Then
            Call LblTipo_Click
        End If
       
        If Me.ActiveControl Is ClasFiscIPI Then
            Call LabelClassificacaoFiscal_Click
        End If
        
    End If

End Sub

Private Sub Label2_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label2(Index), Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2(Index), Button, Shift, X, Y)
End Sub

Private Sub LblTipo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblTipo, Source, X, Y)
End Sub

Private Sub LblTipo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblTipo, Button, Shift, X, Y)
End Sub

Private Sub Label23_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label23, Source, X, Y)
End Sub

Private Sub Label23_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label23, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label16_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label16, Source, X, Y)
End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label16, Button, Shift, X, Y)
End Sub

Private Sub Label20_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label20, Source, X, Y)
End Sub

Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label20, Button, Shift, X, Y)
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub

Private Sub LblUMEstoque_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblUMEstoque, Source, X, Y)
End Sub

Private Sub LblUMEstoque_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblUMEstoque, Button, Shift, X, Y)
End Sub

Private Sub LblUMCompra_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblUMCompra, Source, X, Y)
End Sub

Private Sub LblUMCompra_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblUMCompra, Button, Shift, X, Y)
End Sub

Private Sub NomeUMCompra_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NomeUMCompra, Source, X, Y)
End Sub

Private Sub NomeUMCompra_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NomeUMCompra, Button, Shift, X, Y)
End Sub

Private Sub LblUMVenda_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblUMVenda, Source, X, Y)
End Sub

Private Sub LblUMVenda_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblUMVenda, Button, Shift, X, Y)
End Sub

Private Sub NomeUMVenda_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NomeUMVenda, Source, X, Y)
End Sub

Private Sub NomeUMVenda_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NomeUMVenda, Button, Shift, X, Y)
End Sub

Private Sub NomeUMEstoque_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NomeUMEstoque, Source, X, Y)
End Sub

Private Sub NomeUMEstoque_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NomeUMEstoque, Button, Shift, X, Y)
End Sub

Private Sub LblClasseUM_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblClasseUM, Source, X, Y)
End Sub

Private Sub LblClasseUM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblClasseUM, Button, Shift, X, Y)
End Sub

Private Sub DescricaoClasseUM_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescricaoClasseUM, Source, X, Y)
End Sub

Private Sub DescricaoClasseUM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescricaoClasseUM, Button, Shift, X, Y)
End Sub

Private Sub ContaContabilLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContaContabilLabel, Source, X, Y)
End Sub

Private Sub ContaContabilLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContaContabilLabel, Button, Shift, X, Y)
End Sub

Private Sub LabelContaProducao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelContaProducao, Source, X, Y)
End Sub

Private Sub LabelContaProducao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelContaProducao, Button, Shift, X, Y)
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

Private Sub ClasFiscIPI_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objClassificacaoFiscal As New ClassClassificacaoFiscal
Dim vbMsgRes As VbMsgBoxResult
Dim sCEST As String

On Error GoTo Erro_ClasFiscIPI_Validate

    'Verifica se o cmpo classificação fiscal está preenchido
    If Len(Trim(ClasFiscIPI.Text)) = 0 Then Exit Sub
    
    objClassificacaoFiscal.sCodigo = ClasFiscIPI.ClipText

    'Verifica se existe a Classificação Fiscal informada
    lErro = CF("ClassificacaoFiscal_Le", objClassificacaoFiscal)
    If lErro <> SUCESSO And lErro <> 123494 Then gError 125018
    
    'Se não existe, então pergunta se deseja criar
    If lErro = 123494 Then gError 125019
    
    lErro = CF("NCM_Le_CEST", objClassificacaoFiscal.sCodigo, sCEST)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 125018
    
    If lErro = SUCESSO Then
        CEST.PromptInclude = False
        CEST.Text = sCEST
        CEST.PromptInclude = True
        
        Call CEST_Validate(bSGECancelDummy)
    End If
    
    Exit Sub
    
Erro_ClasFiscIPI_Validate:

    Cancel = True

    Select Case gErr
    
        Case 125018
        
        Case 125019
        
            'Pergunta se deseja criar a Classificação Fiscal
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLASSIFICACAOFISCAL")
            
            If vbMsgRes = vbYes Then
            
                'Chama a tela para cadastrar uma nova Classificação Fiscal
                Call Chama_Tela("ClassificacaoFiscal", objClassificacaoFiscal)
                
            End If
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174947)

    End Select

    Exit Sub

End Sub

Private Sub LabelClassificacaoFiscal_Click()

Dim lErro As Long
Dim objClassificacaoFiscal As New ClassClassificacaoFiscal
Dim colSelecao As New Collection

On Error GoTo Erro_LabelClassificacaoFiscal_Click

    'Preenche na memória o Código passado
    If Len(Trim(ClasFiscIPI.ClipText)) > 0 Then objClassificacaoFiscal.sCodigo = ClasFiscIPI.ClipText

    Call Chama_Tela("ClassificacaoFiscalLista", colSelecao, objClassificacaoFiscal, objEventoClasFiscIPI)

    Exit Sub
    
Erro_LabelClassificacaoFiscal_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174948)

    End Select

    Exit Sub

End Sub

Private Sub objEventoClasFiscIPI_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objClassificacaoFiscal As New ClassClassificacaoFiscal
Dim bCancel As Boolean
    
On Error GoTo Erro_objEventoClasFiscIPI_evSelecao
    
    Set objClassificacaoFiscal = obj1

    'lErro = CF("ClassificacaoFiscal_Le", objClassificacaoFiscal)
    'If lErro <> SUCESSO And lErro <> 123494 Then gError 125020

    'If lErro = 123494 Then gError 125021

    'Preenche o Cliente com o Cliente selecionado
    ClasFiscIPI.PromptInclude = False
    ClasFiscIPI.Text = objClassificacaoFiscal.sCodigo
    ClasFiscIPI.PromptInclude = True
    Call ClasFiscIPI_Validate(bSGECancelDummy)

    Me.Show

    iAlterado = 0
    
    Exit Sub

Erro_objEventoClasFiscIPI_evSelecao:

    Select Case gErr
    
        Case 125020
        
        Case 125021
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CLASSIFICACAOFISCAL_NAO_EXISTENTE", gErr, objClassificacaoFiscal.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174949)

    End Select

    Exit Sub

End Sub

Private Sub ExTIPI_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub SiglaUMTrib_Click()

Dim lErro As Long
Dim sSiglaUMTrib As String

On Error GoTo Erro_SiglaUMTrib_Click

    iAlterado = REGISTRO_ALTERADO

    'Se não selecionou nada --> Sai
    If SiglaUMTrib.ListIndex = -1 Then Exit Sub
    
    sSiglaUMTrib = SiglaUMTrib.Text
        
    'Verifica se Existe e Exibe De acordo com o Parametro
    lErro = SiglaUM_Exibe(sSiglaUMTrib, "NomeUMTrib")
    If lErro <> SUCESSO Then gError 64371
        
    Exit Sub

Erro_SiglaUMTrib_Click:

    Select Case gErr

        Case 64371
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165505)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoEspecifico_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Genero_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ISSQN_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ISSQN_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objISSQN As New ClassISSQN
    
On Error GoTo Erro_ISSQN_Validate

    If Len(Trim(ISSQN.ClipText)) > 0 Then
    
        objISSQN.sCodigo = ISSQN.ClipText
        
        lErro = CF("ISSQN_Le", objISSQN)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 202773
        If lErro <> SUCESSO Then gError 202774
        
        DescISSQN.Caption = objISSQN.sDescricao
        CodServNFe.Caption = CStr(objISSQN.lCodServNFe)
        
    Else
    
        DescISSQN.Caption = ""
        CodServNFe.Caption = ""
        
    End If
    
    Exit Sub
    
Erro_ISSQN_Validate:

    Cancel = True

    Select Case gErr
    
        Case 202773
        
        Case 202774
            Call Rotina_Erro(vbOKOnly, "ERRO_ISSQN_NAO_CADASTRADO", gErr, objISSQN.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202775)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Genero_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProdutoGenero As New ClassProdutoGenero
    
On Error GoTo Erro_Genero_Validate

    If Len(Trim(Genero.ClipText)) > 0 Then
    
        objProdutoGenero.sCodigo = Genero.Text
        
        lErro = CF("ProdutoGenero_Le", objProdutoGenero)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 202776
        If lErro <> SUCESSO Then gError 202777
        
        DescGenero.Caption = objProdutoGenero.sDescricao
        
    Else
        DescGenero.Caption = ""
    End If
    
    Exit Sub
    
Erro_Genero_Validate:

    Cancel = True

    Select Case gErr
    
        Case 202776
        
        Case 202777
            Call Rotina_Erro(vbOKOnly, "ERRO_GENERO_NAO_CADASTRADO", gErr, objProdutoGenero.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202778)
    
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoISSQN_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objISSQN As ClassISSQN
    
On Error GoTo Erro_objEventoISSQN_evSelecao
    
    Set objISSQN = obj1

    ISSQN.PromptInclude = False
    ISSQN.Text = objISSQN.sCodigo
    ISSQN.PromptInclude = True
    Call ISSQN_Validate(bSGECancelDummy)

    Me.Show
   
    Exit Sub

Erro_objEventoISSQN_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165559)

    End Select

    Exit Sub

End Sub

Private Sub objEventoGenero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProdutoGenero As ClassProdutoGenero
    
On Error GoTo Erro_objEventoGenero_evSelecao
    
    Set objProdutoGenero = obj1

    Genero.PromptInclude = False
    Genero.Text = objProdutoGenero.sCodigo
    Genero.PromptInclude = True
    Call Genero_Validate(bSGECancelDummy)

    Me.Show
   
    Exit Sub

Erro_objEventoGenero_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165559)

    End Select

    Exit Sub

End Sub

Private Sub LabelGenero_Click()

Dim objProdutoGenero As New ClassProdutoGenero
Dim colSelecao As Collection

    objProdutoGenero.sCodigo = Genero.ClipText

    Call Chama_Tela("ProdutoGeneroLista", colSelecao, objProdutoGenero, objEventoGenero)

End Sub

Private Sub LabelISSQN_Click()

Dim objISSQN As New ClassISSQN
Dim colSelecao As Collection

    objISSQN.sCodigo = ISSQN.ClipText

    Call Chama_Tela("ISSQNLista", colSelecao, objISSQN, objEventoISSQN, , "Codigo")

End Sub

Private Sub LabelCodServNFe_Click()

Dim objISSQN As New ClassISSQN
Dim colSelecao As Collection

    objISSQN.lCodServNFe = StrParaLong(CodServNFe.Caption)

    Call Chama_Tela("ISSQNLista", colSelecao, objISSQN, objEventoISSQN, , "CodServNFe")

End Sub

Private Sub LabelCEST_Click()

Dim lErro As Long
Dim objCEST As New ClassCEST
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCEST_Click

    'Preenche na memória o Código passado
    If Len(Trim(CEST.ClipText)) > 0 Then objCEST.sCodigo = CEST.ClipText
    
    If Len(Trim(ClasFiscIPI.ClipText)) > 0 Then
    
        colSelecao.Add ClasFiscIPI.ClipText

        Call Chama_Tela("CESTLista", colSelecao, objCEST, objEventoCEST, "? LIKE (NCM + '%') AND NCM <> ''")
        
    Else
    
        Call Chama_Tela("CESTLista", colSelecao, objCEST, objEventoCEST, "NCM = ''")
    
    End If

    Exit Sub
    
Erro_LabelCEST_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216098)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoCEST_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCEST As New ClassCEST
Dim bCancel As Boolean
    
On Error GoTo Erro_objEventoCEST_evSelecao
    
    Set objCEST = obj1

    'Preenche o Cliente com o Cliente selecionado
    CEST.PromptInclude = False
    CEST.Text = objCEST.sCodigo
    CEST.PromptInclude = True
    
    Call CEST_Validate(bSGECancelDummy)

    Me.Show

    iAlterado = 0
    
    Exit Sub

Erro_objEventoCEST_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216099)

    End Select

    Exit Sub

End Sub

Private Sub CEST_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CEST_GotFocus()
    Call MaskEdBox_TrataGotFocus(CEST, iAlterado)
End Sub

Private Sub CEST_Validate(Cancel As Boolean)
Dim lErro As Long
Dim objCEST As New ClassCEST
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_CEST_Validate

    'Verifica se o cmpo classificação fiscal está preenchido
    If Len(Trim(CEST.ClipText)) = 0 Then Exit Sub
    
    objCEST.sCodigo = CEST.ClipText
    objCEST.sNCM = ClasFiscIPI.ClipText

    'Verifica se existe a Classificação Fiscal informada
    lErro = CF("CEST_Le", objCEST)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
    
    'Se não existe, então pergunta se deseja criar
    If lErro = ERRO_LEITURA_SEM_DADOS Then gError 216096
    
    Exit Sub
    
Erro_CEST_Validate:

    Cancel = True

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 216096
            Call Rotina_Erro(vbOKOnly, "ERRO_CEST_NCM_NAO_CADASTRADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 216097)

    End Select

    Exit Sub
    
End Sub
