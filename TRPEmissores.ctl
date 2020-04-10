VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl TRPEmissores 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4365
      Index           =   3
      Left            =   390
      TabIndex        =   31
      Top             =   855
      Visible         =   0   'False
      Width           =   8775
      Begin VB.Frame SSFrame5 
         Height          =   510
         Left            =   90
         TabIndex        =   34
         Top             =   90
         Width           =   8610
         Begin VB.Label Label11 
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
            Height          =   210
            Left            =   165
            TabIndex        =   36
            Top             =   180
            Width           =   1110
         End
         Begin VB.Label Fornecedor 
            Height          =   210
            Index           =   3
            Left            =   1380
            TabIndex        =   35
            Top             =   180
            Width           =   6945
         End
      End
      Begin TelasFATTRP.TabEndereco TabEnd 
         Height          =   3495
         Index           =   0
         Left            =   0
         TabIndex        =   117
         Top             =   795
         Width           =   8490
         _ExtentX        =   14975
         _ExtentY        =   6165
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4395
      Index           =   1
      Left            =   135
      TabIndex        =   27
      Top             =   780
      Width           =   9150
      Begin VB.ComboBox Superior 
         Height          =   315
         Left            =   1455
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2595
         Width           =   3195
      End
      Begin VB.Frame Frame7 
         Caption         =   "Categorias"
         Height          =   3300
         Left            =   4875
         TabIndex        =   103
         Top             =   1080
         Width           =   4155
         Begin VB.ComboBox ComboCategoriaFornecedor 
            Height          =   315
            Left            =   1800
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   540
            Width           =   1545
         End
         Begin VB.ComboBox ComboCategoriaFornecedorItem 
            Height          =   315
            Left            =   3000
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   540
            Width           =   1635
         End
         Begin MSFlexGridLib.MSFlexGrid GridCategoria 
            Height          =   2520
            Left            =   165
            TabIndex        =   9
            Top             =   330
            Width           =   3810
            _ExtentX        =   6720
            _ExtentY        =   4445
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
      Begin VB.TextBox RazaoSocial 
         Height          =   300
         Left            =   1455
         MaxLength       =   40
         TabIndex        =   3
         Top             =   690
         Width           =   7500
      End
      Begin VB.CheckBox Ativo 
         Caption         =   "Ativo"
         Height          =   225
         Left            =   2850
         TabIndex        =   2
         Top             =   240
         Value           =   1  'Checked
         Width           =   930
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   2295
         Picture         =   "TRPEmissores.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Numeração Automática"
         Top             =   210
         Width           =   300
      End
      Begin VB.TextBox Observacao 
         Height          =   1320
         Left            =   1470
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   3060
         Width           =   3165
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   1470
         TabIndex        =   0
         Top             =   210
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         Mask            =   "99999999"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NomeReduzido 
         Height          =   315
         Left            =   1470
         TabIndex        =   4
         Top             =   1170
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CGC 
         Height          =   315
         Left            =   1455
         TabIndex        =   6
         Top             =   2130
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   14
         Mask            =   "99999999999999"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Cliente 
         Height          =   300
         Left            =   1470
         TabIndex        =   5
         Top             =   1650
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label CodigoForn 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   6120
         TabIndex        =   116
         Top             =   225
         Width           =   1320
      End
      Begin VB.Label LabelCodigo 
         Caption         =   "Código de Fornecedor:"
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
         Left            =   4125
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   115
         Top             =   255
         Width           =   2160
      End
      Begin VB.Label LabelCliente 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   735
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   114
         Top             =   1695
         Width           =   660
      End
      Begin VB.Label LabelSuperior 
         AutoSize        =   -1  'True
         Caption         =   "Superior:"
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
         Left            =   615
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   104
         Top             =   2625
         Width           =   780
      End
      Begin VB.Label Label35 
         Caption         =   "CPF:"
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
         Left            =   975
         TabIndex        =   102
         Top             =   2175
         Width           =   555
      End
      Begin VB.Label LabelObservacao 
         AutoSize        =   -1  'True
         Caption         =   "Observação:"
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
         Left            =   300
         TabIndex        =   73
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label LabelNomeReduzido 
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
         Left            =   -15
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   72
         Top             =   1230
         Width           =   1410
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   840
         TabIndex        =   71
         Top             =   750
         Width           =   555
      End
      Begin VB.Label LabelCodigoEmi 
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
         Left            =   735
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   70
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4290
      Index           =   2
      Left            =   135
      TabIndex        =   74
      Top             =   930
      Visible         =   0   'False
      Width           =   9090
      Begin VB.Frame Frame4 
         Caption         =   "Contabilidade"
         Height          =   2265
         Left            =   180
         TabIndex        =   87
         Top             =   780
         Width           =   4200
         Begin MSMask.MaskEdBox ContaContabil 
            Height          =   315
            Left            =   2430
            TabIndex        =   17
            Top             =   420
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ContaFornConsig 
            Height          =   315
            Left            =   2430
            TabIndex        =   18
            Top             =   1035
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ContaDespesa 
            Height          =   315
            Left            =   2430
            TabIndex        =   19
            Top             =   1620
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label ContaContabilLabel 
            AutoSize        =   -1  'True
            Caption         =   "Conta de Fornecedores:"
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
            TabIndex        =   90
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label ContaFornConsigLabel 
            AutoSize        =   -1  'True
            Caption         =   "Conta de Fornecedores em Consignação:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   315
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   89
            Top             =   997
            Width           =   2055
            WordWrap        =   -1  'True
         End
         Begin VB.Label LabelContaDespesa 
            AutoSize        =   -1  'True
            Caption         =   "Conta Estoque/Despesa:"
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
            TabIndex        =   88
            Top             =   1680
            Width           =   2145
         End
      End
      Begin VB.Frame SSFrame6 
         Caption         =   "Conta Corrente"
         Height          =   690
         Left            =   180
         TabIndex        =   83
         Top             =   3285
         Width           =   8520
         Begin MSMask.MaskEdBox ContaCorrente 
            Height          =   315
            Left            =   6330
            TabIndex        =   26
            Top             =   225
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   14
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Agencia 
            Height          =   315
            Left            =   3855
            TabIndex        =   25
            Top             =   225
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   7
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Banco 
            Height          =   315
            Left            =   1680
            TabIndex        =   24
            Top             =   225
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   3
            Mask            =   "999"
            PromptChar      =   " "
         End
         Begin VB.Label Label22 
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
            Left            =   5550
            TabIndex        =   86
            Top             =   270
            Width           =   720
         End
         Begin VB.Label Label25 
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
            Height          =   195
            Left            =   3045
            TabIndex        =   85
            Top             =   270
            Width           =   765
         End
         Begin VB.Label LabelBanco 
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
            Height          =   195
            Left            =   1005
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   84
            Top             =   270
            Width           =   615
         End
      End
      Begin VB.Frame SSFrame8 
         Height          =   510
         Left            =   180
         TabIndex        =   80
         Top             =   0
         Width           =   8490
         Begin VB.Label Fornecedor 
            Height          =   210
            Index           =   2
            Left            =   1395
            TabIndex        =   82
            Top             =   180
            Width           =   6810
         End
         Begin VB.Label Label24 
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
            Height          =   210
            Left            =   165
            TabIndex        =   81
            Top             =   180
            Width           =   1110
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Pagamento"
         Height          =   2265
         Left            =   4470
         TabIndex        =   75
         Top             =   780
         Width           =   4200
         Begin VB.ComboBox CondicaoPagto 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1140
            TabIndex        =   21
            Top             =   795
            Width           =   2070
         End
         Begin VB.ComboBox TipoCobranca 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   1800
            Width           =   2865
         End
         Begin VB.ComboBox TipoFrete 
            Height          =   315
            ItemData        =   "TRPEmissores.ctx":00EA
            Left            =   1140
            List            =   "TRPEmissores.ctx":00F4
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   1305
            Width           =   1125
         End
         Begin MSMask.MaskEdBox Desconto 
            Height          =   315
            Left            =   1140
            TabIndex        =   20
            Top             =   285
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin VB.Label CondicaoPagtoLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Condição:"
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
            Left            =   135
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   79
            Top             =   855
            Width           =   975
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Desconto:"
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
            TabIndex        =   78
            Top             =   345
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cobrança:"
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
            Left            =   225
            TabIndex        =   77
            Top             =   1860
            Width           =   885
         End
         Begin VB.Label Label39 
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
            Left            =   165
            TabIndex        =   76
            Top             =   1365
            Width           =   945
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4245
      Index           =   5
      Left            =   165
      TabIndex        =   32
      Top             =   915
      Visible         =   0   'False
      Width           =   9225
      Begin VB.Frame Frame2 
         Caption         =   "Pedidos de Compra"
         Height          =   1365
         Left            =   4830
         TabIndex        =   62
         Top             =   2490
         Width           =   3825
         Begin VB.Frame SSFrame4 
            Caption         =   "Atraso de Entrega"
            Height          =   600
            Left            =   150
            TabIndex        =   63
            Top             =   600
            Width           =   3585
            Begin VB.Label Label26 
               Caption         =   "Média:"
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
               Left            =   465
               TabIndex        =   67
               Top             =   300
               Width           =   585
            End
            Begin VB.Label Label27 
               Caption         =   "Maior:"
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
               Left            =   2055
               TabIndex        =   66
               Top             =   300
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label MediaAtraso 
               Caption         =   "0"
               Height          =   210
               Left            =   1110
               TabIndex        =   65
               Top             =   285
               Width           =   510
            End
            Begin VB.Label MaiorAtraso 
               Caption         =   "0"
               Height          =   210
               Left            =   2655
               TabIndex        =   64
               Top             =   300
               Visible         =   0   'False
               Width           =   510
            End
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Em Aberto:"
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
            Left            =   240
            TabIndex        =   69
            Top             =   300
            Width           =   945
         End
         Begin VB.Label PedidosEmAberto 
            Caption         =   "0"
            Height          =   210
            Left            =   1275
            TabIndex        =   68
            Top             =   315
            Width           =   585
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Devoluções"
         Height          =   1110
         Left            =   4860
         TabIndex        =   41
         Top             =   975
         Width           =   3780
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Última:"
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
            Left            =   1275
            TabIndex        =   47
            Top             =   825
            Width           =   600
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Valor Acumulado:"
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
            Left            =   375
            TabIndex        =   46
            Top             =   255
            Width           =   1500
         End
         Begin VB.Label ValorAcumuladoDevolucoes 
            Caption         =   "0,00"
            Height          =   210
            Left            =   1980
            TabIndex        =   45
            Top             =   270
            Width           =   1575
         End
         Begin VB.Label Label19 
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
            Left            =   1155
            TabIndex        =   44
            Top             =   540
            Width           =   720
         End
         Begin VB.Label NumTotalDevolucoes 
            Caption         =   "0"
            Height          =   210
            Left            =   1980
            TabIndex        =   43
            Top             =   540
            Width           =   585
         End
         Begin VB.Label DataUltDevolucao 
            Caption         =   "  /  /    "
            Height          =   210
            Left            =   1980
            TabIndex        =   42
            Top             =   825
            Width           =   1170
         End
      End
      Begin VB.Frame SSFrame7 
         Height          =   510
         Left            =   240
         TabIndex        =   38
         Top             =   15
         Width           =   8520
         Begin VB.Label Fornecedor 
            Height          =   210
            Index           =   5
            Left            =   1335
            TabIndex        =   40
            Top             =   180
            Width           =   6750
         End
         Begin VB.Label Label12 
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
            Height          =   210
            Left            =   180
            TabIndex        =   39
            Top             =   180
            Width           =   1110
         End
      End
      Begin VB.Frame SSFrame1 
         Caption         =   "Compras"
         Height          =   1920
         Left            =   240
         TabIndex        =   48
         Top             =   1950
         Width           =   4200
         Begin VB.Label DataUltimaCompra 
            Caption         =   "  /  /    "
            Height          =   210
            Left            =   1830
            TabIndex        =   58
            Top             =   1530
            Width           =   1170
         End
         Begin VB.Label DataPrimeiraCompra 
            Caption         =   "  /  /    "
            Height          =   210
            Left            =   1830
            TabIndex        =   57
            Top             =   1200
            Width           =   1170
         End
         Begin VB.Label ValorAcumuladoCompras 
            Caption         =   "0,00"
            Height          =   210
            Left            =   1830
            TabIndex        =   56
            Top             =   900
            Width           =   1575
         End
         Begin VB.Label MediaCompra 
            Caption         =   "0,00"
            Height          =   210
            Left            =   1830
            TabIndex        =   55
            Top             =   600
            Width           =   1410
         End
         Begin VB.Label NumeroCompras 
            Caption         =   "0"
            Height          =   210
            Left            =   1830
            TabIndex        =   54
            Top             =   270
            Width           =   585
         End
         Begin VB.Label Label18 
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
            Left            =   1035
            TabIndex        =   53
            Top             =   255
            Width           =   720
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Média:"
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
            Left            =   1170
            TabIndex        =   52
            Top             =   585
            Width           =   585
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Última:"
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
            Left            =   1155
            TabIndex        =   51
            Top             =   1530
            Width           =   600
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Primeira:"
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
            TabIndex        =   50
            Top             =   1200
            Width           =   750
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Valor Acumulado:"
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
            Left            =   240
            TabIndex        =   49
            Top             =   900
            Width           =   1500
         End
      End
      Begin VB.Frame SSFrame3 
         Caption         =   "Saldos"
         Height          =   600
         Left            =   240
         TabIndex        =   59
         Top             =   1005
         Width           =   4200
         Begin VB.Label SaldoTitulos 
            Caption         =   "0,00"
            Height          =   210
            Left            =   1815
            TabIndex        =   61
            Top             =   300
            Width           =   1575
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Em Títulos:"
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
            Left            =   780
            TabIndex        =   60
            Top             =   285
            Width           =   990
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4245
      Index           =   4
      Left            =   180
      TabIndex        =   105
      Top             =   870
      Visible         =   0   'False
      Width           =   9225
      Begin VB.Frame Frame12 
         Caption         =   "Exceções de comissão por produto"
         Height          =   3210
         Left            =   240
         TabIndex        =   109
         Top             =   1035
         Width           =   8490
         Begin VB.CommandButton BotaoProdutos 
            Caption         =   "Produtos"
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
            Left            =   375
            TabIndex        =   30
            Top             =   2745
            Width           =   1815
         End
         Begin VB.TextBox DescricaoProduto 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   4215
            MaxLength       =   250
            TabIndex        =   111
            Top             =   810
            Width           =   3720
         End
         Begin MSMask.MaskEdBox PercComissProd 
            Height          =   240
            Left            =   960
            TabIndex        =   110
            Top             =   810
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
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
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Produto 
            Height          =   240
            Left            =   2385
            TabIndex        =   112
            Top             =   840
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridExcecoes 
            Height          =   330
            Left            =   360
            TabIndex        =   29
            Top             =   285
            Width           =   7665
            _ExtentX        =   13520
            _ExtentY        =   582
            _Version        =   393216
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
      End
      Begin VB.Frame Frame11 
         Height          =   510
         Left            =   240
         TabIndex        =   106
         Top             =   15
         Width           =   8520
         Begin VB.Label Label40 
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
            Height          =   210
            Left            =   180
            TabIndex        =   108
            Top             =   180
            Width           =   1110
         End
         Begin VB.Label Fornecedor 
            Height          =   210
            Index           =   4
            Left            =   1395
            TabIndex        =   107
            Top             =   180
            Width           =   6750
         End
      End
      Begin MSMask.MaskEdBox PercComi 
         Height          =   315
         Left            =   2760
         TabIndex        =   28
         Top             =   660
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "% padrão de comissão:"
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
         Left            =   690
         TabIndex        =   113
         Top             =   720
         Width           =   1950
      End
   End
   Begin VB.CommandButton BotaoWebAcesso 
      Caption         =   "Definir Acesso Web"
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
      Left            =   60
      TabIndex        =   12
      Top             =   5355
      Width           =   1995
   End
   Begin VB.Frame Frame6 
      Caption         =   "Frame6"
      Height          =   435
      Left            =   135
      TabIndex        =   91
      Top             =   -135
      Visible         =   0   'False
      Width           =   6630
      Begin VB.CommandButton Filiais 
         Height          =   525
         Left            =   0
         Picture         =   "TRPEmissores.ctx":0102
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   0
         Width           =   1620
      End
      Begin VB.CheckBox GeraICMS 
         Caption         =   "Pode Gerar Credito ICMS"
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
         Left            =   6015
         TabIndex        =   93
         Top             =   1740
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.TextBox Observacao2 
         Height          =   315
         Left            =   1770
         MaxLength       =   150
         MultiLine       =   -1  'True
         TabIndex        =   92
         Top             =   1710
         Width           =   3945
      End
      Begin MSMask.MaskEdBox InscricaoEstadual 
         Height          =   315
         Left            =   1770
         TabIndex        =   94
         Top             =   0
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   18
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox InscricaoMunicipal 
         Height          =   315
         Left            =   1770
         TabIndex        =   95
         Top             =   570
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   18
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox InscricaoINSS 
         Height          =   300
         Left            =   1770
         TabIndex        =   96
         Top             =   1155
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   14
         Mask            =   "999.999.999-99"
         PromptChar      =   " "
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição INSS:"
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
         Left            =   375
         TabIndex        =   100
         Top             =   1215
         Width           =   1350
      End
      Begin VB.Label Label36 
         Caption         =   "Inscrição Municipal:"
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
         Left            =   0
         TabIndex        =   99
         Top             =   615
         Width           =   1725
      End
      Begin VB.Label Label34 
         Caption         =   "Inscrição Estadual:"
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
         Left            =   90
         TabIndex        =   98
         Top             =   45
         Width           =   1650
      End
      Begin VB.Label Label32 
         Caption         =   "Observação:"
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
         Left            =   615
         TabIndex        =   97
         Top             =   1755
         Width           =   1125
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7740
      ScaleHeight     =   495
      ScaleWidth      =   1620
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   120
      Width           =   1680
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "TRPEmissores.ctx":0EA4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "TRPEmissores.ctx":0FFE
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1110
         Picture         =   "TRPEmissores.ctx":1530
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "TRPEmissores.ctx":16AE
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Excluir"
         Top             =   90
         Visible         =   0   'False
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   4815
      Left            =   60
      TabIndex        =   37
      Top             =   450
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8493
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Financeiros"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Endereço"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Comissão"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Estatísticas"
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
Attribute VB_Name = "TRPEmissores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_objUserControl As Object

Public gobjAnotacao As ClassAnotacoes
Public gobjTabEnd As New ClassTabEndereco

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iFrameAtual As Integer
Public iAlterado As Integer

'********** Janaina **********
Dim objGrid2 As AdmGrid
Dim objGridCategoria As AdmGrid
Dim iGrid_Categoria_Col As Integer
Dim iGrid_Valor_Col As Integer
'******************************

Dim objGridExcecoes As AdmGrid
Dim iGrid_Produto_Col As Integer
Dim iGrid_DescricaoProduto_Col As Integer
Dim iGrid_PercComissProd_Col As Integer

Dim sClienteAnt As String
Dim gobjUsuarioWeb As ClassTRPUsuarios

'Alteracao Daniel
Private WithEvents objEventoContaDespesa As AdmEvento
Attribute objEventoContaDespesa.VB_VarHelpID = -1

Private WithEvents objEventoTipoFornecedor As AdmEvento
Attribute objEventoTipoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoCondicaoPagto As AdmEvento
Attribute objEventoCondicaoPagto.VB_VarHelpID = -1
Private WithEvents objEventoPais As AdmEvento
Attribute objEventoPais.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoEmissor As AdmEvento
Attribute objEventoEmissor.VB_VarHelpID = -1
Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1
Private WithEvents objEventoContaFornConsig As AdmEvento
Attribute objEventoContaFornConsig.VB_VarHelpID = -1
Private WithEvents objEventoCidade As AdmEvento
Attribute objEventoCidade.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoBancos As AdmEvento
Attribute objEventoBancos.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_DadosFinanceiros = 2
Private Const TAB_Endereco = 3
Private Const TAB_Comissao = 4
Private Const TAB_Estatisticas = 5

Private Sub Ativo_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub BotaoWebAcesso_Click()
    gobjUsuarioWeb.iTipoUsuario = TRP_USUARIO_EMISSOR
    gobjUsuarioWeb.sNome = NomeReduzido.Text
    gobjUsuarioWeb.lCodigo = StrParaDbl(Codigo.Text)
    Call Chama_Tela_Modal("TRPUsuarios", gobjUsuarioWeb)
End Sub

Private Sub LabelNomeReduzido_Click()

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim colSelecao As New Collection

On Error GoTo Erro_LabelNomeReduzido_Click

    'Verifica se o fornecedor foi preenchido
    If Len(NomeReduzido.ClipText) <> 0 Then

        'Preenche o código de objFornecedor
        objFornecedor.sNomeReduzido = NomeReduzido.Text
        
    End If

    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

    Exit Sub

Erro_LabelNomeReduzido_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160575)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Gera código automático do proximo Fornecedor
    lErro = CF("TRPEmissor_Automatico", lCodigo)
    If lErro <> SUCESSO Then Error 55708

    'Exibe código automático na Tela
    Codigo.Text = CStr(lCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 55708
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160576)
    
    End Select

    Exit Sub

End Sub

Private Sub Agencia_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Banco_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Banco_GotFocus()

    Call MaskEdBox_TrataGotFocus(Banco, iAlterado)
    
End Sub

Private Sub Banco_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Banco_Validate

    'Verifica se foi preenchido o campo Banco
    If Len(Trim(Banco.Text)) = 0 Then Exit Sub

    'Critica se é do tipo positivo
    lErro = Valor_Positivo_Critica(Banco.Text)
    If lErro <> SUCESSO Then Error 43585

    Exit Sub

Erro_Banco_Validate:

    Cancel = True


    Select Case Err

        Case 43585
          
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160577)

    End Select

    Exit Sub

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

    'Verifica se foi preenchido o campo Codigo Fornecedor
    If Len(Trim(Codigo.Text)) = 0 Then Exit Sub

    'Critica se é do tipo positivo
    lErro = Valor_Positivo_Critica(Codigo.Text)
    If lErro <> SUCESSO Then Error 19301

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case Err

        Case 19301 'Tratado na rotina chamada
          
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160578)

    End Select

    Exit Sub

End Sub

Private Sub ComboCategoriaFornecedor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ComboCategoriaFornecedor_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ComboCategoriaFornecedor_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridCategoria)
End Sub

Private Sub ComboCategoriaFornecedor_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCategoria)
End Sub

Private Sub ComboCategoriaFornecedor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCategoria.objControle = ComboCategoriaFornecedor
    lErro = Grid_Campo_Libera_Foco(objGridCategoria)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ComboCategoriaFornecedorItem_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ComboCategoriaFornecedorItem_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ComboCategoriaFornecedorItem_GotFocus()

Dim lErro As Long

On Error GoTo Erro_ComboCategoriaFornecedorItem_GotFocus

    'Preenche com os ítens relacionados a Categoria correspondente
    Call Trata_ComboCategoriaFornecedorItem

    Call Grid_Campo_Recebe_Foco(objGridCategoria)

    Exit Sub

Erro_ComboCategoriaFornecedorItem_GotFocus:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160579)

    End Select

    Exit Sub

End Sub

Private Function Trata_ComboCategoriaFornecedorItem()

Dim lErro As Long
Dim objCategoriaFornecedor As New ClassCategoriaCliente
Dim iIndice As Integer, sValor As String

On Error GoTo Erro_Trata_ComboCategoriaFornecedorItem

    'Armazena o Item em sValor
    sValor = ComboCategoriaFornecedorItem.Text

    'Limpa a Combo
    ComboCategoriaFornecedorItem.Clear

    'Retorna o Item para a Combo
    ComboCategoriaFornecedorItem.Text = sValor

    'Se alguém estiver selecionado
    If Len(Trim(GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Categoria_Col))) > 0 Then

        'Preencher a Combo de Itens desta Categoria
        objCategoriaFornecedor.sCategoria = GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Categoria_Col)

        lErro = Carrega_ComboCategoriaFornecedorItem()
        If lErro <> SUCESSO Then gError 91169

    End If

    'Seleciona o Item na List da Combo
    For iIndice = 0 To ComboCategoriaFornecedorItem.ListCount - 1
        If ComboCategoriaFornecedorItem.List(iIndice) = GridCategoria.Text Then
            ComboCategoriaFornecedorItem.ListIndex = iIndice
            Exit For
        End If
    Next

    Trata_ComboCategoriaFornecedorItem = SUCESSO

    Exit Function

Erro_Trata_ComboCategoriaFornecedorItem:

    Trata_ComboCategoriaFornecedorItem = gErr

    Select Case gErr

        Case 91169

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160580)

    End Select

End Function

Private Function Carrega_ComboCategoriaFornecedorItem() As Long
'Carrega a Combo CategoriaFornecedorItem

Dim lErro As Long
Dim colItensCategoria As New Collection
Dim objCategoriaFornecedorItem As New ClassCategoriaFornItem

On Error GoTo Erro_Carrega_ComboCategoriaFornecedorItem

    'Guarda no obj a categoria que deverá ter os itens lidos
    objCategoriaFornecedorItem.sCategoria = GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Categoria_Col)
    
    'Lê a tabela CategoriaProdutoItem a partir da Categoria
    lErro = CF("CategoriaFornecedor_Le_Itens", objCategoriaFornecedorItem, colItensCategoria)
    If lErro <> SUCESSO And lErro <> 91180 Then gError 91170

    'Insere na combo CategoriaFornecedorItem
    For Each objCategoriaFornecedorItem In colItensCategoria

        'Insere na combo CategoriaFornecedor
        ComboCategoriaFornecedorItem.AddItem objCategoriaFornecedorItem.sItem

    Next

    Carrega_ComboCategoriaFornecedorItem = SUCESSO

    Exit Function

Erro_Carrega_ComboCategoriaFornecedorItem:

    Carrega_ComboCategoriaFornecedorItem = gErr

    Select Case gErr

        Case 91170

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160581)

    End Select

    Exit Function

End Function

Private Sub ComboCategoriaFornecedorItem_KeyPress(KeyAscii As Integer)
Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCategoria)
End Sub

Private Sub ComboCategoriaFornecedorItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCategoria.objControle = ComboCategoriaFornecedorItem
    lErro = Grid_Campo_Libera_Foco(objGridCategoria)
    If lErro <> SUCESSO Then Cancel = True


End Sub

Private Sub CondicaoPagto_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaCorrente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Filiais_Click()

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filiais_Click

    'Verifica se foi preenchido o Codigo
    If Len(Trim(Codigo.Text)) = 0 Then Error 12862
        
    'Preenche objFornecedor
    objFornecedor.lCodigo = StrParaLong(Codigo.Text)
    'Lê o Fornecedor
    lErro = CF("Fornecedor_Le", objFornecedor)
    If lErro <> SUCESSO And lErro <> 12729 Then Error 43452
    
    'Se não achou o Fornecedor
    If lErro <> SUCESSO Then
            
            'Envia aviso perguntando se deseja cadastrar novo Fornecedor
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FORNECEDOR")
    
            If vbMsgRes = vbYes Then
                lErro = Gravar_Registro()
                If lErro <> SUCESSO Then Error 43453
                'Chama a Tela de Filiais de Fornecedor
                objFilialFornecedor.lCodFornecedor = CLng(Codigo.Text)
                objFilialFornecedor.iCodFilial = 1 'p/começar exibindo a matriz
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If
    Else
        'Chama a Tela de Filiais de Fornecedor
        objFilialFornecedor.lCodFornecedor = CLng(Codigo.Text)
        objFilialFornecedor.iCodFilial = 1 'p/começar exibindo a matriz
        Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
    End If
    
    Exit Sub

Erro_Filiais_Click:

    Select Case Err

        Case 12862
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODFORNECEDOR_NAO_PREENCHIDO", Err)

        Case 43452, 43453
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160582)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colCodigoNome As New AdmCollCodigoNome
Dim objCodigoNome As AdmlCodigoNome
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome
Dim colCodigo As New Collection
Dim vCodigo As Variant
Dim iIndice As Integer
Dim sMascaraConta As String
Dim iTipoPadrao As Integer
Dim objTela As Object

On Error GoTo Erro_Form_Load

    'Implementado pois agora é possível ter constantes cutomizadas em função de tamanhos de campos do BD. AdmLib.ClassConsCust
    RazaoSocial.MaxLength = STRING_FORNECEDOR_RAZAO_SOC
    NomeReduzido.MaxLength = STRING_FORNECEDOR_NOME_REDUZIDO
    
    Set objGridCategoria = New AdmGrid
    Set objGridExcecoes = New AdmGrid
    
    sClienteAnt = ""
    Set gobjUsuarioWeb = New ClassTRPUsuarios

    'Inicializa o Grid de Categoria
    lErro = Inicializa_Grid_Categoria(objGridCategoria)
    If lErro <> SUCESSO Then gError 91171

    'Inicializa o Grid de Categoria
    lErro = Inicializa_Grid_Excecoes(objGridExcecoes)
    If lErro <> SUCESSO Then gError 91171

    'Carrega a combobox de Categoria Cliente
    lErro = Carrega_ComboCategoriaFornecedor()
    If lErro <> SUCESSO And lErro <> 91176 Then gError 91172
    
    iFrameAtual = 1

    Set objEventoTipoFornecedor = New AdmEvento
    Set objEventoCondicaoPagto = New AdmEvento
    Set objEventoFornecedor = New AdmEvento
    Set objEventoEmissor = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoBancos = New AdmEvento
    
    Set colCodigoDescricao = New AdmColCodigoNome

    'Lê cada código e descrição reduzida da tabela CondicoesPagto
    lErro = CF("CondicoesPagto_Le_Pagamento", colCodigoDescricao)
    If lErro <> SUCESSO Then gError 12719

    'Preenche a ComboBox CondicaoPagto com os objetos da coleção colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao

        CondicaoPagto.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        CondicaoPagto.ItemData(CondicaoPagto.NewIndex) = objCodigoDescricao.iCodigo

    Next

    Set objTela = Me
    lErro = gobjTabEnd.Inicializa(objTela, TabEnd(0))
    If lErro <> SUCESSO Then gError 12720

    'Verifica se o modulo de contabilidade esta ativo antes das inicializacoes
    If (gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO) Then
        
        'Conta de Fornecedores
        Set objEventoContaContabil = New AdmEvento
        
        'Inicializa propriedade Mask de ContaContabil
        lErro = MascaraConta(sMascaraConta)
        If lErro <> SUCESSO Then gError 12722
    
        ContaContabil.Mask = sMascaraConta
        
        'Conta de Fornecedores em consignação
        Set objEventoContaFornConsig = New AdmEvento
        
        'Inicializa propriedade Mask de ContaContabil
        lErro = MascaraConta(sMascaraConta)
        If lErro <> SUCESSO Then gError 12722
    
        ContaFornConsig.Mask = sMascaraConta
    
        'Alteracao Daniel
        'Conta de Estoque/Despesa
        Set objEventoContaDespesa = New AdmEvento
        
        'Inicializa propriedade Mask de ContaDespesa
        lErro = MascaraConta(sMascaraConta)
        If lErro <> SUCESSO Then gError 106003
    
        ContaDespesa.Mask = sMascaraConta
        'Fim da Alteracao Daniel
    
    Else

        ContaFornConsig.Enabled = False
        ContaFornConsig.Enabled = False
        
        'Alteracao Daniel
        ContaDespesa.Enabled = False
        LabelContaDespesa.Enabled = False
        'Fim da Alteracao Daniel
        
    End If
    
    'Alteracao Daniel
    'Carrega a List da Combo de Tipos de Cobrança
    lErro = Carrega_TipoCobranca()
    If lErro <> SUCESSO Then gError 106006
        
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 12718, 12719, 12720, 12721, 12722, 61394, 91171, 91171, 106003, 106006

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160583)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name
        
            Case GridCategoria.Name
            
                Select Case objGridInt.objGrid.Col
        
                    Case GRID_CATEGORIA_COL
        
                        'Chama SaidaCelula de Categoria
                        lErro = Saida_Celula_Categoria(objGridInt)
                        If lErro <> SUCESSO Then gError 91173
        
                    Case GRID_VALOR_COL
        
                        'Chama SaidaCelula de Valor
                        lErro = Saida_Celula_Valor(objGridInt)
                        If lErro <> SUCESSO Then gError 91174
        
                End Select
                
            Case GridExcecoes.Name
                
                'Verifica qual a coluna atual do Grid
                Select Case objGridInt.objGrid.Col
            
                    'Se for a de Produto
                    Case iGrid_Produto_Col
                        lErro = Saida_Celula_Produto(objGridInt)
                        If lErro <> SUCESSO Then gError 91173
            
                    'Se for a de Unidade de Medida
                    Case iGrid_PercComissProd_Col
                        lErro = Saida_Celula_PercComissProd(objGridInt)
                        If lErro <> SUCESSO Then gError 91174
            
                End Select
            
        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 91175

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 91173, 91174

        Case 91175
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160585)

    End Select

    Exit Function

End Function

Private Function Carrega_ComboCategoriaFornecedor() As Long
'Carrega as Categorias na Combobox

Dim lErro As Long
Dim colCategorias As New Collection
Dim objCategoriaFornecedor As New ClassCategoriaFornecedor

On Error GoTo Erro_Carrega_ComboCategoriaFornecedor
  
    'Lê o código e a descrição de todas as categorias
    lErro = CF("CategoriaFornecedor_Le_Todos", colCategorias)
    If lErro <> SUCESSO And lErro <> 64686 Then gError 91176

    For Each objCategoriaFornecedor In colCategorias

        'Insere na combo CategoriaCliente
        ComboCategoriaFornecedor.AddItem objCategoriaFornecedor.sCategoria

    Next

    Carrega_ComboCategoriaFornecedor = SUCESSO

    Exit Function

Erro_Carrega_ComboCategoriaFornecedor:

    Carrega_ComboCategoriaFornecedor = gErr

    Select Case gErr

        Case 91176

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160586)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Categoria(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Categoria do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim colCategorias As New Collection
Dim objCategoriaFornecedor As New ClassCategoriaFornecedor

On Error GoTo Erro_Saida_Celula_Categoria

    Set objGridInt.objControle = ComboCategoriaFornecedor

    'Testa se .Text da combo é igual ao TextMatrix "embaixo"
    iIndice = GridCategoria.Row
    
    If UCase(GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL)) <> UCase(ComboCategoriaFornecedor.Text) Then

        If Len(Trim(ComboCategoriaFornecedor)) > 0 Then

            'Preenche o objeto com a Categoria
            objCategoriaFornecedor.sCategoria = ComboCategoriaFornecedor.Text

            'Lê Categoria De Cliente no BD
            lErro = CF("CategoriaFornecedor_Le", objCategoriaFornecedor)
            If lErro <> SUCESSO And lErro <> 90592 Then gError 91182

            'Categoria não está cadastrada
            If lErro = 90592 Then gError 91183

            'Verifica se já existe a categoria no Grid
            If GridCategoria.Row > 0 Then

                iIndice1 = GridCategoria.Row

                For iIndice = 1 To objGridCategoria.iLinhasExistentes

                    If iIndice <> iIndice1 Then

                        'Se já existe a categoria, erro
                        If UCase(GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL)) = UCase(ComboCategoriaFornecedor.Text) Then gError 91184

                    End If

                Next

            End If

            If GridCategoria.Row > 0 Then

                iIndice1 = GridCategoria.Row

                'Limpa o valor (ítem) da categoria
                GridCategoria.TextMatrix(iIndice1, GRID_VALOR_COL) = ""

            End If

        Else

             If GridCategoria.Row > 0 Then

                iIndice1 = GridCategoria.Row

                'Categoria vazia, limpa o Item correspondente
                GridCategoria.TextMatrix(iIndice1, GRID_VALOR_COL) = ""

            End If

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 91185

    If (GridCategoria.Row - GridCategoria.FixedRows = objGridInt.iLinhasExistentes) And Len(Trim(ComboCategoriaFornecedor.Text)) > 0 Then
        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If

    Saida_Celula_Categoria = SUCESSO

    Exit Function

Erro_Saida_Celula_Categoria:

    Saida_Celula_Categoria = gErr

    Select Case gErr

        Case 91182
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 91183 'Categoria não está cadastrada

            'Perguntar se deseja criar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DESEJA_CRIAR_CATEGORIAFORNECEDOR")

            If vbMsgRes = vbYes Then

                'Chama a Tela "CategoriaFornecedor"
                Call Chama_Tela("CategoriaFornec", objCategoriaFornecedor)
                
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            Else

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 91184
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_JA_SELECIONADA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 91185
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160587)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Valor(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Item do grid que está deixando de ser a corrente

'******** Janaina ********

Dim lErro As Long
Dim objCategoriaFornecedor As New ClassCategoriaFornecedor
Dim objCategoriaFornecedorItem As New ClassCategoriaFornItem
Dim colItens As New Collection
Dim colItensCategoria As New Collection

Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridCategoria.objControle = ComboCategoriaFornecedorItem

    If Len(Trim(ComboCategoriaFornecedorItem)) > 0 Then

        'se o campo de categoria estiver vazio ==> erro
        If Len(GridCategoria.TextMatrix(GridCategoria.Row, GRID_CATEGORIA_COL)) = 0 Then gError 91186

        'Tenta selecionar na combo
        lErro = Combo_Item_Igual_CI(ComboCategoriaFornecedorItem)
        If lErro <> SUCESSO Then

            'Preenche o objeto com a Categoria
            objCategoriaFornecedorItem.sCategoria = GridCategoria.TextMatrix(GridCategoria.Row, GRID_CATEGORIA_COL)
            objCategoriaFornecedorItem.sItem = ComboCategoriaFornecedorItem.Text

            'Alterado por Luiz Nogueira em 03/03/04
            'Lê Item De Categoria De Fornecedor no BD
            lErro = CF("CategoriaFornecedorItem_Le", objCategoriaFornecedorItem, colItensCategoria)
            If lErro <> SUCESSO And lErro <> 127089 Then Error 91187

            'Item da Categoria não está cadastrado
            If lErro = 127089 Then gError 91188

        End If

        If GridCategoria.Row - GridCategoria.FixedRows = objGridInt.iLinhasExistentes Then

            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 91189

    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = gErr

    Select Case gErr

        Case 91187
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 91188 'Item da Categoria não está cadastrado

            'Perguntar se deseja criar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DESEJA_CRIAR_CATEGORIAFORNECEDORITEM", objCategoriaFornecedorItem.sItem)

            If vbMsgRes = vbYes Then

                'Preenche o objeto com a Categoria
                objCategoriaFornecedor.sCategoria = GridCategoria.TextMatrix(GridCategoria.Row, GRID_CATEGORIA_COL)

                'Chama a Tela "CategoriaFornecedor"
                Call Chama_Tela("CategoriaFornecedor", objCategoriaFornecedor, objCategoriaFornecedorItem.sItem)
                
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            Else

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 91189
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 91186
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_CATEGORIA_NAO_PREENCHIDA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160588)

    End Select

    Exit Function

End Function

Public Sub GridCategoria_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridCategoria, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCategoria, iAlterado)
    End If

End Sub

Public Sub GridCategoria_EnterCell()
    Call Grid_Entrada_Celula(objGridCategoria, iAlterado)
End Sub

Public Sub GridCategoria_GotFocus()
    Call Grid_Recebe_Foco(objGridCategoria)
End Sub

Public Sub GridCategoria_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridCategoria)
End Sub

Public Sub GridCategoria_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCategoria, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCategoria, iAlterado)
    End If

End Sub

Public Sub GridCategoria_LeaveCell()
    Call Saida_Celula(objGridCategoria)
End Sub

Public Sub GridCategoria_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridCategoria)
End Sub

Public Sub GridCategoria_RowColChange()
    Call Grid_RowColChange(objGridCategoria)
End Sub

Public Sub GridCategoria_Scroll()
    Call Grid_Scroll(objGridCategoria)
End Sub

Private Function Inicializa_Grid_Categoria(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Categoria")
    objGridInt.colColuna.Add ("Item")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ComboCategoriaFornecedor.Name)
    objGridInt.colCampo.Add (ComboCategoriaFornecedorItem.Name)

    'Colunas do Grid
    iGrid_Categoria_Col = 1
    iGrid_Valor_Col = 2

    'Grid do GridInterno
    objGridInt.objGrid = GridCategoria

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 21

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridCategoria.ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Categoria = SUCESSO

    Exit Function

End Function

Private Sub LabelCodigoEmi_Click()

Dim lErro As Long
Dim objEmi As New ClassCliEmissoresTRP
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigoEmi_Click

    'Verifica se o fornecedor foi preenchido
    If Len(Codigo.ClipText) <> 0 Then

        'Preenche o código de objFornecedor
        objEmi.lCodigo = StrParaLong(Codigo.Text)
        
    End If

    Call Chama_Tela("TRPEmissoresLista", colSelecao, objEmi, objEventoEmissor)

    Exit Sub

Erro_LabelCodigoEmi_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160589)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o fornecedor foi preenchido
    If Len(CodigoForn.Caption) <> 0 Then

        'Preenche o código de objFornecedor
        objFornecedor.lCodigo = StrParaLong(CodigoForn.Caption)
        
    End If

    colSelecao.Add TRP_TIPO_CLIENTE_EMISSOR

    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor, "Tipo = ?")

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160589)

    End Select

    Exit Sub

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objFornecedor As ClassFornecedor
Dim objFornecedorEstatistica As New ClassFilialFornecedorEst

On Error GoTo Erro_objEventoFornecedor_evSelecao

    Set objFornecedor = obj1
    
    lErro = CF("Fornecedor_Le_Estendida", objFornecedor, objFornecedorEstatistica)
    If lErro <> SUCESSO And lErro <> 52701 Then gError 82678

    'Verifica se Fornecedor não está cadastrado
    If lErro <> SUCESSO Then gError 82679

    lErro = Exibe_Dados_Fornecedor(objFornecedor, objFornecedorEstatistica)
    If lErro <> SUCESSO Then gError 82680

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Me.Show
    
    Exit Sub

Erro_objEventoFornecedor_evSelecao:

    Select Case gErr

    Case 82678, 82680

    Case 82679
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INEXISTENTE", gErr)

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160590)

    End Select

    Exit Sub

    

End Sub

Private Sub objEventoEmissor_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objEmi As ClassCliEmissoresTRP
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_objEventoEmissor_evSelecao

    Set objEmi = obj1
    
    objFornecedor.lCodigo = objEmi.lFornEmissor
    
    lErro = Trata_Parametros(objFornecedor, objEmi)
    If lErro <> SUCESSO Then gError 82680

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    Me.Show
    
    Exit Sub

Erro_objEventoEmissor_evSelecao:

    Select Case gErr

    Case 82678, 82680

    Case 82679
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INEXISTENTE", gErr)

    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160590)

    End Select

    Exit Sub

End Sub

Private Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Observacao2_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub RazaoSocial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CondicaoPagto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CondicaoPagto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim iCodigo As Integer

On Error GoTo Erro_CondicaoPagto_Validate

    'Verifica se foi preenchida a ComboBox CondicaoPagto
    If Len(Trim(CondicaoPagto.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox CondicaoPagto
    If CondicaoPagto.Text = CondicaoPagto.List(CondicaoPagto.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(CondicaoPagto, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 33548

    'Não existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objCondicaoPagto.iCodigo = iCodigo

        'Tenta ler CondicaoPagto com esse código no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then Error 33549
        
        If lErro <> SUCESSO Then Error 33550 'Não encontrou CondicaoPagto no BD

        'Encontrou CondicaoPagto no BD e não é de Pagamento
        If objCondicaoPagto.iEmPagamento = 0 Then Error 33551

        'Coloca no Text da Combo
        CondicaoPagto.Text = CStr(objCondicaoPagto.iCodigo) & SEPARADOR & objCondicaoPagto.sDescricao

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 33552

    Exit Sub

Erro_CondicaoPagto_Validate:

    Cancel = True


    Select Case Err

        Case 33551
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_RECEBIMENTO", Err, iCodigo)
    
        Case 33548, 33549
    
        Case 33550  'Não encontrou CondicaoPagto no BD
    
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAO_PAGAMENTO")
    
            If vbMsgRes = vbYes Then
                'Chama a tela de CondicaoPagto
                Call Chama_Tela("CondicoesPagto", objCondicaoPagto)
    
            Else
                'Segura o foco
    
            End If
    
        Case 33551
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_RECEBIMENTO", Err, iCodigo)
    
        Case 33552
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", Err, CondicaoPagto.Text)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160594)

    End Select

    Exit Sub

End Sub

Private Sub Desconto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Desconto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sDesconto As String

On Error GoTo Erro_Desconto_Validate

    sDesconto = Desconto.Text

    'Verifica se foi preenchido o Desconto
    If Len(Trim(Desconto.Text)) = 0 Then Exit Sub

    'Critica se é porcentagem
    lErro = Porcentagem_Critica(Desconto.Text)
    If lErro <> SUCESSO Then Error 12776

    Desconto.Text = Format(sDesconto, "Fixed")

    Exit Sub

Erro_Desconto_Validate:

    Cancel = True


    Select Case Err

        Case 12776
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160595)
    
        End Select

    Exit Sub

End Sub

Private Sub ContaContabil_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaFornConsig_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaContabil_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaMascarada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ContaContabil_Validate

    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica_Modulo", ContaContabil.Text, ContaContabil.ClipText, objPlanoConta, MODULO_CONTASAPAGAR)
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then Error 90540
    
    If lErro = SUCESSO Then
    
        sContaFormatada = objPlanoConta.sConta
        
        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)
        
        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 39798
        
        ContaContabil.PromptInclude = False
        ContaContabil.Text = sContaMascarada
        ContaContabil.PromptInclude = True
    
    'se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then
    
        lErro = CF("Conta_Critica", ContaContabil.Text, sContaFormatada, objPlanoConta, MODULO_CONTASAPAGAR)
        If lErro <> SUCESSO And lErro <> 5700 Then Error 12777
        
        If lErro = 5700 Then Error 12778

    End If
    
    Exit Sub

Erro_ContaContabil_Validate:

    Cancel = True

    Select Case Err

    Case 12777, 39797
        
    Case 12778
        lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", Err, ContaContabil.Text)

    Case 39798
        lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160596)

    End Select

    Exit Sub

End Sub

Private Sub ContaFornConsig_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaMascarada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ContaFornConsig_Validate

    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica_Modulo", ContaFornConsig.Text, ContaFornConsig.ClipText, objPlanoConta, MODULO_CONTASAPAGAR)
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 39797
    
    If lErro = SUCESSO Then
    
        sContaFormatada = objPlanoConta.sConta
        
        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)
        
        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then gError 90541
        
        ContaFornConsig.PromptInclude = False
        ContaFornConsig.Text = sContaMascarada
        ContaFornConsig.PromptInclude = True
    
    'se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then
    
        lErro = CF("Conta_Critica", ContaFornConsig.Text, sContaFormatada, objPlanoConta, MODULO_CONTASAPAGAR)
        If lErro <> SUCESSO And lErro <> 5700 Then gError 90542
        
        If lErro = 5700 Then gError 90543

    End If
    
    Exit Sub

Erro_ContaFornConsig_Validate:

    Cancel = True


    Select Case gErr

    Case 90542, 90540
        
    Case 90543
        Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", gErr, ContaFornConsig.Text)

    Case 90541
        Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160597)

    End Select

    Exit Sub

End Sub

Private Sub CGC_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CGC_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CGC, iAlterado)
    
End Sub

Private Sub CGC_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CGC_Validate

    If Len(Trim(CGC.Text)) = 0 Then Exit Sub

    Select Case Len(Trim(CGC.Text))

        Case STRING_CPF 'CPF
    
            lErro = Cpf_Critica(CGC.Text)
            If lErro <> SUCESSO Then Error 12780
    
            CGC.Format = "000\.000\.000-00; ; ; "
            CGC.Text = CGC.Text
    
        Case STRING_CGC  'CGC
    
            lErro = Cgc_Critica(CGC.Text)
            If lErro <> SUCESSO Then Error 12781
    
            CGC.Format = "00\.000\.000\/0000-00; ; ; "
            CGC.Text = CGC.Text
    
    Case Else

        Error 12782

    End Select

    Exit Sub

Erro_CGC_Validate:

    Cancel = True



    Select Case Err

        Case 12780, 12781

        Case 12782
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC_CPF", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160598)

    End Select

    Exit Sub

End Sub

Private Sub InscricaoEstadual_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub InscricaoMunicipal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava o Fornecedor
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 47890

    'Limpa a Tela
    lErro = Limpa_Tela_Fornecedores
    If lErro <> SUCESSO Then gError 78017

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 47890, 78017

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160601)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 12858

    'Limpa a Tela
    lErro = Limpa_Tela_Fornecedores
    If lErro <> SUCESSO Then gError 78018

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 12858, 78018

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160602)

    End Select

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim colCodNomeFiliais As New AdmColCodigoNome
Dim objFornecedor As New ClassFornecedor
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim lCodigo As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    Parent.MousePointer = vbHourglass
    
    'Verifica se o código foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 12822

    objFornecedor.lCodigo = CLng(Codigo.Text)

    'Lê os dados do Fornecedor a ser excluido
    lErro = CF("Fornecedor_Le", objFornecedor)
    If lErro <> SUCESSO And lErro <> 12729 Then gError 12823

    'Verifica se Fornecedor não está cadastrado
    If lErro <> SUCESSO Then gError 12824

    'Envia aviso perguntando se realmente deseja excluir Fornecedor e suas filiais
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_FORNECEDOR", objFornecedor.lCodigo)

    If vbMsgRes = vbYes Then

        'Exclui Fornecedor
        lErro = CF("Fornecedor_Exclui", objFornecedor)
        If lErro <> SUCESSO Then gError 12826

        'Limpa a Tela
        lErro = Limpa_Tela_Fornecedores
        If lErro <> SUCESSO Then gError 78019
        
        iAlterado = 0

    End If

    Parent.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    Parent.MousePointer = vbDefault
    
    Select Case gErr

        Case 12822, 78019
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODFORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 12823, 12826, 91191

        Case 12824
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160603)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

 Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
   
End Sub

Function Limpa_Tela_Fornecedores() As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Tela_Fornecedores

    Call Limpa_Tela(Me)

    CondicaoPagto.Text = ""

    'Inicializa os campos cliente existentes nos frames
    For iIndice = 2 To 5
        Fornecedor(iIndice).Caption = ""
    Next
    
    'cyntia em 14/06/02
    'Limpa os labels
    SaldoTitulos.Caption = Format(0, "standard")
    NumeroCompras.Caption = Format(0, "standard")
    MediaCompra.Caption = Format(0, "standard")
    ValorAcumuladoCompras.Caption = Format(0, "standard")
    DataPrimeiraCompra.Caption = ""
    DataUltimaCompra.Caption = ""
    ValorAcumuladoDevolucoes.Caption = Format(0, "standard")
    NumTotalDevolucoes.Caption = 0
    DataUltDevolucao.Caption = ""
    PedidosEmAberto.Caption = 0
    MediaAtraso.Caption = 0
    MaiorAtraso.Caption = 0
    
    Ativo.Value = MARCADO
    
    Call gobjTabEnd.Limpa_Tela
    
    '******** Janaina ********
    Call Grid_Limpa(objGridCategoria)
    '*************************
    
    sClienteAnt = ""
    Set gobjUsuarioWeb = New ClassTRPUsuarios
    Superior.Clear
    Call Grid_Limpa(objGridExcecoes)
    CodigoForn.Caption = ""
    
    Set gobjAnotacao = Nothing
        
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Limpa_Tela_Fornecedores = SUCESSO
    
    Exit Function
    
Erro_Limpa_Tela_Fornecedores:

    Limpa_Tela_Fornecedores = gErr
    
    Select Case gErr
        
        Case 78015
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160604)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objFornecedor As New ClassFornecedor
Dim objEmi As New ClassCliEmissoresTRP
Dim objCategoriaFornecedor As ClassCategoriaFornecedor
Dim objFornecedorList As New AdmlCodigoNome
Dim objEndereco As New ClassEndereco
Dim colItensCategoria As New Collection
Dim lCodigo As Long
Dim iAchou As Integer
Dim colEnderecos As New Collection

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se foi preenchido o Codigo
    If Len(Trim(Codigo.Text)) = 0 Then gError 12792

    'Verifica se foi preenchida a Razao Social
    If Len(Trim(RazaoSocial.Text)) = 0 Then gError 12793

    'Verifica se foi preenchido o Nome Reduzido
    If Len(Trim(NomeReduzido.Text)) = 0 Then gError 12794
    
    'Verifica se foi preenchido o Cliente
    If Len(Trim(Cliente.Text)) = 0 Then gError 200350
    
    If CondicaoPagto.ListIndex = -1 Then gError 200844
    
    'Lê os dados do Endereço
    lErro = gobjTabEnd.Move_Endereco_Memoria(colEnderecos)
    If lErro <> SUCESSO Then gError 12795
    
    Set objEndereco = colEnderecos.Item(1)

    'Verifica se para cada Categoria existente no grid de Categorias existe um Valor correspondente
    For iIndice = 1 To objGridCategoria.iLinhasExistentes
        'Verifica se a Categoria foi preenchida
         If Len(Trim(GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL))) <> 0 And Len(Trim(GridCategoria.TextMatrix(iIndice, GRID_VALOR_COL))) = 0 Then gError 91190
    Next
    
    'Lê os dados da tela do Fornecedor
    lErro = Le_Dados_Fornecedor(objFornecedor, objEmi)
    If lErro <> SUCESSO Then gError 12798
    
    lErro = CF("InscricaoEstadual_Critica", objFornecedor.sInscricaoEstadual, objEndereco.sSiglaEstado)
    If lErro <> SUCESSO Then gError 189577
    
    lErro = Trata_Alteracao(objFornecedor, objFornecedor.lCodigo)
    If lErro <> SUCESSO Then gError 80453

    'Se o CGC estiver Preenchido
    If Len(Trim(objFornecedor.sCgc)) > 0 Then
        'Verifica se tem outro fornecedor com o mesmo CGC e dá aviso
        lErro = CF("FilialFornecedor_Testa_CGC", objFornecedor.lCodigo, 0, objFornecedor.sCgc)
        If lErro <> SUCESSO Then gError 80217
    End If
    
    lErro = CF("TRPEmissores_Grava", objFornecedor, objEndereco, objEmi)
    If lErro <> SUCESSO Then gError 12799
    
    'gravar anotacao, se houver
    If Not (gobjAnotacao Is Nothing) Then
    
        If Len(Trim(gobjAnotacao.sTextoCompleto)) <> 0 Or Len(Trim(gobjAnotacao.sTitulo)) <> 0 Then
        
            gobjAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_FORNECEDOR
            gobjAnotacao.sID = CStr(objFornecedor.lCodigo)
            gobjAnotacao.dtDataAlteracao = gdtDataHoje
            
            lErro = CF("Anotacoes_Grava", gobjAnotacao)
            If lErro <> SUCESSO Then gError 128149
            
        End If
        
    End If
    
    iAchou = 0

    iAlterado = 0

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = gErr

    Select Case gErr

        Case 12792
            Call Rotina_Erro(vbOKOnly, "ERRO_CODFORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 12793
            Call Rotina_Erro(vbOKOnly, "ERRO_RAZ_SOC_NAO_PREENCHIDA", gErr)

        Case 12794
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_PREENCHIDO", gErr)

        Case 12795, 12798, 80217, 80453, 91191, 12799, 128149, 189577
        
        Case 62672
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TIPO_FORNECEDOR_NAO_PREENCHIDO", gErr)
            
        Case 91190
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAFORNECEDORITEM_ITEM_NAO_SELECIONADO1", gErr, iIndice)
        
        Case 200350
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 200844
            Call Rotina_Erro(vbOKOnly, "ERRO_PREENCH_COND_PAGTO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160605)

    End Select

    Exit Function

End Function

Private Function Le_Dados_Fornecedor(ByVal objFornecedor As ClassFornecedor, ByVal objEmi As ClassCliEmissoresTRP) As Long
'Lê os dados que estão na tela de Fornecedores e coloca em objFornecedor

Dim lErro As Long
Dim iContaPreenchida As Integer
Dim objFilialFornCategoria As ClassFilialFornCategoria
Dim iIndice As Integer
Dim sConta As String
Dim sContaFornConsig As String
Dim objCliente As New ClassCliente
Dim objEmiExc As ClassCliEmissoresExcTRP
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Le_Dados_Fornecedor

    'IDENTIFICACAO :
    
    objEmi.lCodigo = StrParaLong(Codigo.Text)
    If Len(Trim(PercComi.Text)) > 0 Then objEmi.dPercComissao = CDbl(PercComi.Text) / 100

    objCliente.sNomeReduzido = Cliente.Text

    'Lê o Cliente através do Nome Reduzido
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 12800

    objEmi.lCliente = objCliente.lCodigo
    
    objEmi.lSuperior = LCodigo_Extrai(Superior.Text)
    
    Set objEmi.objUsuarioWeb = gobjUsuarioWeb

    objFornecedor.lCodigo = StrParaLong(CodigoForn.Caption)
    
    objFornecedor.sRazaoSocial = RazaoSocial.Text
    objFornecedor.sNomeReduzido = NomeReduzido.Text
    objFornecedor.iTipo = TRP_TIPO_CLIENTE_FORNECEDORES
    objFornecedor.sObservacao = Observacao.Text
    objFornecedor.iAtivo = Ativo.Value

'********** Janaina ***************

    'Preenche uma coleção com todas as linhas "existentes" do grid de categorias
    For iIndice = 1 To objGridCategoria.iLinhasExistentes

        'Verifica se a Categoria foi preenchida
        If Len(Trim(GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL))) <> 0 Then

            Set objFilialFornCategoria = New ClassFilialFornCategoria
           
            objFilialFornCategoria.lFornecedor = objFornecedor.lCodigo
            objFilialFornCategoria.sCategoria = GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL)
            objFilialFornCategoria.sItem = GridCategoria.TextMatrix(iIndice, GRID_VALOR_COL)

            objFornecedor.colCategoriaItem.Add objFilialFornCategoria

        End If

    Next
    
'****************************
    'DADOS FINANCEIROS :

    If Len(Trim(Desconto.Text)) > 0 Then objFornecedor.dDesconto = CDbl(Desconto.Text) / 100

    If Len(Trim(CondicaoPagto.Text)) > 0 Then objFornecedor.iCondicaoPagto = Codigo_Extrai(CondicaoPagto.Text)

    lErro = CF("Conta_Formata", ContaContabil.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then gError 12800
    If iContaPreenchida = CONTA_VAZIA Then
        objFornecedor.sContaContabil = ""
    Else
        objFornecedor.sContaContabil = sConta
    End If
    
    lErro = CF("Conta_Formata", ContaFornConsig.Text, sContaFornConsig, iContaPreenchida)
    If lErro <> SUCESSO Then gError 90547
    
    If iContaPreenchida = CONTA_VAZIA Then
        objFornecedor.sContaFornConsig = ""
    Else
        objFornecedor.sContaFornConsig = sContaFornConsig
    End If
    
    objFornecedor.sContaCorrente = ContaCorrente.Text
    objFornecedor.sAgencia = Agencia.Text
    If Len(Trim(Banco.Text)) > 0 Then
       If CInt(Banco.Text) < 0 Then gError 12790
        objFornecedor.iBanco = CInt(Banco.Text)
    End If
    
    'INSCRICOES :

    objFornecedor.sCgc = CGC.Text
    objFornecedor.sInscricaoEstadual = InscricaoEstadual.Text
    objFornecedor.sInscricaoMunicipal = InscricaoMunicipal.Text
    objFornecedor.sObservacao2 = Observacao2.Text
    
    'Alteracao Daniel
    objFornecedor.sInscricaoINSS = InscricaoINSS.Text
    objFornecedor.iGeraCredICMS = GeraICMS.Value
    objFornecedor.iTipoFrete = TipoFrete.ListIndex
    objFornecedor.iTipoCobranca = Codigo_Extrai(TipoCobranca.Text)

    'ESTATISTICAS :

    objFornecedor.dSaldoTitulos = CDbl(SaldoTitulos.Caption)
    objFornecedor.lNumeroCompras = CLng(NumeroCompras.Caption)
    objFornecedor.dMediaCompra = CDbl(MediaCompra.Caption)
    objFornecedor.dValorAcumuladoCompras = CDbl(ValorAcumuladoCompras.Caption)
    If Trim(DataPrimeiraCompra.Caption) = "" Then
        objFornecedor.dtDataPrimeiraCompra = DATA_NULA
    Else
        objFornecedor.dtDataPrimeiraCompra = StrParaDate(DataPrimeiraCompra.Caption)
    End If
    If Trim(DataUltimaCompra.Caption) = "" Then
        objFornecedor.dtDataUltimaCompra = DATA_NULA
    Else
        objFornecedor.dtDataUltimaCompra = StrParaDate(DataUltimaCompra.Caption)
    End If
    objFornecedor.sAgencia = Agencia.Text
    objFornecedor.lMediaAtraso = CLng(MediaAtraso.Caption)
    objFornecedor.lMaiorAtraso = CLng(MaiorAtraso.Caption)
    objFornecedor.dValorAcumuladoDevolucoes = CDbl(ValorAcumuladoDevolucoes.Caption)
    objFornecedor.lNumTotalDevolucoes = CLng(NumTotalDevolucoes.Caption)
    If Trim(DataUltDevolucao.Caption) = "" Then
        objFornecedor.dtDataUltDevolucao = DATA_NULA
    Else
        objFornecedor.dtDataUltDevolucao = StrParaDate(DataUltDevolucao.Caption)
    End If
    
    'Alteracao Daniel
    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaDespesa.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then gError 12966

    If iContaPreenchida = CONTA_VAZIA Then
        objFornecedor.sContaDespesa = ""
    Else
        objFornecedor.sContaDespesa = sConta
    End If
    'Fim da Alteracao Daniel
    
    For iIndice = 1 To objGridExcecoes.iLinhasExistentes
    
        lErro = CF("Produto_Formata", GridExcecoes.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 195801
        
        If iProdutoPreenchido = PRODUTO_VAZIO Then gError 195807
        
        Set objEmiExc = New ClassCliEmissoresExcTRP
        
        objEmiExc.sProduto = sProdutoFormatado
        
        If Len(Trim(GridExcecoes.TextMatrix(iIndice, iGrid_PercComissProd_Col))) = 0 Then gError 195808
        
        objEmiExc.sProduto = sProdutoFormatado
        objEmiExc.dPercComissao = PercentParaDbl(GridExcecoes.TextMatrix(iIndice, iGrid_PercComissProd_Col))
    
        objEmi.colCliEmissoresExcTRP.Add objEmiExc
        
    Next

    Le_Dados_Fornecedor = SUCESSO

    Exit Function

Erro_Le_Dados_Fornecedor:

    Le_Dados_Fornecedor = gErr

    Select Case gErr

        Case 12800, 19301, 90547
        
        Case 12790
            Call Rotina_Erro(vbOKOnly, "ERRO_BANCO_NAO_POSITIVO", Err, Banco.Text)
       
        Case 195801
        
        Case 195807
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO_GRID", gErr, iIndice)
        
        Case 195808
            Call Rotina_Erro(vbOKOnly, "ERRO_PERCENTUAL_NAO_PREENCHIDO_GRID", gErr, iIndice)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160607)

    End Select

    Exit Function

End Function

Function Exibe_Dados_TipoFornecedor(objTipoFornecedor As ClassTipoFornecedor) As Long
'Exibe dados do Tipo de Fornecedor na tela

Dim lErro As Long
Dim sContaEnxuta As String

On Error GoTo Erro_Exibe_Dados_TipoFornecedor

    If objTipoFornecedor.iCondicaoPagto = 0 Then
        CondicaoPagto.Text = ""
    Else
        CondicaoPagto.Text = CStr(objTipoFornecedor.iCondicaoPagto)
        Call CondicaoPagto_Validate(bSGECancelDummy)
    End If

    If objTipoFornecedor.dDesconto = 0# Then
        Desconto.Text = ""
    Else
        Desconto.Text = CStr(100 * objTipoFornecedor.dDesconto)
    End If

    Observacao.Text = objTipoFornecedor.sObservacao

    ContaContabil.PromptInclude = False

    If objTipoFornecedor.sContaDespesa = "" Then
        ContaContabil.Text = ""
    Else
        lErro = Mascara_RetornaContaEnxuta(objTipoFornecedor.sContaDespesa, sContaEnxuta)
        If lErro <> SUCESSO Then Error 12766

        ContaContabil.Text = sContaEnxuta
    End If

    ContaContabil.PromptInclude = True

    Exibe_Dados_TipoFornecedor = SUCESSO

Exit Function

Erro_Exibe_Dados_TipoFornecedor:

    Exibe_Dados_TipoFornecedor = Err

    Select Case Err

        Case 12766
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objTipoFornecedor.sContaDespesa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160608)

    End Select

    Exit Function

End Function

Function Trata_Parametros(Optional obj1 As Object, Optional obj2 As Object) As Long

Dim lErro As Long
Dim lCodigo As Long
Dim objFornecedorEstatistica As New ClassFilialFornecedorEst
Dim objFornecedor As ClassFornecedor
Dim objEmi As ClassCliEmissoresTRP
Dim sClasse As String

On Error GoTo Erro_Trata_Parametros

    If Not (obj2 Is Nothing) Then
        sClasse = TypeName(obj2)
        Select Case sClasse
            Case "ClassFornecedor"
                Set objFornecedor = obj2
            Case "ClassCliEmissoresTRP"
                Set objEmi = obj2
        End Select
    End If
    
    If Not (obj1 Is Nothing) Then
        sClasse = TypeName(obj1)
        Select Case sClasse
            Case "ClassFornecedor"
                Set objFornecedor = obj1
            Case "ClassCliEmissoresTRP"
                Set objEmi = obj1
        End Select
    End If

    If Not (objEmi Is Nothing) Then
        Set objFornecedor = New ClassFornecedor
        
        lErro = CF("TRPEmissores_Le", objEmi)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 12723
        
        objFornecedor.lCodigo = objEmi.lFornEmissor
    End If
    
    If Not (objFornecedor Is Nothing) Then

        lErro = CF("Fornecedor_Le_Estendida", objFornecedor, objFornecedorEstatistica)
        If lErro <> SUCESSO And lErro <> 52701 Then gError 12723

        If lErro = SUCESSO Then
            lErro = Exibe_Dados_Fornecedor(objFornecedor, objFornecedorEstatistica)
            If lErro <> SUCESSO Then gError 12724
        Else
            lErro = Limpa_Tela_Fornecedores()
            If lErro <> SUCESSO Then gError 78016
            If objFornecedor.lCodigo <> 0 Then
                Codigo.Text = objFornecedor.lCodigo
            End If
            If objFornecedor.sNomeReduzido <> "" Then NomeReduzido.Text = objFornecedor.sNomeReduzido
        End If
    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 12723, 12724, 12745, 78016

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160609)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Sub NomeReduzido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NomeReduzido_Validate(Cancel As Boolean)

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_NomeReduzido_Validate
    
    'Se está preenchido, testa se começa por letra
    If Len(Trim(NomeReduzido.Text)) > 0 Then
    
        If Not IniciaLetra(NomeReduzido.Text) Then Error 25004
        
    End If

    'Preenche os label Fornecedor existentes nos Frames
    For iIndice = 2 To 5
        Fornecedor(iIndice).Caption = NomeReduzido.Text
    Next
    
    Exit Sub

Erro_NomeReduzido_Validate:

    Cancel = True


    Select Case Err

        Case 25004
            
           lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_COMECA_LETRA", Err, NomeReduzido.Text)
    
        Case Else
    
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160610)

    End Select

End Sub

Private Sub Opcao_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index
            
        Select Case iFrameAtual
        
            Case TAB_Identificacao
                Parent.HelpContextID = IDH_FORNECEDOR_IDENT

            Case TAB_DadosFinanceiros
                Parent.HelpContextID = IDH_FORNECEDOR_DADOS_FIN
                
            
            Case TAB_Endereco
                Parent.HelpContextID = IDH_FORNECEDOR_ENDERECO
            
            Case TAB_Estatisticas
                Parent.HelpContextID = IDH_FORNECEDOR_ESTATISTICAS
        
        End Select
    End If

End Sub

Function Exibe_Dados_Fornecedor(objFornecedor As ClassFornecedor, objFornecedorEstatistica As ClassFilialFornecedorEst) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iCodigo As Integer
Dim sContaEnxuta As String
Dim objEndereco As New ClassEndereco
Dim colCategoriaItem As New Collection
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objTipoFornecedor As New ClassTipoFornecedor
Dim objEmi As New ClassCliEmissoresTRP
Dim objEmiExc As New ClassCliEmissoresExcTRP
Dim objProduto As ClassProduto
Dim sProdutoEnxuto As String
Dim colEnderecos As New Collection

On Error GoTo Erro_Exibe_Dados_Fornecedor

    'pesquisa a primeira compra
    lErro = CF("Pesquisa_Primeira_Compra_Forn", objFornecedorEstatistica)
    If lErro <> SUCESSO Then gError 89402

    'pesquisa a ultima compra
    lErro = CF("Pesquisa_Ultima_Compra_Forn", objFornecedorEstatistica)
    If lErro <> SUCESSO Then gError 89403

    'pesquisa a ultima nota fiscal de Devolução de Devolução da FilialEmpresa (data de emissao mais recente) para o Fornecedor em questão.
    lErro = CF("Pesquisa_Ultima_Devolucao_Forn", objFornecedorEstatistica)
    If lErro <> SUCESSO Then gError 89415
    
    objEmi.lFornEmissor = objFornecedor.lCodigo
    
    lErro = CF("TRPEmissores_Le_Forn", objEmi)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 89415

    'IDENTIFICAÇÃO :
    Codigo.Text = CStr(objEmi.lCodigo)
    CodigoForn.Caption = CStr(objFornecedor.lCodigo)
    
    Cliente.Text = objEmi.lCliente
    Call Cliente_Validate(bSGECancelDummy)
    
    Set gobjUsuarioWeb = objEmi.objUsuarioWeb
    
    If objEmi.lSuperior <> 0 Then
        Call Combo_Seleciona_ItemData(Superior, objEmi.lSuperior)
    End If
    
    If objEmi.dPercComissao <> 0 Then
        PercComi.Text = CStr(100 * objEmi.dPercComissao)
    Else
        PercComi.Text = ""
    End If
    
    RazaoSocial.Text = objFornecedor.sRazaoSocial
    NomeReduzido.Text = objFornecedor.sNomeReduzido

    Ativo.Value = objFornecedor.iAtivo

    If objFornecedor.sObservacao <> "" Then
        Observacao.Text = objFornecedor.sObservacao
    Else
        Observacao.Text = ""
    End If
    
    'Preenche os campos Fornecedor existentes nos frames
    For iIndice = 2 To 5
        Fornecedor(iIndice).Caption = objFornecedor.sNomeReduzido
    Next

    'DADOS FINANCEIROS :

    If objFornecedor.dDesconto <> 0 Then
        Desconto.Text = CStr(100 * objFornecedor.dDesconto)
    Else
        Desconto.Text = ""
    End If
    
    If objFornecedor.iCondicaoPagto = 0 Then
        CondicaoPagto.Text = ""
    Else
        CondicaoPagto.Text = CStr(objFornecedor.iCondicaoPagto)
        Call CondicaoPagto_Validate(bSGECancelDummy)
    End If

    ContaContabil.PromptInclude = False
    If Len(Trim(objFornecedor.sContaContabil)) = 0 Then
        ContaContabil.Text = ""
    Else
        lErro = Mascara_RetornaContaEnxuta(objFornecedor.sContaContabil, sContaEnxuta)
        If lErro <> SUCESSO Then gError 12819

        ContaContabil.Text = sContaEnxuta
        
    End If
    ContaContabil.PromptInclude = True

    ContaFornConsig.PromptInclude = False
    If Len(Trim(objFornecedor.sContaFornConsig)) = 0 Then
        ContaFornConsig.Text = ""
    Else
        lErro = Mascara_RetornaContaEnxuta(objFornecedor.sContaFornConsig, sContaEnxuta)
        If lErro <> SUCESSO Then gError 90546

        ContaFornConsig.Text = sContaEnxuta
        
    End If
    ContaFornConsig.PromptInclude = True

    ContaCorrente.Text = objFornecedor.sContaCorrente

    Agencia.Text = objFornecedor.sAgencia

    If objFornecedor.iBanco = 0 Then
        Banco.Text = ""
    Else
        Banco.Text = CStr(objFornecedor.iBanco)
    End If

    'INSCRIÇÕES :

    CGC.Text = objFornecedor.sCgc
    Call CGC_Validate(bSGECancelDummy)
    InscricaoEstadual.Text = objFornecedor.sInscricaoEstadual
    InscricaoMunicipal.Text = objFornecedor.sInscricaoMunicipal
    Observacao2.Text = objFornecedor.sObservacao2

    'ENDEREÇO :

    objEndereco.lCodigo = objFornecedor.lEndereco
    'Lê os dados do endereço do Fornecedor
    lErro = CF("Endereco_Le", objEndereco)
    If lErro <> SUCESSO Then gError 12820

    colEnderecos.Add objEndereco
    
    lErro = gobjTabEnd.Traz_Endereco_Tela(colEnderecos)
    If lErro <> SUCESSO Then gError 12820

    'ESTATÍSTICAS :

    SaldoTitulos.Caption = Format(objFornecedorEstatistica.dSaldoTitulos, "Standard")
    NumeroCompras.Caption = Format(objFornecedorEstatistica.lNumeroCompras, FORMATO_INTEIRO)
    If objFornecedorEstatistica.lNumeroCompras > 0 Then
        MediaCompra.Caption = Format(objFornecedorEstatistica.dValorAcumuladoCompras / objFornecedorEstatistica.lNumeroCompras, "Standard")
    Else
        MediaCompra.Caption = CStr(0)
    End If
    ValorAcumuladoCompras.Caption = Format(objFornecedorEstatistica.dValorAcumuladoCompras, "Standard")
    
    If objFornecedorEstatistica.dtDataPrimeiraCompra = DATA_NULA Then
        DataPrimeiraCompra.Caption = ""
    Else
        DataPrimeiraCompra.Caption = Format(objFornecedorEstatistica.dtDataPrimeiraCompra, "dd/mm/yyyy")
    End If
    
    If objFornecedorEstatistica.dtDataUltimaCompra = DATA_NULA Then
        DataUltimaCompra.Caption = ""
    Else
        DataUltimaCompra.Caption = Format(objFornecedorEstatistica.dtDataUltimaCompra, "dd/mm/yyyy")
    End If

    If objFornecedorEstatistica.lItensPedidosRecebidos > 0 Then
        MediaAtraso.Caption = Format(objFornecedorEstatistica.lAtrasoAcumulado / objFornecedorEstatistica.lItensPedidosRecebidos, FORMATO_MEDIA_ATRASO)
    Else
        MediaAtraso.Caption = CStr(0)
    End If
    
'    MaiorAtraso.Caption = CStr(objFornecedorEstatistica.lMaiorAtraso)

    ValorAcumuladoDevolucoes.Caption = Format(objFornecedorEstatistica.dValorAcumuladoDevolucoes, "Standard")
    NumTotalDevolucoes.Caption = Format(objFornecedorEstatistica.lNumTotalDevolucoes, FORMATO_INTEIRO)

    If objFornecedorEstatistica.dtDataUltDevolucao = DATA_NULA Then
        DataUltDevolucao.Caption = ""
    Else
        DataUltDevolucao.Caption = Format(objFornecedorEstatistica.dtDataUltDevolucao, "dd/mm/yyyy")
    End If
    
   'Lê as categoria deste fornecedor
    objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
    objFilialFornecedor.iCodFilial = FILIAL_MATRIZ
    
    'Alterado em 03/03/04 por Luiz Nogueira
    lErro = CF("FilialFornecedorCategorias_Le_FornFilial", objFilialFornecedor, colCategoriaItem)
    If lErro <> SUCESSO And lErro <> 127085 Then gError 33010

    'Limpa o grid de categorias
    Call Grid_Limpa(objGridCategoria)
    
    'Carrega o Grid Categoria
    If colCategoriaItem.Count > 0 Then
        For iIndice = 1 To colCategoriaItem.Count
            GridCategoria.TextMatrix(iIndice, iGrid_Categoria_Col) = colCategoriaItem.Item(iIndice).sCategoria
            GridCategoria.TextMatrix(iIndice, iGrid_Valor_Col) = colCategoriaItem.Item(iIndice).sItem
        Next
        objGridCategoria.iLinhasExistentes = colCategoriaItem.Count
    End If

    'Alteracao Daniel
    ContaDespesa.PromptInclude = False

    If Len(Trim(objFornecedor.sContaDespesa)) = 0 Then
        ContaDespesa.Text = objFornecedor.sContaDespesa
    Else
        lErro = Mascara_RetornaContaEnxuta(objFornecedor.sContaDespesa, sContaEnxuta)
        If lErro <> SUCESSO Then gError 106004

        ContaDespesa.Text = sContaEnxuta
    End If

    ContaDespesa.PromptInclude = True
    'Fim da Alteracao
    
    'Alteracao Daniel
    InscricaoINSS.Text = objFornecedor.sInscricaoINSS
    
    'Alteracao Daniel
    GeraICMS.Value = objFornecedor.iGeraCredICMS
    TipoFrete.ListIndex = objFornecedor.iTipoFrete
    
    'Busca no ItemData da Combo de Tipo de Cobranca o código
    For iIndice = 0 To TipoCobranca.ListCount - 1
        If TipoCobranca.ItemData(iIndice) = objFornecedor.iTipoCobranca Then
            TipoCobranca.ListIndex = iIndice
            Exit For
        End If
    Next
    'Fim da Alteracao Daniel

    Call Grid_Limpa(objGridExcecoes)
    
    iIndice = 0
    For Each objEmiExc In objEmi.colCliEmissoresExcTRP
    
        iIndice = iIndice + 1
        
        Set objProduto = New ClassProduto
    
        lErro = Mascara_RetornaProdutoTela(objEmiExc.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 195802

        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True

        objProduto.sCodigo = objEmiExc.sProduto
        
        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 195803
        
        'Preenche o Grid
        GridExcecoes.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
        GridExcecoes.TextMatrix(iIndice, iGrid_DescricaoProduto_Col) = objProduto.sDescricao
        GridExcecoes.TextMatrix(iIndice, iGrid_PercComissProd_Col) = Format(objEmiExc.dPercComissao, "Percent")
    
    Next
    
    objGridExcecoes.iLinhasExistentes = objEmi.colCliEmissoresExcTRP.Count
    
    iAlterado = 0

    Exibe_Dados_Fornecedor = SUCESSO

Exit Function

Erro_Exibe_Dados_Fornecedor:

    Exibe_Dados_Fornecedor = gErr

    Select Case gErr

        Case 12733, 12820, 12355, 89402, 89403, 89415
        
        Case 12320
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_SELECIONADO", gErr, objTipoFornecedor.iCodigo)
        
        Case 12319
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_FORNECEDOR_NAO_CADASTRADO", gErr, objTipoFornecedor.iCodigo)

        Case 12819
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objFornecedor.sContaContabil)
        
        Case 90546
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objFornecedor.sContaFornConsig)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160611)

    End Select

    Exit Function

End Function

'"""""""""""""""""""""""""""""""""""""""""""
'"      ROTINAS DAS TELAS DE BROWSE        "
'"""""""""""""""""""""""""""""""""""""""""""

'BROWSE CONDICAO_PAGTO :

Private Sub CondicaoPagtoLabel_Click()

Dim objCondicaoPagto As New ClassCondicaoPagto
Dim colSelecao As New Collection

    If CondicaoPagto.ListIndex <> -1 Then objCondicaoPagto.iCodigo = Codigo_Extrai(CondicaoPagto.Text)

    Call Chama_Tela("CondicaoPagtoCPLista", colSelecao, objCondicaoPagto, objEventoCondicaoPagto)

End Sub

'BROWSE PLANO_CONTA :

Private Sub ContaContabilLabel_Click()

Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim iContaPreenchida As Integer
Dim sConta As String
Dim lErro As Long

On Error GoTo Erro_ContaContabilLabel_Click

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaContabil.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 12774

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    Call Chama_Tela("PlanoContaCPLista", colSelecao, objPlanoConta, objEventoContaContabil)

    Exit Sub

Erro_ContaContabilLabel_Click:

    Select Case Err

    Case 12774

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160612)

    End Select

    Exit Sub

End Sub

Private Sub ContaFornConsigLabel_Click()

Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim iContaPreenchida As Integer
Dim sConta As String
Dim lErro As Long

On Error GoTo Erro_ContaFornConsigLabel_Click

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaFornConsig.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 12774

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    Call Chama_Tela("PlanoContaCPLista", colSelecao, objPlanoConta, objEventoContaFornConsig)

    Exit Sub

Erro_ContaFornConsigLabel_Click:

    Select Case Err

    Case 12774

    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160613)

    End Select

    Exit Sub

End Sub

'RETORNO CONDICAO_PAGTO :

Private Sub objEventoCondicaoPagto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCondicaoPagto As ClassCondicaoPagto

On Error GoTo Erro_objEventoCondicaoPagto_evSelecao

    Set objCondicaoPagto = obj1

    CondicaoPagto.Text = CStr(objCondicaoPagto.iCodigo)
    Call CondicaoPagto_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

Erro_objEventoCondicaoPagto_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160615)

    End Select

    Exit Sub

End Sub

'RETORNO PLANO_CONTA :

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
        If lErro <> SUCESSO Then Error 19108

        ContaContabil.Text = sContaEnxuta

        ContaContabil.PromptInclude = True

    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case Err

        Case 19108
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160616)

    End Select

    Exit Sub

End Sub

Private Sub objEventoContaFornConsig_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaFornConsig_evSelecao

    Set objPlanoConta = obj1

    If objPlanoConta.sConta = "" Then

        ContaFornConsig.Text = ""

    Else

        ContaFornConsig.PromptInclude = False

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then Error 19108

        ContaFornConsig.Text = sContaEnxuta

        ContaFornConsig.PromptInclude = True

    End If

    Me.Show

    Exit Sub

Erro_objEventoContaFornConsig_evSelecao:

    Select Case Err

        Case 19108
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160617)

    End Select

    Exit Sub

End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA
'""""""""""""""""""""""""""""""""""""""""""""""

'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objEmi As New ClassCliEmissoresTRP

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TRPEmissores"

    'Lê os dados da Tela Fornecedores
    lErro = Le_Dados_Fornecedor(objFornecedor, objEmi)
    If lErro <> SUCESSO Then Error 12863

    'Preenche a coleção colCampoValor, com nome do campo,
    'Valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo

    colCampoValor.Add "Codigo", objEmi.lCodigo, 0, "Codigo"
    colCampoValor.Add "NomeReduzido", objFornecedor.sNomeReduzido, STRING_FORNECEDOR_NOME_REDUZIDO, "NomeReduzido"


    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 12863

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160618)

    End Select

    Exit Sub

End Sub

'Preenche os campos da tela com os correspondentes do BD
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objEmi As New ClassCliEmissoresTRP
'Dim objFornecedorEstatistica As New ClassFilialFornecedorEst
'Dim objFilialFornecedor As New ClassFilialFornecedor

On Error GoTo Erro_Tela_Preenche

    objEmi.lCodigo = colCampoValor.Item("Codigo").vValor
    
    lErro = Trata_Parametros(objEmi)
    If lErro <> SUCESSO Then gError 19215

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 19215
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INEXISTENTE", gErr)

        Case 19214, 19216

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160619)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Alteracao Daniel
    Set objEventoContaDespesa = Nothing
    
    Set objEventoTipoFornecedor = Nothing
    Set objEventoCondicaoPagto = Nothing
    Set objEventoContaContabil = Nothing
    Set objEventoContaFornConsig = Nothing
    Set objEventoFornecedor = Nothing
    Set objGridExcecoes = Nothing
    Set objEventoCliente = Nothing
    Set objEventoProduto = Nothing
    Set objEventoBancos = Nothing
    Set objEventoEmissor = Nothing
    
    Set gobjAnotacao = Nothing

    Set gobjUsuarioWeb = Nothing
    
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object
    
    Parent.HelpContextID = IDH_FORNECEDOR_IDENT
    Set Form_Load_Ocx = Me
    Caption = "Emissores"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "TRPEmissores"
    
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
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is ContaDespesa Then
            Call LabelContaDespesa_Click
        ElseIf Me.ActiveControl Is CondicaoPagto Then
            Call CondicaoPagtoLabel_Click
        ElseIf Me.ActiveControl Is ContaContabil Then
            Call ContaContabilLabel_Click
        
        ElseIf Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        ElseIf Me.ActiveControl Is NomeReduzido Then
            Call LabelNomeReduzido_Click
        ElseIf Me.ActiveControl Is ContaFornConsig Then
            Call ContaFornConsigLabel_Click
        ElseIf Me.ActiveControl Is Cliente Then
            Call LabelCliente_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
        ElseIf Me.ActiveControl Is Banco Then
            Call LabelBanco_Click
        End If
        
    End If
    
End Sub

Private Sub Fornecedor_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Fornecedor(Index), Source, X, Y)
End Sub

Private Sub Fornecedor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Fornecedor(Index), Button, Shift, X, Y)
End Sub

Private Sub LabelObservacao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelObservacao, Source, X, Y)
End Sub

Private Sub LabelObservacao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelObservacao, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeReduzido_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeReduzido, Source, X, Y)
End Sub

Private Sub LabelNomeReduzido_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeReduzido, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label26_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label26, Source, X, Y)
End Sub

Private Sub Label26_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label26, Button, Shift, X, Y)
End Sub

Private Sub Label27_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label27, Source, X, Y)
End Sub

Private Sub Label27_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label27, Button, Shift, X, Y)
End Sub

Private Sub MediaAtraso_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(MediaAtraso, Source, X, Y)
End Sub

Private Sub MediaAtraso_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(MediaAtraso, Button, Shift, X, Y)
End Sub

Private Sub MaiorAtraso_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(MaiorAtraso, Source, X, Y)
End Sub

Private Sub MaiorAtraso_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(MaiorAtraso, Button, Shift, X, Y)
End Sub

Private Sub Label29_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label29, Source, X, Y)
End Sub

Private Sub Label29_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label29, Button, Shift, X, Y)
End Sub

Private Sub PedidosEmAberto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PedidosEmAberto, Source, X, Y)
End Sub

Private Sub PedidosEmAberto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PedidosEmAberto, Button, Shift, X, Y)
End Sub

Private Sub Label23_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label23, Source, X, Y)
End Sub

Private Sub Label23_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label23, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub ValorAcumuladoDevolucoes_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorAcumuladoDevolucoes, Source, X, Y)
End Sub

Private Sub ValorAcumuladoDevolucoes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorAcumuladoDevolucoes, Button, Shift, X, Y)
End Sub

Private Sub Label19_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label19, Source, X, Y)
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label19, Button, Shift, X, Y)
End Sub

Private Sub NumTotalDevolucoes_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumTotalDevolucoes, Source, X, Y)
End Sub

Private Sub NumTotalDevolucoes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumTotalDevolucoes, Button, Shift, X, Y)
End Sub

Private Sub DataUltDevolucao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataUltDevolucao, Source, X, Y)
End Sub

Private Sub DataUltDevolucao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataUltDevolucao, Button, Shift, X, Y)
End Sub

Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub

Private Sub DataUltimaCompra_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataUltimaCompra, Source, X, Y)
End Sub

Private Sub DataUltimaCompra_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataUltimaCompra, Button, Shift, X, Y)
End Sub

Private Sub DataPrimeiraCompra_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataPrimeiraCompra, Source, X, Y)
End Sub

Private Sub DataPrimeiraCompra_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataPrimeiraCompra, Button, Shift, X, Y)
End Sub

Private Sub ValorAcumuladoCompras_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorAcumuladoCompras, Source, X, Y)
End Sub

Private Sub ValorAcumuladoCompras_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorAcumuladoCompras, Button, Shift, X, Y)
End Sub

Private Sub MediaCompra_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(MediaCompra, Source, X, Y)
End Sub

Private Sub MediaCompra_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(MediaCompra, Button, Shift, X, Y)
End Sub

Private Sub NumeroCompras_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumeroCompras, Source, X, Y)
End Sub

Private Sub NumeroCompras_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumeroCompras, Button, Shift, X, Y)
End Sub

Private Sub Label18_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label18, Source, X, Y)
End Sub

Private Sub Label18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label18, Button, Shift, X, Y)
End Sub

Private Sub Label17_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label17, Source, X, Y)
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label17, Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label15, Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15, Button, Shift, X, Y)
End Sub

Private Sub Label16_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label16, Source, X, Y)
End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label16, Button, Shift, X, Y)
End Sub

Private Sub Label21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label21, Source, X, Y)
End Sub

Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label21, Button, Shift, X, Y)
End Sub

Private Sub SaldoTitulos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SaldoTitulos, Source, X, Y)
End Sub

Private Sub SaldoTitulos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SaldoTitulos, Button, Shift, X, Y)
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

Private Sub Label36_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label36, Source, X, Y)
End Sub

Private Sub Label36_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label36, Button, Shift, X, Y)
End Sub

Private Sub Label34_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label34, Source, X, Y)
End Sub

Private Sub Label34_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label34, Button, Shift, X, Y)
End Sub

Private Sub Label35_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label35, Source, X, Y)
End Sub

Private Sub Label35_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label35, Button, Shift, X, Y)
End Sub

Private Sub Label32_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label32, Source, X, Y)
End Sub

Private Sub Label32_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label32, Button, Shift, X, Y)
End Sub

Private Sub Label22_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label22, Source, X, Y)
End Sub

Private Sub Label22_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label22, Button, Shift, X, Y)
End Sub

Private Sub Label25_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label25, Source, X, Y)
End Sub

Private Sub Label25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label25, Button, Shift, X, Y)
End Sub

Private Sub LabelBanco_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelBanco, Source, X, Y)
End Sub

Private Sub LabelBanco_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelBanco, Button, Shift, X, Y)
End Sub

Private Sub Label24_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label24, Source, X, Y)
End Sub

Private Sub Label24_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label24, Button, Shift, X, Y)
End Sub

Private Sub CondicaoPagtoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CondicaoPagtoLabel, Source, X, Y)
End Sub

Private Sub CondicaoPagtoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CondicaoPagtoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub ContaContabilLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContaContabilLabel, Source, X, Y)
End Sub

Private Sub ContaFornConsigLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContaFornConsigLabel, Source, X, Y)
End Sub

Private Sub ContaContabilLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContaContabilLabel, Button, Shift, X, Y)
End Sub

Private Sub ContaFornConsigLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContaFornConsigLabel, Button, Shift, X, Y)
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

'Alteracao Daniel
Private Sub GeraICMS_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

'Alteracao Daniel
Private Sub InscricaoINSS_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

'Alteracao Daniel
Private Sub ContaDespesa_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

'Alteracao Daniel
Private Sub TipoCobranca_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

'Alteracao Daniel
Private Sub TipoFrete_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

'Alteracao Daniel
Private Sub LabelContaDespesa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelContaDespesa, Source, X, Y)
End Sub

'Alteracao Daniel
Private Sub LabelContaDespesa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelContaDespesa, Button, Shift, X, Y)
End Sub

'Alteracao Daniel
Private Sub ContaDespesa_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaMascarada As String
Dim sContaFormatada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ContaDespesa_Validate

    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica_Modulo", ContaDespesa.Text, ContaDespesa.ClipText, objPlanoConta, MODULO_CONTASAPAGAR)
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 101499

    If lErro = SUCESSO Then

        sContaFormatada = objPlanoConta.sConta

        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then gError 106000
        
        ContaDespesa.PromptInclude = False
        ContaDespesa.Text = sContaMascarada
        ContaDespesa.PromptInclude = True
        
    'se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then

        lErro = CF("Conta_Critica", ContaDespesa.Text, sContaFormatada, objPlanoConta, MODULO_CONTASAPAGAR)
        If lErro <> SUCESSO And lErro <> 5700 Then gError 106001

        If lErro = 5700 Then gError 106002

    End If

    Exit Sub

Erro_ContaDespesa_Validate:

    Cancel = True

    Select Case gErr

        Case 106001, 101499

        Case 106002
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", gErr, ContaDespesa.Text)

        Case 106000
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160620)

    End Select

End Sub

'Alteracao Daniel
Private Sub objEventoContaDespesa_evSelecao(obj1 As Object)

Dim objPlanoConta As ClassPlanoConta
Dim lErro As Long
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaDespesa_evSelecao

    Set objPlanoConta = obj1

    If objPlanoConta.sConta = "" Then

        ContaDespesa.Text = ""

    Else

        ContaDespesa.PromptInclude = False

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then gError 106005

        ContaDespesa.Text = sContaEnxuta

        ContaDespesa.PromptInclude = True

    End If

    Me.Show

    Exit Sub

Erro_objEventoContaDespesa_evSelecao:

    Select Case gErr

        Case 106005
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160621)

    End Select

End Sub

'Alteracao Daniel
Private Function Carrega_TipoCobranca() As Long
'Carrega na combobox os Tipos de Cobrança

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome

On Error GoTo Erro_Carrega_TipoCobranca

    'Lê o nome e o codigo de todos os Tipos de Cobrança
    lErro = CF("Cod_Nomes_Le", "TiposDeCobranca", "Codigo", "Descricao", 255, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 106007

    'Carrega na combo de Tipos de Cobrança os Tipos que estão em colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao

        TipoCobranca.AddItem CInt(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        TipoCobranca.ItemData(TipoCobranca.NewIndex) = objCodigoDescricao.iCodigo

    Next

    Carrega_TipoCobranca = SUCESSO

    Exit Function

Erro_Carrega_TipoCobranca:

    Carrega_TipoCobranca = gErr

    Select Case gErr

        Case 106007

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160622)

    End Select

End Function

'Alteracao Daniel
Private Sub LabelContaDespesa_Click()

Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim iContaPreenchida As Integer
Dim sConta As String
Dim lErro As Long

On Error GoTo Erro_LabelContaDespesa_Click

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaDespesa.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then gError 106008

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    Call Chama_Tela("PlanoContaCPLista", colSelecao, objPlanoConta, objEventoContaDespesa)

    Exit Sub

Erro_LabelContaDespesa_Click:

    Select Case gErr

    Case 106008

    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160623)

    End Select

End Sub

Public Sub Anotacao_Extrai(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long

On Error GoTo Erro_Anotacao_Extrai

    objAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_FORNECEDOR
    If Len(Trim(Codigo.Text)) > 0 Then
        objAnotacao.sID = Codigo.Text
    Else
        objAnotacao.sID = ""
        If Not (gobjAnotacao Is Nothing) Then
            objAnotacao.sTextoCompleto = gobjAnotacao.sTextoCompleto
            objAnotacao.sTitulo = gobjAnotacao.sTitulo
        End If
    End If
    
    Exit Sub
     
Erro_Anotacao_Extrai:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160626)
     
    End Select
     
    Exit Sub

End Sub

Public Sub Anotacao_Preenche(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long

On Error GoTo Erro_Anotacao_Preenche

    'guarda o texto digitado
    Set gobjAnotacao = objAnotacao
        
    Exit Sub
     
Erro_Anotacao_Preenche:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160627)
     
    End Select
     
    Exit Sub

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome
Dim iCodFilial As Integer

On Error GoTo Erro_Cliente_Validate

    'Verifica se o Cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        'Busca o Cliente no BD
        lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
        If lErro <> SUCESSO Then gError 190635
    
        If sClienteAnt <> Cliente.Text Then
        
            Superior.Clear
        
            'Lê cada código e descrição da tabela Paises
            lErro = CF("Cod_Nomes_Le", "WebEmissores", "Codigo", "NomeReduzido", STRING_FORNECEDOR_NOME_REDUZIDO, colCodigoDescricao, "Cliente = " & CStr(objCliente.lCodigo))
            If lErro <> SUCESSO Then gError 12721
        
            'Preenche cada ComboBox País com os objetos da coleção colCodigoDescricao
            For Each objCodigoDescricao In colCodigoDescricao
        
                Superior.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
                Superior.ItemData(Superior.NewIndex) = objCodigoDescricao.iCodigo
        
            Next
            
        End If
    
    End If
    
    sClienteAnt = Cliente.Text
    
    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 190635, 190636, 198634, 198635

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190637)

    End Select

    Exit Sub

End Sub

Private Sub Cliente_GotFocus()
    Call MaskEdBox_TrataGotFocus(Cliente, iAlterado)
End Sub

Private Sub Cliente_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Function Inicializa_Grid_Excecoes(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("% Comissão")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoProduto.Name)
    objGridInt.colCampo.Add (PercComissProd.Name)

    'Colunas do Grid
    iGrid_Produto_Col = 1
    iGrid_DescricaoProduto_Col = 2
    iGrid_PercComissProd_Col = 3

    'Grid do GridInterno
    objGridInt.objGrid = GridExcecoes

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 201

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 6

    'Largura da primeira coluna
    GridExcecoes.ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Excecoes = SUCESSO

    Exit Function

End Function

Public Sub GridExcecoes_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridExcecoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridExcecoes, iAlterado)
    End If
    
End Sub

Public Sub GridExcecoes_EnterCell()
    Call Grid_Entrada_Celula(objGridExcecoes, iAlterado)
End Sub

Public Sub GridExcecoes_GotFocus()
    Call Grid_Recebe_Foco(objGridExcecoes)
End Sub

Public Sub GridExcecoes_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridExcecoes)
End Sub

Public Sub GridExcecoes_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridExcecoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridExcecoes, iAlterado)
    End If
    
End Sub

Public Sub GridExcecoes_LeaveCell()
    Call Saida_Celula(objGridExcecoes)
End Sub

Public Sub GridExcecoes_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridExcecoes)
End Sub

Public Sub GridExcecoes_RowColChange()
    Call Grid_RowColChange(objGridExcecoes)
End Sub

Public Sub GridExcecoes_Scroll()
    Call Grid_Scroll(objGridExcecoes)
End Sub


Public Sub Produto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Produto_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridExcecoes)
End Sub

Public Sub Produto_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridExcecoes)
End Sub

Public Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Produto_Validate

    Set objGridExcecoes.objControle = Produto
    
    lErro = Grid_Campo_Libera_Foco(objGridExcecoes)
    If lErro <> SUCESSO Then gError 195835
    
    Exit Sub
    
Erro_Produto_Validate:

    Cancel = True

    Select Case gErr
    
        Case 195835
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195836)
     
    End Select
    
End Sub

Public Sub PercComissProd_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercComissProd_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridExcecoes)
End Sub

Public Sub PercComissProd_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridExcecoes)
End Sub

Public Sub PercComissProd_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercComissProd_Validate

    Set objGridExcecoes.objControle = PercComissProd
    lErro = Grid_Campo_Libera_Foco(objGridExcecoes)
    If lErro <> SUCESSO Then gError 195838
    
    Exit Sub
    
Erro_PercComissProd_Validate:

    Cancel = True

    Select Case gErr
    
        Case 195838
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195839)
     
    End Select
    
End Sub

Public Function Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    Select Case objControl.Name

        Case Produto.Name
            objControl.Enabled = True
 
        Case DescricaoProduto.Name
            objControl.Enabled = False
 
        Case PercComissProd.Name
            
            lErro = CF("Produto_Formata", GridExcecoes.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 195801
            
            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
 
    End Select
    
    Rotina_Grid_Enable = SUCESSO

    Exit Function

Erro_Rotina_Grid_Enable:

    Rotina_Grid_Enable = gErr

    Select Case gErr
    
        Case 195801
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195804)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Produto que está deixando de ser a corrente

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iIndice As Integer
Dim iProdutoPreenchido As Integer
Dim sProdutoEnxuto As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    If Len(Trim(Produto.ClipText)) > 0 Then

        'Critica o Produto
        lErro = CF("Produto_Critica_Filial2", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then gError 195791
        
        'Se o produto é gerencial ==> erro
        If lErro = 86295 Then gError 195792
               
        'Se o produto não foi encontrado ==> Pergunta se deseja criar
        If lErro = 51381 Then gError 195793

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 195794
    
        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True
    
        'Verifica se já está em outra linha do Grid
        For iIndice = 1 To objGridInt.iLinhasExistentes
            If iIndice <> GridExcecoes.Row Then
                If GridExcecoes.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text Then gError 195795
            End If
        Next
    
        If GridExcecoes.Row - GridExcecoes.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
        GridExcecoes.TextMatrix(GridExcecoes.Row, iGrid_DescricaoProduto_Col) = objProduto.sDescricao

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195796

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 195791, 195796
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 195792
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 195793
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)
            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 195794
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 195795
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_PREENCHIDO_LINHA_GRID", gErr, iIndice)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195797)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Function Saida_Celula_PercComissProd(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Percentual Desconto que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentDesc As Double

On Error GoTo Erro_Saida_Celula_PercComissProd

    Set objGridInt.objControle = PercComissProd

    If Len(PercComissProd.Text) > 0 Then
        
        'Critica a porcentagem
        lErro = Porcentagem_Critica(PercComissProd.Text)
        If lErro <> SUCESSO Then gError 195798

        dPercentDesc = CDbl(PercComissProd.Text)
        
        PercComissProd.Text = Format(dPercentDesc, "Fixed")
        
        If GridExcecoes.Row - GridExcecoes.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195799

    Saida_Celula_PercComissProd = SUCESSO

    Exit Function

Erro_Saida_Celula_PercComissProd:

    Saida_Celula_PercComissProd = gErr

    Select Case gErr

        Case 195798, 195799
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195800)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Sub PercComi_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PercComi_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sPercComi As String

On Error GoTo Erro_PercComi_Validate

    sPercComi = PercComi.Text

    'Verifica se foi preenchido o PercComi
    If Len(Trim(PercComi.Text)) = 0 Then Exit Sub

    'Critica se é porcentagem
    lErro = Porcentagem_Critica(PercComi.Text)
    If lErro <> SUCESSO Then gError 12776

    PercComi.Text = Format(sPercComi, "Fixed")

    Exit Sub

Erro_PercComi_Validate:

    Cancel = True

    Select Case gErr

        Case 12776
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160595)
    
        End Select

    Exit Sub

End Sub

Public Sub BotaoProdutos_Click()

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As Collection
Dim sProduto1 As String

On Error GoTo Erro_BotaoProdutos_Click

    If Me.ActiveControl Is Produto Then
        sProduto1 = Produto.Text
    Else
        'Verifica se tem alguma linha selecionada no Grid
        If GridExcecoes.Row = 0 Then gError 195827

        sProduto1 = GridExcecoes.TextMatrix(GridExcecoes.Row, iGrid_Produto_Col)
    End If
    
    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 195828
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    objProduto.sCodigo = sProduto
    
    'Chama a Tela ProdutoVendaLista
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoProduto)

    Exit Sub
    
Erro_BotaoProdutos_Click:

    Select Case gErr
    
        Case 195827
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 195828
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195829)
    
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProdutoEnxuto As String
Dim iIndice As Integer

On Error GoTo Erro_objEventoProduto_evSelecao

    'verifica se tem alguma linha do Grid selecionada
    If GridExcecoes.Row > 0 Then

        Set objProduto = obj1

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 195831

        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True
        
        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 195832
        
        If Not (Me.ActiveControl Is Produto) Then
    
            'Verifica se já está em outra linha do Grid
            For iIndice = 1 To objGridExcecoes.iLinhasExistentes
                If iIndice <> GridExcecoes.Row Then
                    If GridExcecoes.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text Then gError 195839
                End If
            Next
    
            'Preenche o Grid
            GridExcecoes.TextMatrix(GridExcecoes.Row, iGrid_Produto_Col) = Produto.Text
            GridExcecoes.TextMatrix(GridExcecoes.Row, iGrid_DescricaoProduto_Col) = objProduto.sDescricao
    
            If GridExcecoes.Row - GridExcecoes.FixedRows = objGridExcecoes.iLinhasExistentes Then
                objGridExcecoes.iLinhasExistentes = objGridExcecoes.iLinhasExistentes + 1
            End If
    
        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    GridExcecoes.TextMatrix(GridExcecoes.Row, iGrid_Produto_Col) = ""

    Select Case gErr

        Case 195830, 195832
        
        Case 195831
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
        
        Case 195839
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_PREENCHIDO_LINHA_GRID", gErr, iIndice)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195833)

    End Select

    Exit Sub

End Sub

Public Sub LabelCliente_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    objCliente.sNomeReduzido = Cliente.Text

    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    'Preenche campo Cliente
    Cliente.Text = objCliente.sNomeReduzido

    'Executa o Validate
    Call Cliente_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Public Sub LabelBanco_Click()

Dim objBanco As New ClassBanco
Dim colSelecao As New Collection

    objBanco.iCodBanco = StrParaInt(Banco.Text)

    Call Chama_Tela("BancoLista", colSelecao, objBanco, objEventoBancos)

End Sub

Private Sub objEventoBancos_evSelecao(obj1 As Object)

Dim objBanco As ClassBanco
Dim bCancel As Boolean

    Set objBanco = obj1

    Banco.PromptInclude = False
    Banco.Text = CStr(objBanco.iCodBanco)
    Banco.PromptInclude = True

    Me.Show

    Exit Sub

End Sub

