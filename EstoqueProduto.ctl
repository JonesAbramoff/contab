VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl EstoqueProduto 
   ClientHeight    =   5760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9210
   KeyPreview      =   -1  'True
   ScaleHeight     =   5760
   ScaleWidth      =   9210
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4725
      Index           =   2
      Left            =   225
      TabIndex        =   7
      Top             =   915
      Visible         =   0   'False
      Width           =   8760
      Begin VB.Frame Frame4 
         Caption         =   "Saldos"
         Height          =   3750
         Left            =   345
         TabIndex        =   23
         Top             =   855
         Width           =   8295
         Begin VB.Label Label40 
            Caption         =   "Em Pedido de Compra:"
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
            Left            =   615
            TabIndex        =   120
            Top             =   3405
            Width           =   1965
         End
         Begin VB.Label Label39 
            Caption         =   "Em Requisição de Compra:"
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
            Left            =   285
            TabIndex        =   119
            Top             =   3075
            Width           =   2295
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Ordem de Produção:"
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
            Left            =   825
            TabIndex        =   118
            Top             =   2400
            Width           =   1755
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Empenhada:"
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
            Left            =   1515
            TabIndex        =   117
            Top             =   2085
            Width           =   1065
         End
         Begin VB.Label QuantPedidoLabel 
            AutoSize        =   -1  'True
            Caption         =   "Em Pedido de Venda:"
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
            Left            =   735
            TabIndex        =   116
            Top             =   2730
            Width           =   1845
         End
         Begin VB.Label LabelValPedCompra 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4500
            TabIndex        =   115
            Top             =   3375
            Width           =   1590
         End
         Begin VB.Label LabelValReqCompra 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4500
            TabIndex        =   114
            Top             =   3045
            Width           =   1590
         End
         Begin VB.Label LabelQuantPedCompra 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2700
            TabIndex        =   113
            Top             =   3375
            Width           =   1590
         End
         Begin VB.Label LabelQuantReqCompra 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2700
            TabIndex        =   112
            Top             =   3045
            Width           =   1590
         End
         Begin VB.Label ValorDispTotal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4500
            TabIndex        =   111
            Top             =   360
            Width           =   1590
         End
         Begin VB.Label ValorInd 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4500
            TabIndex        =   110
            Top             =   690
            Width           =   1590
         End
         Begin VB.Label QuantDispTotal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2700
            TabIndex        =   109
            Top             =   360
            Width           =   1590
         End
         Begin VB.Label QuantInd 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2700
            TabIndex        =   108
            Top             =   690
            Width           =   1590
         End
         Begin VB.Label ValorTotal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4500
            TabIndex        =   95
            Top             =   1695
            Width           =   1590
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Totais:"
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
            Left            =   1965
            TabIndex        =   94
            Top             =   1695
            Width           =   600
         End
         Begin VB.Label QuantTotal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2700
            TabIndex        =   93
            Top             =   1695
            Width           =   1590
         End
         Begin VB.Label QuantPedido 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2700
            TabIndex        =   90
            Top             =   2715
            Width           =   1590
         End
         Begin VB.Label ValorPedido 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4500
            TabIndex        =   89
            Top             =   2715
            Width           =   1590
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Não Disponível:"
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
            Left            =   1185
            TabIndex        =   37
            ToolTipText     =   "Quantidade recebida mas ainda não disponível"
            Top             =   735
            Width           =   1395
         End
         Begin VB.Label QuantDefeituosa 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2700
            TabIndex        =   36
            Top             =   1020
            Width           =   1590
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Defeituosa:"
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
            Left            =   1590
            TabIndex        =   35
            Top             =   1065
            Width           =   990
         End
         Begin VB.Label QuantEmpenhada 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2700
            TabIndex        =   34
            Top             =   2055
            Width           =   1590
         End
         Begin VB.Label QuantReservada 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2700
            TabIndex        =   33
            Top             =   1365
            Width           =   1590
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Reservada:"
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
            Left            =   1590
            TabIndex        =   32
            Top             =   1380
            Width           =   990
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Disponível:"
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
            Left            =   1590
            TabIndex        =   31
            Top             =   390
            Width           =   990
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Quantidades"
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
            Left            =   2955
            TabIndex        =   30
            Top             =   135
            Width           =   1080
         End
         Begin VB.Label ValorDefeituosa 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4500
            TabIndex        =   29
            Top             =   1020
            Width           =   1590
         End
         Begin VB.Label ValorEmpenhada 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4500
            TabIndex        =   28
            Top             =   2055
            Width           =   1590
         End
         Begin VB.Label ValorReservada 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4500
            TabIndex        =   27
            Top             =   1365
            Width           =   1590
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Valores"
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
            Left            =   4905
            TabIndex        =   26
            Top             =   135
            Width           =   645
         End
         Begin VB.Label ValorOP 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4500
            TabIndex        =   25
            Top             =   2385
            Width           =   1590
         End
         Begin VB.Label QuantOP 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2700
            TabIndex        =   24
            Top             =   2385
            Width           =   1590
         End
      End
      Begin VB.Label CustoLabel 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   7500
         TabIndex        =   92
         Top             =   135
         Width           =   1155
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "Custo:"
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
         Left            =   6885
         TabIndex        =   91
         Top             =   165
         Width           =   555
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Data do Último Inventário:"
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
         TabIndex        =   45
         Top             =   570
         Width           =   2250
      End
      Begin VB.Label DataInventario 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   7500
         TabIndex        =   44
         Top             =   540
         Width           =   1155
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Almoxarifado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   450
         TabIndex        =   43
         Top             =   570
         Width           =   1140
      End
      Begin VB.Label UnidMedLabel 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   0
         Left            =   4500
         TabIndex        =   42
         Top             =   150
         Width           =   780
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "U.M.:"
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
         Left            =   3960
         TabIndex        =   41
         Top             =   165
         Width           =   480
      End
      Begin VB.Label ProdutoLabel 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   1665
         TabIndex        =   40
         Top             =   135
         Width           =   1245
      End
      Begin VB.Label AlmoxarifadoLabel 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   1665
         TabIndex        =   39
         Top             =   540
         Width           =   2760
      End
      Begin VB.Label Label31 
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
         Height          =   165
         Left            =   870
         TabIndex        =   38
         Top             =   165
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4725
      Index           =   1
      Left            =   165
      TabIndex        =   0
      Top             =   915
      Width           =   8760
      Begin VB.CheckBox Padrao 
         Caption         =   "Padrão"
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
         Left            =   4890
         TabIndex        =   3
         Top             =   2302
         Width           =   945
      End
      Begin VB.ComboBox Almoxarifado 
         Height          =   315
         ItemData        =   "EstoqueProduto.ctx":0000
         Left            =   1950
         List            =   "EstoqueProduto.ctx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2242
         Width           =   2670
      End
      Begin VB.TextBox LocalizacaoFisica 
         Height          =   315
         Left            =   1950
         MaxLength       =   20
         TabIndex        =   4
         Top             =   3033
         Width           =   3840
      End
      Begin MSMask.MaskEdBox Produto 
         Height          =   315
         Left            =   1950
         TabIndex        =   1
         Top             =   675
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSComctlLib.TreeView TvwProduto 
         Height          =   3870
         Left            =   6000
         TabIndex        =   6
         Top             =   420
         Visible         =   0   'False
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   6826
         _Version        =   393217
         Indentation     =   453
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin MSMask.MaskEdBox ContaContabil 
         Height          =   315
         Left            =   1950
         TabIndex        =   5
         Top             =   3825
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "U.M.:"
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
         Left            =   4470
         TabIndex        =   22
         Top             =   720
         Width           =   480
      End
      Begin VB.Label UnidMed 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5010
         TabIndex        =   21
         Top             =   660
         Width           =   780
      End
      Begin VB.Label LabelProduto 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   6075
         TabIndex        =   20
         Top             =   150
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Localização Física:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   210
         TabIndex        =   19
         Top             =   3075
         Width           =   1680
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Almoxarifado:"
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
         Height          =   165
         Left            =   735
         TabIndex        =   18
         Top             =   2317
         Width           =   1155
      End
      Begin VB.Label Descricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1950
         TabIndex        =   17
         Top             =   1451
         Width           =   3840
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
         Height          =   165
         Left            =   960
         TabIndex        =   16
         Top             =   1500
         Width           =   930
      End
      Begin VB.Label ProdutoLabel1 
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
         Height          =   165
         Left            =   1155
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   15
         Top             =   720
         Width           =   735
      End
      Begin VB.Label ContaContabilLabel 
         AutoSize        =   -1  'True
         Caption         =   "Conta de Estoque:"
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
         Left            =   300
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   14
         ToolTipText     =   "Conta contábil de estoque"
         Top             =   3870
         Width           =   1590
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4725
      Index           =   3
      Left            =   195
      TabIndex        =   8
      Top             =   900
      Visible         =   0   'False
      Width           =   8760
      Begin VB.Frame Frame3 
         Caption         =   "Saldos de Terceiros em Nosso Poder"
         Height          =   1965
         Left            =   165
         TabIndex        =   64
         Top             =   2550
         Width           =   8475
         Begin VB.Label CustoBenef3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6675
            TabIndex        =   107
            Top             =   1620
            Width           =   1590
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "Custos"
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
            TabIndex        =   106
            Top             =   195
            Width           =   585
         End
         Begin VB.Label CustoConserto3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6675
            TabIndex        =   105
            Top             =   480
            Width           =   1590
         End
         Begin VB.Label CustoConsig3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6675
            TabIndex        =   104
            Top             =   765
            Width           =   1590
         End
         Begin VB.Label CustoDemo3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6675
            TabIndex        =   103
            Top             =   1050
            Width           =   1590
         End
         Begin VB.Label CustoOutras3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6675
            TabIndex        =   102
            Top             =   1335
            Width           =   1590
         End
         Begin VB.Label QuantOutras3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3135
            TabIndex        =   81
            Top             =   1335
            Width           =   1590
         End
         Begin VB.Label QuantDemo3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3135
            TabIndex        =   80
            Top             =   1050
            Width           =   1590
         End
         Begin VB.Label QuantConsig3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3135
            TabIndex        =   79
            Top             =   765
            Width           =   1590
         End
         Begin VB.Label QuantConserto3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3135
            TabIndex        =   78
            Top             =   480
            Width           =   1590
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Outras:"
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
            Left            =   2385
            TabIndex        =   77
            Top             =   1335
            Width           =   630
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Em Demonstração:"
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
            Left            =   1410
            TabIndex        =   76
            Top             =   1065
            Width           =   1605
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Em Consignação:"
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
            Left            =   1530
            TabIndex        =   75
            Top             =   780
            Width           =   1485
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Em Conserto:"
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
            Left            =   1875
            TabIndex        =   74
            Top             =   480
            Width           =   1140
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Quantidades"
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
            Left            =   3375
            TabIndex        =   73
            Top             =   225
            Width           =   1080
         End
         Begin VB.Label ValorOutras3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4920
            TabIndex        =   72
            Top             =   1335
            Width           =   1590
         End
         Begin VB.Label ValorDemo3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4920
            TabIndex        =   71
            Top             =   1050
            Width           =   1590
         End
         Begin VB.Label ValorConsig3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4920
            TabIndex        =   70
            Top             =   780
            Width           =   1590
         End
         Begin VB.Label ValorConserto3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4920
            TabIndex        =   69
            Top             =   480
            Width           =   1590
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Valores"
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
            TabIndex        =   68
            Top             =   195
            Width           =   645
         End
         Begin VB.Label ValorBenef3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4935
            TabIndex        =   67
            Top             =   1620
            Width           =   1590
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Em Beneficiamento:"
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
            TabIndex        =   66
            Top             =   1620
            Width           =   1695
         End
         Begin VB.Label QuantBenef3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3135
            TabIndex        =   65
            Top             =   1620
            Width           =   1590
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Saldos Nosso em Poder de Terceiros"
         Height          =   1965
         Left            =   165
         TabIndex        =   46
         Top             =   510
         Width           =   8475
         Begin VB.Label CustoBenef 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6705
            TabIndex        =   101
            Top             =   1635
            Width           =   1590
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "Custos"
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
            Left            =   7065
            TabIndex        =   100
            Top             =   225
            Width           =   585
         End
         Begin VB.Label CustoConserto 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6705
            TabIndex        =   99
            Top             =   465
            Width           =   1590
         End
         Begin VB.Label CustoConsig 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6705
            TabIndex        =   98
            Top             =   750
            Width           =   1590
         End
         Begin VB.Label CustoDemo 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6705
            TabIndex        =   97
            Top             =   1035
            Width           =   1590
         End
         Begin VB.Label CustoOutras 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6705
            TabIndex        =   96
            Top             =   1335
            Width           =   1590
         End
         Begin VB.Label QuantOutras 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3165
            TabIndex        =   63
            Top             =   1335
            Width           =   1590
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Outras:"
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
            Left            =   2445
            TabIndex        =   62
            Top             =   1335
            Width           =   630
         End
         Begin VB.Label QuantDemo 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3165
            TabIndex        =   61
            Top             =   1035
            Width           =   1590
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Em Demonstração:"
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
            TabIndex        =   60
            Top             =   1050
            Width           =   1605
         End
         Begin VB.Label QuantConsig 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3165
            TabIndex        =   59
            Top             =   750
            Width           =   1590
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Em Consignação:"
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
            Left            =   1590
            TabIndex        =   58
            Top             =   780
            Width           =   1485
         End
         Begin VB.Label QuantConserto 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3165
            TabIndex        =   57
            Top             =   480
            Width           =   1590
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Em Conserto:"
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
            Left            =   1935
            TabIndex        =   56
            Top             =   495
            Width           =   1140
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Quantidades"
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
            Left            =   3375
            TabIndex        =   55
            Top             =   210
            Width           =   1080
         End
         Begin VB.Label ValorOutras 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4950
            TabIndex        =   54
            Top             =   1335
            Width           =   1590
         End
         Begin VB.Label ValorDemo 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4950
            TabIndex        =   53
            Top             =   1035
            Width           =   1590
         End
         Begin VB.Label ValorConsig 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4950
            TabIndex        =   52
            Top             =   750
            Width           =   1590
         End
         Begin VB.Label ValorConserto 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4950
            TabIndex        =   51
            Top             =   465
            Width           =   1590
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Valores"
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
            Left            =   5325
            TabIndex        =   50
            Top             =   210
            Width           =   690
         End
         Begin VB.Label ValorBenef 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4950
            TabIndex        =   49
            Top             =   1635
            Width           =   1590
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Em Beneficiamento:"
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
            TabIndex        =   48
            Top             =   1635
            Width           =   1695
         End
         Begin VB.Label QuantBenef 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3165
            TabIndex        =   47
            Top             =   1635
            Width           =   1590
         End
      End
      Begin VB.Label Label24 
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
         Height          =   165
         Left            =   615
         TabIndex        =   87
         Top             =   165
         Width           =   735
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Almoxarifado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   4830
         TabIndex        =   86
         Top             =   165
         Width           =   1155
      End
      Begin VB.Label UnidMedLabel 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   3612
         TabIndex        =   85
         Top             =   150
         Width           =   780
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "U.M.:"
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
         Left            =   3105
         TabIndex        =   84
         Top             =   180
         Width           =   480
      End
      Begin VB.Label ProdutoLabel 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   1404
         TabIndex        =   83
         Top             =   150
         Width           =   1260
      End
      Begin VB.Label AlmoxarifadoLabel 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   6015
         TabIndex        =   82
         Top             =   150
         Width           =   2595
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6885
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   135
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "EstoqueProduto.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "EstoqueProduto.ctx":015E
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "EstoqueProduto.ctx":02E8
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "EstoqueProduto.ctx":081A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5130
      Left            =   135
      TabIndex        =   88
      Top             =   540
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   9049
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Principais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Saldos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Saldos Em/De Terceiros"
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
Attribute VB_Name = "EstoqueProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTEstoqueProduto
Attribute objCT.VB_VarHelpID = -1

Private Sub Almoxarifado_Change()
     Call objCT.Almoxarifado_Change
End Sub

Private Sub Almoxarifado_Click()
     Call objCT.Almoxarifado_Click
End Sub

Private Sub BotaoExcluir_Click()
     Call objCT.BotaoExcluir_Click
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub BotaoGravar_Click()
     Call objCT.BotaoGravar_Click
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Public Sub Form_Activate()
     Call objCT.Form_Activate
End Sub

Public Sub Form_Deactivate()
     Call objCT.Form_Deactivate
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub LocalizacaoFisica_Change()
     Call objCT.LocalizacaoFisica_Change
End Sub

Private Sub ContaContabil_Change()
     Call objCT.ContaContabil_Change
End Sub

Private Sub ContaContabil_Validate(Cancel As Boolean)
     Call objCT.ContaContabil_Validate(Cancel)
End Sub

Private Sub ContaContabilLabel_Click()
     Call objCT.ContaContabilLabel_Click
End Sub

Private Sub Padrao_Click()
     Call objCT.Padrao_Click
End Sub

Private Sub ProdutoLabel1_Click()
     Call objCT.ProdutoLabel1_Click
End Sub

Private Sub Produto_Change()
     Call objCT.Produto_Change
End Sub

Private Sub Produto_Validate(Cancel As Boolean)
     Call objCT.Produto_Validate(Cancel)
End Sub

Private Sub TabStrip1_Click()
     Call objCT.TabStrip1_Click
End Sub

'Private Sub TvwProduto_Expand(ByVal Node As MSComctlLib.Node)
'     Call objCT.TvwProduto_Expand(Node)
'End Sub
'
'Private Sub TvwProduto_NodeClick(ByVal Node As MSComctlLib.Node)
'     Call objCT.TvwProduto_NodeClick(Node)
'End Sub

Function Trata_Parametros(Optional objEstoqueProduto As ClassEstoqueProduto) As Long
     Trata_Parametros = objCT.Trata_Parametros(objEstoqueProduto)
End Function

Public Function Form_Load_Ocx() As Object

    Call objCT.Form_Load_Ocx
    Set Form_Load_Ocx = Me

End Function

Public Sub Form_Unload(Cancel As Integer)
    If Not (objCT Is Nothing) Then
        Call objCT.Form_Unload(Cancel)
        Set objCT.objUserControl = Nothing
        Set objCT = Nothing
    End If
End Sub

Private Sub objCT_Unload()
   RaiseEvent Unload
End Sub

Public Function Name() As String
    Name = objCT.Name
End Function

Public Sub Show()
    Call objCT.Show
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

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property

Private Sub UserControl_Initialize()
    Set objCT = New CTEstoqueProduto
    Set objCT.objUserControl = Me
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Property Get Caption() As String
    Caption = objCT.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    objCT.Caption = New_Caption
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Call objCT.UserControl_KeyDown(KeyCode, Shift)
End Sub

Private Sub CustoLabel_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(CustoLabel(Index), Source, X, Y)
End Sub

Private Sub CustoLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CustoLabel(Index), Button, Shift, X, Y)
End Sub

Private Sub UnidMedLabel_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(UnidMedLabel(Index), Source, X, Y)
End Sub

Private Sub UnidMedLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(UnidMedLabel(Index), Button, Shift, X, Y)
End Sub

Private Sub ProdutoLabel_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(ProdutoLabel(Index), Source, X, Y)
End Sub

Private Sub ProdutoLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ProdutoLabel(Index), Button, Shift, X, Y)
End Sub


Private Sub AlmoxarifadoLabel_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(AlmoxarifadoLabel(Index), Source, X, Y)
End Sub

Private Sub AlmoxarifadoLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(AlmoxarifadoLabel(Index), Button, Shift, X, Y)
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub

Private Sub UnidMed_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(UnidMed, Source, X, Y)
End Sub

Private Sub UnidMed_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(UnidMed, Button, Shift, X, Y)
End Sub

'Private Sub LabelProduto_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(LabelProduto, Source, X, Y)
'End Sub
'
'Private Sub LabelProduto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(LabelProduto, Button, Shift, X, Y)
'End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Descricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Descricao, Source, X, Y)
End Sub

Private Sub Descricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Descricao, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub ProdutoLabel1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ProdutoLabel1, Source, X, Y)
End Sub

Private Sub ProdutoLabel1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ProdutoLabel1, Button, Shift, X, Y)
End Sub

Private Sub ContaContabilLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContaContabilLabel, Source, X, Y)
End Sub

Private Sub ContaContabilLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContaContabilLabel, Button, Shift, X, Y)
End Sub

Private Sub ValorTotal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorTotal, Source, X, Y)
End Sub

Private Sub ValorTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorTotal, Button, Shift, X, Y)
End Sub

Private Sub Label37_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label37, Source, X, Y)
End Sub

Private Sub Label37_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label37, Button, Shift, X, Y)
End Sub

Private Sub QuantTotal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantTotal, Source, X, Y)
End Sub

Private Sub QuantTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantTotal, Button, Shift, X, Y)
End Sub

Private Sub QuantPedido_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantPedido, Source, X, Y)
End Sub

Private Sub QuantPedido_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantPedido, Button, Shift, X, Y)
End Sub

Private Sub QuantPedidoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantPedidoLabel, Source, X, Y)
End Sub

Private Sub QuantPedidoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantPedidoLabel, Button, Shift, X, Y)
End Sub

Private Sub LabelQuantReqCompra_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelQuantReqCompra, Source, X, Y)
End Sub

Private Sub LabelQuantReqCompra_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelQuantReqCompra, Button, Shift, X, Y)
End Sub

Private Sub LabelValReqCompra_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelValReqCompra, Source, X, Y)
End Sub

Private Sub LabelValReqCompra_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelValReqCompra, Button, Shift, X, Y)
End Sub

Private Sub LabelValPedCompra_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelValPedCompra, Source, X, Y)
End Sub

Private Sub LabelValPedCompra_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelValPedCompra, Button, Shift, X, Y)
End Sub

Private Sub LabelQuantPedCompra_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelQuantPedCompra, Source, X, Y)
End Sub

Private Sub LabelQuantPedCompra_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelQuantPedCompra, Button, Shift, X, Y)
End Sub

Private Sub ValorPedido_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorPedido, Source, X, Y)
End Sub

Private Sub ValorPedido_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorPedido, Button, Shift, X, Y)
End Sub

Private Sub QuantInd_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantInd, Source, X, Y)
End Sub

Private Sub QuantInd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantInd, Button, Shift, X, Y)
End Sub

Private Sub Label16_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label16, Source, X, Y)
End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label16, Button, Shift, X, Y)
End Sub

Private Sub QuantDefeituosa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantDefeituosa, Source, X, Y)
End Sub

Private Sub QuantDefeituosa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantDefeituosa, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub QuantEmpenhada_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantEmpenhada, Source, X, Y)
End Sub

Private Sub QuantEmpenhada_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantEmpenhada, Button, Shift, X, Y)
End Sub

Private Sub Label20_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label20, Source, X, Y)
End Sub

Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label20, Button, Shift, X, Y)
End Sub

Private Sub QuantReservada_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantReservada, Source, X, Y)
End Sub

Private Sub QuantReservada_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantReservada, Button, Shift, X, Y)
End Sub

Private Sub Label21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label21, Source, X, Y)
End Sub

Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label21, Button, Shift, X, Y)
End Sub

Private Sub QuantDispTotal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantDispTotal, Source, X, Y)
End Sub

Private Sub QuantDispTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantDispTotal, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label32_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label32, Source, X, Y)
End Sub

Private Sub Label32_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label32, Button, Shift, X, Y)
End Sub

Private Sub ValorInd_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorInd, Source, X, Y)
End Sub

Private Sub ValorInd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorInd, Button, Shift, X, Y)
End Sub

Private Sub ValorDefeituosa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorDefeituosa, Source, X, Y)
End Sub

Private Sub ValorDefeituosa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorDefeituosa, Button, Shift, X, Y)
End Sub

Private Sub ValorEmpenhada_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorEmpenhada, Source, X, Y)
End Sub

Private Sub ValorEmpenhada_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorEmpenhada, Button, Shift, X, Y)
End Sub

Private Sub ValorReservada_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorReservada, Source, X, Y)
End Sub

Private Sub ValorReservada_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorReservada, Button, Shift, X, Y)
End Sub

Private Sub ValorDispTotal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorDispTotal, Source, X, Y)
End Sub

Private Sub ValorDispTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorDispTotal, Button, Shift, X, Y)
End Sub

Private Sub Label33_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label33, Source, X, Y)
End Sub

Private Sub Label33_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label33, Button, Shift, X, Y)
End Sub

Private Sub ValorOP_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorOP, Source, X, Y)
End Sub

Private Sub ValorOP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorOP, Button, Shift, X, Y)
End Sub

Private Sub Label36_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label36, Source, X, Y)
End Sub

Private Sub Label36_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label36, Button, Shift, X, Y)
End Sub

Private Sub QuantOP_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantOP, Source, X, Y)
End Sub

Private Sub QuantOP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantOP, Button, Shift, X, Y)
End Sub

Private Sub Label50_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label50, Source, X, Y)
End Sub

Private Sub Label50_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label50, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub DataInventario_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataInventario, Source, X, Y)
End Sub

Private Sub DataInventario_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataInventario, Button, Shift, X, Y)
End Sub

Private Sub Label27_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label27, Source, X, Y)
End Sub

Private Sub Label27_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label27, Button, Shift, X, Y)
End Sub

Private Sub Label30_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label30, Source, X, Y)
End Sub

Private Sub Label30_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label30, Button, Shift, X, Y)
End Sub

Private Sub Label31_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label31, Source, X, Y)
End Sub

Private Sub Label31_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label31, Button, Shift, X, Y)
End Sub

Private Sub QuantOutras3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantOutras3, Source, X, Y)
End Sub

Private Sub QuantOutras3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantOutras3, Button, Shift, X, Y)
End Sub

Private Sub QuantDemo3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantDemo3, Source, X, Y)
End Sub

Private Sub QuantDemo3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantDemo3, Button, Shift, X, Y)
End Sub

Private Sub QuantConsig3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantConsig3, Source, X, Y)
End Sub

Private Sub QuantConsig3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantConsig3, Button, Shift, X, Y)
End Sub

Private Sub QuantConserto3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantConserto3, Source, X, Y)
End Sub

Private Sub QuantConserto3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantConserto3, Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label15, Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15, Button, Shift, X, Y)
End Sub

Private Sub Label17_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label17, Source, X, Y)
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label17, Button, Shift, X, Y)
End Sub

Private Sub Label18_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label18, Source, X, Y)
End Sub

Private Sub Label18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label18, Button, Shift, X, Y)
End Sub

Private Sub Label19_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label19, Source, X, Y)
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label19, Button, Shift, X, Y)
End Sub

Private Sub Label22_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label22, Source, X, Y)
End Sub

Private Sub Label22_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label22, Button, Shift, X, Y)
End Sub

Private Sub ValorOutras3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorOutras3, Source, X, Y)
End Sub

Private Sub ValorOutras3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorOutras3, Button, Shift, X, Y)
End Sub

Private Sub ValorDemo3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorDemo3, Source, X, Y)
End Sub

Private Sub ValorDemo3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorDemo3, Button, Shift, X, Y)
End Sub

Private Sub ValorConsig3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorConsig3, Source, X, Y)
End Sub

Private Sub ValorConsig3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorConsig3, Button, Shift, X, Y)
End Sub

Private Sub ValorConserto3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorConserto3, Source, X, Y)
End Sub

Private Sub ValorConserto3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorConserto3, Button, Shift, X, Y)
End Sub

Private Sub Label23_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label23, Source, X, Y)
End Sub

Private Sub Label23_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label23, Button, Shift, X, Y)
End Sub

Private Sub ValorBenef3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorBenef3, Source, X, Y)
End Sub

Private Sub ValorBenef3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorBenef3, Button, Shift, X, Y)
End Sub

Private Sub Label35_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label35, Source, X, Y)
End Sub

Private Sub Label35_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label35, Button, Shift, X, Y)
End Sub

Private Sub QuantBenef3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantBenef3, Source, X, Y)
End Sub

Private Sub QuantBenef3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantBenef3, Button, Shift, X, Y)
End Sub

Private Sub QuantOutras_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantOutras, Source, X, Y)
End Sub

Private Sub QuantOutras_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantOutras, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub QuantDemo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantDemo, Source, X, Y)
End Sub

Private Sub QuantDemo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantDemo, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub QuantConsig_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantConsig, Source, X, Y)
End Sub

Private Sub QuantConsig_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantConsig, Button, Shift, X, Y)
End Sub

Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub

Private Sub QuantConserto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantConserto, Source, X, Y)
End Sub

Private Sub QuantConserto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantConserto, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub ValorOutras_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorOutras, Source, X, Y)
End Sub

Private Sub ValorOutras_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorOutras, Button, Shift, X, Y)
End Sub

Private Sub ValorDemo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorDemo, Source, X, Y)
End Sub

Private Sub ValorDemo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorDemo, Button, Shift, X, Y)
End Sub

Private Sub ValorConsig_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorConsig, Source, X, Y)
End Sub

Private Sub ValorConsig_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorConsig, Button, Shift, X, Y)
End Sub

Private Sub ValorConserto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorConserto, Source, X, Y)
End Sub

Private Sub ValorConserto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorConserto, Button, Shift, X, Y)
End Sub

Private Sub Label26_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label26, Source, X, Y)
End Sub

Private Sub Label26_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label26, Button, Shift, X, Y)
End Sub

Private Sub ValorBenef_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorBenef, Source, X, Y)
End Sub

Private Sub ValorBenef_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorBenef, Button, Shift, X, Y)
End Sub

Private Sub Label34_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label34, Source, X, Y)
End Sub

Private Sub Label34_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label34, Button, Shift, X, Y)
End Sub

Private Sub QuantBenef_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantBenef, Source, X, Y)
End Sub

Private Sub QuantBenef_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantBenef, Button, Shift, X, Y)
End Sub

Private Sub Label24_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label24, Source, X, Y)
End Sub

Private Sub Label24_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label24, Button, Shift, X, Y)
End Sub

Private Sub Label25_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label25, Source, X, Y)
End Sub

Private Sub Label25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label25, Button, Shift, X, Y)
End Sub

Private Sub Label28_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label28, Source, X, Y)
End Sub

Private Sub Label28_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label28, Button, Shift, X, Y)
End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub

Private Sub CustoConserto3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CustoConserto3, Source, X, Y)
End Sub

Private Sub CustoConserto3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CustoConserto3, Button, Shift, X, Y)
End Sub

Private Sub CustoConsig3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CustoConsig3, Source, X, Y)
End Sub

Private Sub CustoConsig3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CustoConsig3, Button, Shift, X, Y)
End Sub

Private Sub CustoDemo3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CustoDemo3, Source, X, Y)
End Sub

Private Sub CustoDemo3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CustoDemo3, Button, Shift, X, Y)
End Sub

Private Sub CustoOutras3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CustoOutras3, Source, X, Y)
End Sub

Private Sub CustoOutras3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CustoOutras3, Button, Shift, X, Y)
End Sub

Private Sub CustoBenef3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CustoBenef3, Source, X, Y)
End Sub

Private Sub CustoBenef3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CustoBenef3, Button, Shift, X, Y)
End Sub

Private Sub CustoConserto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CustoConserto, Source, X, Y)
End Sub

Private Sub CustoConserto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CustoConserto, Button, Shift, X, Y)
End Sub

Private Sub CustoConsig_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CustoConsig, Source, X, Y)
End Sub

Private Sub CustoConsig_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CustoConsig, Button, Shift, X, Y)
End Sub

Private Sub CustoDemo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CustoDemo, Source, X, Y)
End Sub

Private Sub CustoDemo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CustoDemo, Button, Shift, X, Y)
End Sub

Private Sub CustoOutras_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CustoOutras, Source, X, Y)
End Sub

Private Sub CustoOutras_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CustoOutras, Button, Shift, X, Y)
End Sub

Private Sub CustoBenef_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CustoBenef, Source, X, Y)
End Sub

Private Sub CustoBenef_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CustoBenef, Button, Shift, X, Y)
End Sub

Private Sub Label41_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label41, Source, X, Y)
End Sub

Private Sub Label41_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label41, Button, Shift, X, Y)
End Sub

Private Sub Label42_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label42, Source, X, Y)
End Sub

Private Sub Label42_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label42, Button, Shift, X, Y)
End Sub

