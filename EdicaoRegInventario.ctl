VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.UserControl EdicaoRegInventarioOcx 
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8925
   KeyPreview      =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   8925
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4605
      Index           =   1
      Left            =   90
      TabIndex        =   25
      Top             =   870
      Width           =   8730
      Begin VB.CommandButton BotaoExcluirRegInv 
         Caption         =   "Excluir Registro de Inventário"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3720
         TabIndex        =   3
         Top             =   1470
         Width           =   4065
      End
      Begin VB.CommandButton BotaoRegCadastrado 
         Caption         =   "Registros de Inventário Cadastrados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3720
         TabIndex        =   2
         Top             =   844
         Width           =   4065
      End
      Begin VB.CommandButton BotaoGerar 
         Caption         =   "Gerar Registro de Inventário"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3720
         TabIndex        =   1
         Top             =   240
         Width           =   4065
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Left            =   2475
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   300
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   300
         Left            =   1380
         TabIndex        =   0
         Top             =   300
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
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
         Left            =   870
         TabIndex        =   26
         Top             =   330
         Width           =   480
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Caption         =   "Frame22"
      Height          =   3870
      Index           =   2
      Left            =   780
      TabIndex        =   24
      Top             =   930
      Visible         =   0   'False
      Width           =   7290
      Begin VB.Frame Frame3 
         Caption         =   "Almoxarifado"
         Height          =   1005
         Left            =   210
         TabIndex        =   44
         Top             =   1620
         Width           =   6765
         Begin VB.ComboBox Almoxarifado 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   540
            Width           =   3825
         End
         Begin VB.OptionButton OptionTodosTipos 
            Caption         =   "Todos"
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
            Left            =   180
            TabIndex        =   6
            Top             =   285
            Value           =   -1  'True
            Width           =   960
         End
         Begin VB.OptionButton OptionUmTipo 
            Caption         =   "Apenas do Almoxarifado:"
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
            Left            =   180
            TabIndex        =   7
            Top             =   600
            Width           =   2445
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Conta Contabil"
         Height          =   1005
         Left            =   210
         TabIndex        =   43
         Top             =   2730
         Width           =   6765
         Begin VB.CommandButton BotaoContaContabil 
            Caption         =   "Conta Contabil"
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
            Left            =   4770
            TabIndex        =   12
            Top             =   518
            Width           =   1605
         End
         Begin VB.OptionButton OptionContaTodas 
            Caption         =   "Todas"
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
            TabIndex        =   9
            Top             =   315
            Value           =   -1  'True
            Width           =   960
         End
         Begin VB.OptionButton OptionContaUma 
            Caption         =   "Apenas da Conta Contabil:"
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
            TabIndex        =   10
            Top             =   600
            Width           =   2625
         End
         Begin MSMask.MaskEdBox ContaContabil 
            Height          =   300
            Left            =   2790
            TabIndex        =   11
            Top             =   540
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   529
            _Version        =   393216
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
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
            PromptChar      =   " "
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Produtos"
         Height          =   1395
         Index           =   0
         Left            =   225
         TabIndex        =   29
         Top             =   120
         Width           =   6765
         Begin MSMask.MaskEdBox ProdutoInicial 
            Height          =   315
            Left            =   705
            TabIndex        =   4
            Top             =   360
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdutoFinal 
            Height          =   315
            Left            =   705
            TabIndex        =   5
            Top             =   840
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label DescProdFim 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2280
            TabIndex        =   33
            Top             =   840
            Width           =   4065
         End
         Begin VB.Label DescProdInic 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2280
            TabIndex        =   32
            Top             =   360
            Width           =   4065
         End
         Begin VB.Label ProdutoInicialLabel 
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
            Height          =   255
            Left            =   345
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   31
            Top             =   405
            Width           =   360
         End
         Begin VB.Label ProdutoFinalLabel 
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
            Height          =   255
            Left            =   315
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   30
            Top             =   885
            Width           =   435
         End
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4575
      Index           =   3
      Left            =   75
      TabIndex        =   23
      Top             =   870
      Visible         =   0   'False
      Width           =   8760
      Begin VB.ComboBox NaturezaProduto 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "EdicaoRegInventario.ctx":0000
         Left            =   2400
         List            =   "EdicaoRegInventario.ctx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1731
         Width           =   4200
      End
      Begin VB.TextBox Observacoes 
         Height          =   315
         Left            =   2220
         TabIndex        =   22
         Top             =   4020
         Width           =   4005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Valor :"
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
         Left            =   1800
         TabIndex        =   103
         Top             =   3619
         Width           =   570
      End
      Begin VB.Label ValorTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2400
         TabIndex        =   102
         Top             =   3559
         Width           =   1365
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
         Height          =   195
         Left            =   3690
         TabIndex        =   46
         Top             =   420
         Width           =   1155
      End
      Begin VB.Label AlmoxarifadoLabel 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   4890
         TabIndex        =   45
         Top             =   360
         Width           =   1710
      End
      Begin VB.Label ValorUnitario 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2400
         TabIndex        =   21
         Top             =   3102
         Width           =   1365
      End
      Begin VB.Label Quantidade 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2400
         TabIndex        =   18
         Top             =   2645
         Width           =   1200
      End
      Begin VB.Label IPICodigo 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5250
         TabIndex        =   16
         Top             =   1274
         Width           =   1350
      End
      Begin VB.Label Modelo 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2400
         TabIndex        =   15
         Top             =   1274
         Width           =   1635
      End
      Begin VB.Label Descricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2400
         TabIndex        =   14
         Top             =   817
         Width           =   4200
      End
      Begin VB.Label Produto 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2400
         TabIndex        =   13
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label UM 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2400
         TabIndex        =   20
         Top             =   2188
         Width           =   735
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "Observações:"
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
         Left            =   960
         TabIndex        =   42
         Top             =   4065
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade:"
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
         Left            =   1275
         TabIndex        =   41
         Top             =   2705
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Custo Unitário:"
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
         Left            =   1095
         TabIndex        =   40
         Top             =   3162
         Width           =   1275
      End
      Begin VB.Label LblUMEstoque 
         AutoSize        =   -1  'True
         Caption         =   "UM:"
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
         TabIndex        =   39
         Top             =   2248
         Width           =   360
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
         Height          =   195
         Index           =   2
         Left            =   1395
         TabIndex        =   38
         Top             =   877
         Width           =   930
      End
      Begin VB.Label ProdutoLabel 
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
         Height          =   195
         Left            =   1590
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   37
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "IPI Codigo:"
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
         Left            =   4290
         TabIndex        =   36
         Top             =   1334
         Width           =   960
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Modelo:"
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
         Left            =   1635
         TabIndex        =   35
         Top             =   1334
         Width           =   690
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
         Height          =   195
         Index           =   1
         Left            =   1485
         TabIndex        =   34
         Top             =   1791
         Width           =   840
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   4575
      Index           =   4
      Left            =   90
      TabIndex        =   47
      Top             =   870
      Visible         =   0   'False
      Width           =   8760
      Begin VB.Frame Frame5 
         Caption         =   "Saldos Nosso em Poder de Terceiros"
         Height          =   1965
         Left            =   165
         TabIndex        =   72
         Top             =   510
         Width           =   8475
         Begin VB.Label QuantBenef 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3165
            TabIndex        =   95
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
            TabIndex        =   94
            Top             =   1635
            Width           =   1695
         End
         Begin VB.Label ValorBenef 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4950
            TabIndex        =   93
            Top             =   1635
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
            TabIndex        =   92
            Top             =   210
            Width           =   690
         End
         Begin VB.Label ValorConserto 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4950
            TabIndex        =   91
            Top             =   465
            Width           =   1590
         End
         Begin VB.Label ValorConsig 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4950
            TabIndex        =   90
            Top             =   750
            Width           =   1590
         End
         Begin VB.Label ValorDemo 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4950
            TabIndex        =   89
            Top             =   1035
            Width           =   1590
         End
         Begin VB.Label ValorOutras 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4950
            TabIndex        =   88
            Top             =   1335
            Width           =   1590
         End
         Begin VB.Label Label7 
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
            TabIndex        =   87
            Top             =   210
            Width           =   1080
         End
         Begin VB.Label Label4 
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
            TabIndex        =   86
            Top             =   495
            Width           =   1140
         End
         Begin VB.Label QuantConserto 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3165
            TabIndex        =   85
            Top             =   480
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
            TabIndex        =   84
            Top             =   780
            Width           =   1485
         End
         Begin VB.Label QuantConsig 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3165
            TabIndex        =   83
            Top             =   750
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
            TabIndex        =   82
            Top             =   1050
            Width           =   1605
         End
         Begin VB.Label QuantDemo 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3165
            TabIndex        =   81
            Top             =   1035
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
            TabIndex        =   80
            Top             =   1335
            Width           =   630
         End
         Begin VB.Label QuantOutras 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3165
            TabIndex        =   79
            Top             =   1335
            Width           =   1590
         End
         Begin VB.Label CustoOutras 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6705
            TabIndex        =   78
            Top             =   1335
            Width           =   1590
         End
         Begin VB.Label CustoDemo 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6705
            TabIndex        =   77
            Top             =   1035
            Width           =   1590
         End
         Begin VB.Label CustoConsig 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6705
            TabIndex        =   76
            Top             =   750
            Width           =   1590
         End
         Begin VB.Label CustoConserto 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6705
            TabIndex        =   75
            Top             =   465
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
            TabIndex        =   74
            Top             =   225
            Width           =   585
         End
         Begin VB.Label CustoBenef 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6705
            TabIndex        =   73
            Top             =   1635
            Width           =   1590
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Saldos de Terceiros em Nosso Poder"
         Height          =   1965
         Left            =   165
         TabIndex        =   48
         Top             =   2550
         Width           =   8475
         Begin VB.Label QuantBenef3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3135
            TabIndex        =   71
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
            TabIndex        =   70
            Top             =   1620
            Width           =   1695
         End
         Begin VB.Label ValorBenef3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4935
            TabIndex        =   69
            Top             =   1620
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
         Begin VB.Label ValorConserto3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4920
            TabIndex        =   67
            Top             =   480
            Width           =   1590
         End
         Begin VB.Label ValorConsig3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4920
            TabIndex        =   66
            Top             =   780
            Width           =   1590
         End
         Begin VB.Label ValorDemo3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4920
            TabIndex        =   65
            Top             =   1050
            Width           =   1590
         End
         Begin VB.Label ValorOutras3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4920
            TabIndex        =   64
            Top             =   1335
            Width           =   1590
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
            TabIndex        =   63
            Top             =   225
            Width           =   1080
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
            TabIndex        =   62
            Top             =   480
            Width           =   1140
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
            TabIndex        =   61
            Top             =   780
            Width           =   1485
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
            TabIndex        =   60
            Top             =   1065
            Width           =   1605
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
            TabIndex        =   59
            Top             =   1335
            Width           =   630
         End
         Begin VB.Label QuantConserto3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3135
            TabIndex        =   58
            Top             =   480
            Width           =   1590
         End
         Begin VB.Label QuantConsig3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3135
            TabIndex        =   57
            Top             =   765
            Width           =   1590
         End
         Begin VB.Label QuantDemo3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3135
            TabIndex        =   56
            Top             =   1050
            Width           =   1590
         End
         Begin VB.Label QuantOutras3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3135
            TabIndex        =   55
            Top             =   1335
            Width           =   1590
         End
         Begin VB.Label CustoOutras3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6675
            TabIndex        =   54
            Top             =   1335
            Width           =   1590
         End
         Begin VB.Label CustoDemo3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6675
            TabIndex        =   53
            Top             =   1050
            Width           =   1590
         End
         Begin VB.Label CustoConsig3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6675
            TabIndex        =   52
            Top             =   765
            Width           =   1590
         End
         Begin VB.Label CustoConserto3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6675
            TabIndex        =   51
            Top             =   480
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
            TabIndex        =   50
            Top             =   195
            Width           =   585
         End
         Begin VB.Label CustoBenef3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6675
            TabIndex        =   49
            Top             =   1620
            Width           =   1590
         End
      End
      Begin VB.Label AlmoxarifadoCaption 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6465
         TabIndex        =   101
         Top             =   150
         Width           =   1785
      End
      Begin VB.Label ProdutoTerc 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1380
         TabIndex        =   100
         Top             =   150
         Width           =   1260
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
         Left            =   3300
         TabIndex        =   99
         Top             =   165
         Width           =   480
      End
      Begin VB.Label UnidMedLabel 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3825
         TabIndex        =   98
         Top             =   150
         Width           =   780
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
         Left            =   5265
         TabIndex        =   97
         Top             =   165
         Width           =   1155
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
         Left            =   585
         TabIndex        =   96
         Top             =   165
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6690
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   135
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "EdicaoRegInventario.ctx":00A4
         Style           =   1  'Graphical
         TabIndex        =   107
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   630
         Picture         =   "EdicaoRegInventario.ctx":01FE
         Style           =   1  'Graphical
         TabIndex        =   106
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1125
         Picture         =   "EdicaoRegInventario.ctx":0388
         Style           =   1  'Graphical
         TabIndex        =   105
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1635
         Picture         =   "EdicaoRegInventario.ctx":08BA
         Style           =   1  'Graphical
         TabIndex        =   104
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4965
      Left            =   60
      TabIndex        =   19
      Top             =   540
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   8758
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Filtro de Produtos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Edição"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
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
Attribute VB_Name = "EdicaoRegInventarioOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Variáveis Globais
Dim iAlterado As Integer
Dim iFrameAtual As Integer
Dim iFrameFiltroAlterado As Integer

'Eventos dos Browses
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1
Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1
Private WithEvents objEventoBotaoInventario As AdmEvento
Attribute objEventoBotaoInventario.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_IDENTIFICACAO = 1
Private Const TAB_FILTRO = 2
Private Const TAB_EDICAO = 3
Private Const TAB_TERCEIROS = 4

Function Trata_Parametros(Optional objRegInventario As ClassRegInventario) As Long

Dim lErro As Long
Dim objRegInventarioAlmox As New ClassRegInventarioAlmox

On Error GoTo Erro_Trata_Parametros

    'Se foi passada um Registro de Inventário como parâmetro
    If Not objRegInventario Is Nothing Then

        'Se a Data e o Produto vieram preenchidos
        If objRegInventario.dtData <> DATA_NULA And Len(Trim(objRegInventario.sProduto)) > 0 Then

            'Guarda a Filial Empresa
            objRegInventario.iFilialEmpresa = giFilialEmpresa

            'Lê o Registro de Inventário passado
            lErro = CF("RegInventario_Le", objRegInventario)
            If lErro <> SUCESSO And lErro <> 70308 Then gError 70343

            'Se encontrou o Registro de Inventário
            If lErro = SUCESSO Then

                'Traz RegInventario para a tela
                lErro = Traz_RegInventario_Tela(objRegInventario, objRegInventarioAlmox)
                If lErro <> SUCESSO Then gError 70543

            End If

        'Se apenas a data veio preenchida
        ElseIf objRegInventario.dtData <> DATA_NULA And Len(Trim(objRegInventario.sProduto)) = 0 Then

            objRegInventario.iFilialEmpresa = giFilialEmpresa

            'Lê o Registro de Inventário apenas com a data passada
            lErro = CF("RegInventario_Le_Data", objRegInventario)
            If lErro <> SUCESSO And lErro <> 70237 Then gError 70544

            'Coloca a data do Registro de inventário na tela
            Call DateParaMasked(Data, objRegInventario.dtData)
        End If

    End If

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 70343, 70543, 70544

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159271)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iFrameAtual = 1

    'Eventos dos Browses
    Set objEventoProduto = New AdmEvento
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    Set objEventoContaContabil = New AdmEvento
    Set objEventoBotaoInventario = New AdmEvento

    'Carrega Almoxarifados cadastrados
    lErro = Almoxarifados_Carrega()
    If lErro <> SUCESSO Then gError 70278

    'Mascara o Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoInicial)
    If lErro <> SUCESSO Then gError 70288

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoFinal)
    If lErro <> SUCESSO Then gError 70522

    'Mascara a Conta Contábil
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabil)
    If lErro <> SUCESSO Then gError 70525

    iAlterado = 0
    iFrameFiltroAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 70278, 70288, 70522, 70525

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159272)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Almoxarifados_Carrega() As Long

Dim lErro As Long
Dim colAlmoxarifados As New Collection
Dim objAlmoxarifado As ClassAlmoxarifado

On Error GoTo Erro_Almoxarifados_Carrega

    'Lê Códigos e NomesReduzidos da tabela Almoxarifado e devolve na coleção
    lErro = CF("Almoxarifados_Le_FilialEmpresa", giFilialEmpresa, colAlmoxarifados)
    If lErro <> SUCESSO Then gError 70279

    'Preenche a ListBox AlmoxarifadoList com os objetos da coleção
    For Each objAlmoxarifado In colAlmoxarifados
        Almoxarifado.AddItem objAlmoxarifado.sNomeReduzido
        Almoxarifado.ItemData(Almoxarifado.NewIndex) = objAlmoxarifado.iCodigo
    Next

    Almoxarifados_Carrega = SUCESSO

    Exit Function

Erro_Almoxarifados_Carrega:

    Almoxarifados_Carrega = gErr

    Select Case gErr

        Case 70279

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159273)

    End Select

    Exit Function

End Function

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Libera variáveis globais
    Set objEventoProduto = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    Set objEventoContaContabil = Nothing
    Set objEventoBotaoInventario = Nothing

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub Almoxarifado_Click()

    iFrameFiltroAlterado = REGISTRO_ALTERADO

End Sub

Private Sub BotaoExcluirRegInv_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objRegInventario As New ClassRegInventario

On Error GoTo Erro_BotaoExcluirRegInv_Click

    'Verifica se a data foi preenchida
    If Len(Trim(Data.ClipText)) = 0 Then gError 70943

    'Lê o Registro de Inventário a partir da Data e FilialEmpresa
    objRegInventario.dtData = CDate(Data.Text)
    objRegInventario.iFilialEmpresa = giFilialEmpresa
    lErro = CF("RegInventario_Le_Data", objRegInventario)
    If lErro <> SUCESSO And lErro <> 70308 Then gError 70944

    'Se não encontrou, erro
    If lErro = 70309 Then gError 70945

    'Pede a confirmação da exclusão do Registro de Inventário
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_REGIVENTARIOTODOS", objRegInventario.dtData)
    If vbMsgRes = vbNo Then Exit Sub

    GL_objMDIForm.MousePointer = vbHourglass

    'Exclui todos os Registro de Inventário com a data passada
    lErro = CF("RegInventarioTodos_Exclui", objRegInventario)
    If lErro <> SUCESSO Then gError 70946

    'Limpa a tela
    Call Limpa_Tela_RegInventario

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluirRegInv_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 70943
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)

        Case 70944, 70946

        Case 70945
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGINVENTARIO_NAO_CADASTRADO1", gErr, objRegInventario.dtData)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159274)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGerar_Click()

Dim lErro As Long
Dim objRegInventario As New ClassRegInventario
Dim objLivroFilial As New ClassLivrosFilial
Dim sNomeArqParam As String
Dim objEstoqueMes As New ClassEstoqueMes
Dim dtData As Date, ret As VbMsgBoxResult

On Error GoTo Erro_BotaoGerar_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se a data está preenchida
    If Len(Trim(Data.ClipText)) = 0 Then gError 70232

    objRegInventario.iFilialEmpresa = giFilialEmpresa
    objRegInventario.dtData = CDate(Data.Text)

    'Verifica se existe um outro Registro de inventário para essa mesma data
    lErro = CF("RegInventario_Le_Data", objRegInventario)
    If lErro <> SUCESSO And lErro <> 70237 Then gError 70233

    'Se encontrou, erro
    If lErro = SUCESSO Then gError 70238

    'Verifica se o Mês da data passada foi fechado e apurado
    objEstoqueMes.iAno = Year(objRegInventario.dtData)
    objEstoqueMes.iMes = Month(objRegInventario.dtData)
    objEstoqueMes.iFilialEmpresa = giFilialEmpresa
    
    lErro = CF("EstoqueMes_Le", objEstoqueMes)
    If lErro <> SUCESSO And lErro <> 36513 Then gError 70535

    'Se não encontrou o EstoqueMes, erro
    If lErro = 36513 Then gError 70536

    'Se o mês não estiver fechado ou não tiver sido apurado, erro
    If objEstoqueMes.iFechamento <> ESTOQUEMES_FECHAMENTO_FECHADO Or objEstoqueMes.iCustoProdApurado <> CUSTO_APURADO Then
    
        ret = Rotina_Aviso(vbYesNo, "ERRO_ESTOQUEMES_ABERTO_NAOAPURADO", objEstoqueMes.iMes, giFilialEmpresa)
        If ret <> vbYes Then gError 70537
        
    End If

''    lErro = Sistema_Preparar_Batch(sNomeArqParam)
''    If lErro <> SUCESSO Then gError 70240
''
''    'Gera Registro de Inventário
''    lErro = Gerar_RegInventario(sNomeArqParam, objRegInventario.dtData)
''    If lErro <> SUCESSO Then gError 70239

    'Verifica se a data está dentro do intervalo de data do Livro Fiscal de Registro de Inventário
    objLivroFilial.iFilialEmpresa = giFilialEmpresa
    dtData = StrParaDate(Data.Text)
    objLivroFilial.iCodLivro = LIVRO_REG_INVENTARIO_CODIGO
    
''    lErro = CF("RegInventario_IntervaloData_Critica", objLivroFilial, dtData)
''    If lErro <> SUCESSO And lErro <> 70925 Then gError 70926
''
''    'Se não está, erro
''    If lErro = 70925 Then gError 70927

    '??? Apagar a chamada abaixo e descomentar as funções acima quando virar Batch
    lErro = Rotina_Geracao_RegInventario(objRegInventario.dtData)
    If lErro <> SUCESSO Then gError 70239

    Call Limpa_Tela_RegInventario

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoGerar_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 70232
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)

        Case 70238
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGINVENTARIO_EXISTENTE", gErr, objRegInventario.dtData)

        Case 70233, 70239, 70240, 70535, 70537, 70926

        Case 70536
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEMES_INEXISTENTE1", gErr, giFilialEmpresa)

        Case 70927
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_REGINVENTARIO_FORA_INTERVALO", gErr, dtData)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159275)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objRegInventario As New ClassRegInventario
Dim objRegInventarioAlmox As New ClassRegInventarioAlmox

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "RegInventario_RegInvAlmox"

    'Move os dados da tela para a memória
    lErro = Move_Tela_Memoria(objRegInventario)
    If lErro <> SUCESSO Then gError 70263

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Produto", objRegInventario.sProduto, STRING_PRODUTO, "Produto"
    colCampoValor.Add "Descricao", objRegInventario.sDescricao, STRING_DESCRICAO_CAMPO, "Descricao"
    colCampoValor.Add "Modelo", objRegInventario.sModelo, STRING_PRODUTO_MODELO, "Modelo"
    colCampoValor.Add "IPICodigo", objRegInventario.sIPICodigo, STRING_PRODUTO_IPI_CODIGO, "IPICodigo"
    colCampoValor.Add "SiglaUMEstoque", objRegInventario.sSiglaUMEstoque, STRING_UM_SIGLA, "SiglaUMEstoque"
    colCampoValor.Add "QuantidadeUMEstoque", objRegInventario.dQuantidadeUMEstoque, 0, "QuantidadeUMEstoque"
    colCampoValor.Add "ValorUnitario", objRegInventario.dValorUnitario, 0, "ValorUnitario"
    colCampoValor.Add "Natureza", objRegInventario.iNatureza, 0, "Natureza"
    colCampoValor.Add "QtdeNossaEmTerc", objRegInventario.dQtdeNossaEmTerc, 0, "QtdeNossaEmTerc"
    colCampoValor.Add "QtdeDeTercConosco", objRegInventario.dQtdeDeTercConosco, 0, "QtdeDeTercConosco"
    colCampoValor.Add "Observacoes", objRegInventario.sObservacoes, STRING_OBSERVACAO, "Observacoes"
    colCampoValor.Add "ContaContabil", objRegInventario.sContaContabil, STRING_CONTA, "ContaContabil"
    colCampoValor.Add "QuantConsig3", objRegInventario.dQuantConsig3, 0, "QuantConsig3"
    colCampoValor.Add "QuantConsig", objRegInventario.dQuantConsig, 0, "QuantConsig"
    colCampoValor.Add "QuantDemo3", objRegInventario.dQuantDemo3, 0, "QuantDemo3"
    colCampoValor.Add "QuantDemo", objRegInventario.dQuantDemo, 0, "QuantDemo"
    colCampoValor.Add "QuantConserto3", objRegInventario.dQuantConserto3, 0, "QuantConserto3"
    colCampoValor.Add "QuantConserto", objRegInventario.dQuantConserto, 0, "QuantConserto"
    colCampoValor.Add "QuantOutras3", objRegInventario.dQuantOutras3, 0, "QuantOutras3"
    colCampoValor.Add "QuantOutras", objRegInventario.dQuantOutras, 0, "QuantOutras"
    colCampoValor.Add "QuantBenef", objRegInventario.dQuantBenef, 0, "QuantBenef"
    colCampoValor.Add "QuantBenef3", objRegInventario.dQuantBenef3, 0, "QuantBenef3"
    colCampoValor.Add "CustoConsig3", objRegInventario.dCustoConsig3, 0, "CustoConsig3"
    colCampoValor.Add "CustoConsig", objRegInventario.dCustoConsig, 0, "CustoConsig"
    colCampoValor.Add "CustoDemo3", objRegInventario.dCustoDemo3, 0, "CustoDemo3"
    colCampoValor.Add "CustoDemo", objRegInventario.dCustoDemo, 0, "CustoDemo"
    colCampoValor.Add "CustoConserto3", objRegInventario.dCustoConserto3, 0, "CustoConserto3"
    colCampoValor.Add "CustoConserto", objRegInventario.dCustoConserto, 0, "CustoConserto"
    colCampoValor.Add "CustoOutras3", objRegInventario.dCustoOutras3, 0, "CustoOutras3"
    colCampoValor.Add "CustoOutras", objRegInventario.dCustoOutras, 0, "CustoOutras"
    colCampoValor.Add "CustoBenef", objRegInventario.dCustoBenef, 0, "CustoBenef"
    colCampoValor.Add "CustoBenef3", objRegInventario.dCustoBenef3, 0, "CustoBenef3"
    colCampoValor.Add "Almoxarifado", objRegInventarioAlmox.iAlmoxarifado, 0, "Almoxarifado"
    colCampoValor.Add "QuantConsig3Almox", objRegInventarioAlmox.dQuantConsig3, 0, "QuantConsig3Almox"
    colCampoValor.Add "QuantDemo3Almox", objRegInventarioAlmox.dQuantDemo3, 0, "QuantDemo3Almox"
    colCampoValor.Add "QuantConserto3Almox", objRegInventarioAlmox.dQuantConserto3, 0, "QuantConserto3Almox"
    colCampoValor.Add "QuantOutras3Almox", objRegInventarioAlmox.dQuantOutras3, 0, "QuantOutras3Almox"
    colCampoValor.Add "QuantBenef3Almox", objRegInventarioAlmox.dQuantBenef3, 0, "QuantBenef3Almox"
    colCampoValor.Add "QuantidadeUMEstoqueAlmox", objRegInventarioAlmox.dQuantidadeUMEstoque, 0, "QuantidadeUMEstoqueAlmox"
    colCampoValor.Add "QtdeDeTercConoscoAlmox", objRegInventarioAlmox.dQtdeDeTercConosco, 0, "QtdeDeTercConoscoAlmox"

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    'Trata Filtros da Tela
    lErro = Trata_Filtro(colSelecao)
    If lErro <> SUCESSO Then gError 70274

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 70263, 70274

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159276)

    End Select

    Exit Sub

End Sub

Function Trata_Filtro(colSelecao As AdmColFiltro) As Long
'Trata filtro do sistema de setas

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sContaFormatada As String
Dim iContaPreenchida As Integer

On Error GoTo Erro_Trata_Filtro

    'Se a data estiver preenchida
    If Len(Trim(Data.ClipText)) > 0 Then
        colSelecao.Add "Data", OP_IGUAL, CDate(Data.Text)
    End If

    'Se o Produto Inicial foi preenchido
    If Len(Trim(ProdutoInicial.ClipText)) > 0 Then

        'Formata o Produto
        lErro = CF("Produto_Formata", ProdutoInicial.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 70275

        'Adiciona em no Filtro, Produto >= ProdutoInicial
        colSelecao.Add "Produto", OP_MAIOR_OU_IGUAL, sProdutoFormatado

    End If

    'Se o Produto Final foi preenchido
    If Len(Trim(ProdutoFinal.ClipText)) > 0 Then

        'Formata o Produto
        lErro = CF("Produto_Formata", ProdutoFinal.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 70276

        'Adiciona em no Filtro, Produto <= ProdutoFinal
        colSelecao.Add "Produto", OP_MENOR_OU_IGUAL, sProdutoFormatado

    End If

    'Se selecionou apenas um almoxarifado e ele está preenchido
    If OptionUmTipo.Value = True And Almoxarifado.ListIndex <> -1 Then

        'Adiciona no Filtro código do Almoxarifado
        colSelecao.Add "Almoxarifado", OP_IGUAL, Almoxarifado.ItemData(Almoxarifado.ListIndex)

    End If

    'Se selecionou apenas uma Conta Contábil e ela foi preenchida
    If OptionContaUma.Value = True And Len(Trim(ContaContabil.ClipText)) > 0 Then

        'Formata a Conta Contábil
        lErro = CF("Conta_Formata", ContaContabil.Text, sContaFormatada, iContaPreenchida)
        If lErro <> SUCESSO Then gError 70277

        'Adiciona Conta Contábil no Sistema de Setas
        colSelecao.Add "ContaContabil", OP_IGUAL, sContaFormatada

    End If

    Trata_Filtro = SUCESSO

    Exit Function

Erro_Trata_Filtro:

    Trata_Filtro = gErr

    Select Case gErr

        Case 70275, 70276, 70277

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159277)

    End Select

    Exit Function

End Function

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim objRegInventario As New ClassRegInventario
Dim objRegInventarioAlmox As New ClassRegInventarioAlmox
Dim lErro As Long

On Error GoTo Erro_Tela_Preenche

    'Carrega objRegInventario com os dados passados em colCampoValor
    objRegInventario.sProduto = colCampoValor.Item("Produto").vValor
    objRegInventario.iAlmoxarifado = colCampoValor.Item("Almoxarifado").vValor
    objRegInventario.sDescricao = colCampoValor.Item("Descricao").vValor
    objRegInventario.sModelo = colCampoValor.Item("Modelo").vValor
    objRegInventario.sIPICodigo = colCampoValor.Item("IPICodigo").vValor
    objRegInventario.sSiglaUMEstoque = colCampoValor.Item("SiglaUMEstoque").vValor
    objRegInventario.dQuantidadeUMEstoque = colCampoValor.Item("QuantidadeUMEstoque").vValor
    objRegInventario.dValorUnitario = colCampoValor.Item("ValorUnitario").vValor
    objRegInventario.iNatureza = colCampoValor.Item("Natureza").vValor
    objRegInventario.dQtdeNossaEmTerc = colCampoValor.Item("QtdeNossaEmTerc").vValor
    objRegInventario.dQtdeDeTercConosco = colCampoValor.Item("QtdeDeTercConosco").vValor
    objRegInventario.sObservacoes = colCampoValor.Item("Observacoes").vValor
    objRegInventario.sContaContabil = colCampoValor.Item("ContaContabil").vValor
    objRegInventario.dQuantConsig3 = colCampoValor.Item("QuantConsig3").vValor
    objRegInventario.dQuantConsig = colCampoValor.Item("QuantConsig").vValor
    objRegInventario.dQuantDemo3 = colCampoValor.Item("QuantDemo3").vValor
    objRegInventario.dQuantDemo = colCampoValor.Item("QuantDemo").vValor
    objRegInventario.dQuantConserto3 = colCampoValor.Item("QuantConserto3").vValor
    objRegInventario.dQuantConserto = colCampoValor.Item("QuantConserto").vValor
    objRegInventario.dQuantOutras3 = colCampoValor.Item("QuantOutras3").vValor
    objRegInventario.dQuantOutras = colCampoValor.Item("QuantOutras").vValor
    objRegInventario.dQuantBenef = colCampoValor.Item("QuantBenef").vValor
    objRegInventario.dQuantBenef3 = colCampoValor.Item("QuantBenef3").vValor
    objRegInventario.dCustoConsig3 = colCampoValor.Item("CustoConsig3").vValor
    objRegInventario.dCustoConsig = colCampoValor.Item("CustoConsig").vValor
    objRegInventario.dCustoDemo3 = colCampoValor.Item("CustoDemo3").vValor
    objRegInventario.dCustoDemo = colCampoValor.Item("CustoDemo").vValor
    objRegInventario.dCustoConserto3 = colCampoValor.Item("CustoConserto3").vValor
    objRegInventario.dCustoConserto = colCampoValor.Item("CustoConserto").vValor
    objRegInventario.dCustoOutras3 = colCampoValor.Item("CustoOutras3").vValor
    objRegInventario.dCustoOutras = colCampoValor.Item("CustoOutras").vValor
    objRegInventario.dCustoBenef = colCampoValor.Item("CustoBenef").vValor
    objRegInventario.dCustoBenef3 = colCampoValor.Item("CustoBenef3").vValor
    objRegInventarioAlmox.iAlmoxarifado = colCampoValor.Item("Almoxarifado").vValor
    objRegInventarioAlmox.dQuantConsig3 = colCampoValor.Item("QuantConsig3Almox").vValor
    objRegInventarioAlmox.dQuantDemo3 = colCampoValor.Item("QuantDemo3Almox").vValor
    objRegInventarioAlmox.dQuantConserto3 = colCampoValor.Item("QuantConserto3Almox").vValor
    objRegInventarioAlmox.dQuantOutras3 = colCampoValor.Item("QuantOutras3Almox").vValor
    objRegInventarioAlmox.dQuantBenef3 = colCampoValor.Item("QuantBenef3Almox").vValor
    objRegInventarioAlmox.dQuantidadeUMEstoque = colCampoValor.Item("QuantidadeUMEstoqueAlmox").vValor
    objRegInventarioAlmox.dQtdeDeTercConosco = colCampoValor.Item("QtdeDeTercConoscoAlmox").vValor

    'Se a data não foi preenchida, erro
    If Len(Trim(Data.ClipText)) = 0 Then gError 70586

    objRegInventario.dtData = CDate(Data.Text)

    'Se o NumIntDoc estiver preenchido
    If Len(Trim(objRegInventario.sProduto)) > 0 Then

        'Traz dados do Registro de Inventário para a tela
        lErro = Traz_RegInventario_Tela(objRegInventario, objRegInventarioAlmox)
        If lErro <> SUCESSO Then gError 70268

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 70268

        Case 70586
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159278)

    End Select

    Exit Sub

End Sub

Function Traz_RegInventario_Tela(objRegInventario As ClassRegInventario, objRegInventarioAlmox As ClassRegInventarioAlmox) As Long
'Traz os dados do Registro de Inventário para a tela

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoMascarado As String
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_Traz_RegInventario_Tela

    'Coloca Data do Inventário na tela
    If objRegInventario.dtData <> DATA_NULA Then
        Call DateParaMasked(Data, objRegInventario.dtData)
    End If

    'Mascara o Produto
    lErro = Mascara_MascararProduto(objRegInventario.sProduto, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 70269

    'Dados do Produto
    Produto.Caption = sProdutoMascarado
    ProdutoTerc.Caption = sProdutoMascarado
    Descricao.Caption = objRegInventario.sDescricao
    Modelo.Caption = objRegInventario.sModelo
    IPICodigo.Caption = objRegInventario.sIPICodigo

    'Seleciona a Natureza
    For iIndice = 0 To NaturezaProduto.ListCount - 1
        If NaturezaProduto.ItemData(iIndice) = objRegInventario.iNatureza Then
            NaturezaProduto.ListIndex = iIndice
            Exit For
        End If
    Next
    
    UM.Caption = objRegInventario.sSiglaUMEstoque
    UnidMedLabel.Caption = objRegInventario.sSiglaUMEstoque
    Observacoes.Text = objRegInventario.sObservacoes
        
    'Se tem Filtro por Almoxarifado
    If Len(Trim(Almoxarifado.Text)) > 0 Then
        
        'Se possui Almoxarifado
        If objRegInventario.iAlmoxarifado <> 0 Then
            objAlmoxarifado.iCodigo = objRegInventario.iAlmoxarifado
            lErro = CF("Almoxarifado_Le", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then gError 70541
    
            'Almoxarifado não cadastrado
            If lErro = 25056 Then gError 70542
    
            AlmoxarifadoLabel.Caption = objAlmoxarifado.sNomeReduzido
            AlmoxarifadoCaption.Caption = objAlmoxarifado.sNomeReduzido
        End If
            
        'Coloca as Quantidade de Acordo com o que tem no Almoxarifado
        CustoConserto.Caption = ""
        CustoConsig.Caption = ""
        CustoDemo.Caption = ""
        CustoOutras.Caption = ""
        CustoBenef.Caption = ""
            
        If objRegInventarioAlmox.dQuantidadeUMEstoque > QTDE_ESTOQUE_DELTA Then
            ValorUnitario.Caption = Format(objRegInventario.dValorUnitario, FORMATO_CUSTO)
        Else
            ValorUnitario.Caption = Format(0, FORMATO_CUSTO)
        End If
                
        If objRegInventarioAlmox.dQuantConserto3 > QTDE_ESTOQUE_DELTA Then
            CustoConserto3.Caption = Format(objRegInventario.dCustoConserto3, FORMATO_CUSTO)
        Else
            CustoConserto3.Caption = Format(0, FORMATO_CUSTO)
        End If
                
        If objRegInventarioAlmox.dQuantConsig3 > QTDE_ESTOQUE_DELTA Then
            CustoConsig3.Caption = Format(objRegInventario.dCustoConsig3, FORMATO_CUSTO)
        Else
            CustoConsig3.Caption = Format(0, FORMATO_CUSTO)
        End If
                
        If objRegInventarioAlmox.dQuantDemo3 > QTDE_ESTOQUE_DELTA Then
            CustoDemo3.Caption = Format(objRegInventario.dCustoDemo3, FORMATO_CUSTO)
        Else
            CustoDemo3.Caption = Format(0, FORMATO_CUSTO)
        End If
                
        If objRegInventarioAlmox.dQuantOutras3 > QTDE_ESTOQUE_DELTA Then
            CustoOutras3.Caption = Format(objRegInventario.dCustoOutras3, FORMATO_CUSTO)
        Else
            CustoOutras3.Caption = Format(0, FORMATO_CUSTO)
        End If
                
        If objRegInventarioAlmox.dQuantBenef3 > QTDE_ESTOQUE_DELTA Then
           CustoBenef3.Caption = Format(objRegInventario.dCustoBenef3, FORMATO_CUSTO)
        Else
           CustoBenef3.Caption = Format(0, FORMATO_CUSTO)
        End If
        
        Quantidade.Caption = Format(objRegInventarioAlmox.dQuantidadeUMEstoque, "Standard")
        ValorTotal.Caption = Format(objRegInventario.dValorUnitario, "standard")
        
        QuantConserto.Caption = ""
        ValorConserto.Caption = ""
        
        QuantConserto3.Caption = Formata_Estoque(objRegInventarioAlmox.dQuantConserto3)
        ValorConserto3.Caption = Format(objRegInventarioAlmox.dQuantConserto3 * objRegInventario.dCustoConserto3, "standard")
        
        QuantConsig.Caption = ""
        ValorConsig.Caption = ""
        
        QuantConsig3.Caption = Formata_Estoque(objRegInventarioAlmox.dQuantConsig3)
        ValorConsig3.Caption = Format(objRegInventarioAlmox.dQuantConsig3 * objRegInventario.dCustoConsig3, "standard")
        
        QuantDemo.Caption = ""
        ValorDemo.Caption = ""
        
        QuantDemo3.Caption = Formata_Estoque(objRegInventarioAlmox.dQuantDemo3)
        ValorDemo3.Caption = Format(objRegInventarioAlmox.dQuantDemo3 * objRegInventario.dCustoDemo3, "standard")
        
        QuantOutras.Caption = ""
        ValorOutras.Caption = ""
        
        QuantOutras3.Caption = Formata_Estoque(objRegInventarioAlmox.dQuantOutras3)
        ValorOutras3.Caption = Format(objRegInventarioAlmox.dQuantOutras3 * objRegInventario.dCustoOutras3, "standard")
        
        QuantBenef.Caption = ""
        ValorBenef.Caption = ""
        
        QuantBenef3.Caption = Formata_Estoque(objRegInventarioAlmox.dQuantBenef3)
        ValorBenef3.Caption = Format(objRegInventarioAlmox.dQuantBenef3 * objRegInventario.dCustoBenef3, "standard")
    
    Else
        'Se não tem Filtro de Almoxarifado coloca quantidade e Custo da Filial
        If objRegInventario.dQuantidadeUMEstoque > QTDE_ESTOQUE_DELTA Then
            ValorUnitario.Caption = Format(objRegInventario.dValorUnitario / objRegInventario.dQuantidadeUMEstoque, FORMATO_CUSTO)
        Else
            ValorUnitario.Caption = Format(0, FORMATO_CUSTO)
        End If
                
        If objRegInventario.dQuantConserto3 > QTDE_ESTOQUE_DELTA Then
            CustoConserto3.Caption = Format(objRegInventario.dCustoConserto3 / objRegInventario.dQuantConserto3, FORMATO_CUSTO)
        Else
            CustoConserto3.Caption = Format(0, FORMATO_CUSTO)
        End If
                
        If objRegInventario.dQuantConsig3 > QTDE_ESTOQUE_DELTA Then
            CustoConsig3.Caption = Format(objRegInventario.dCustoConsig3 / objRegInventario.dQuantConsig3, FORMATO_CUSTO)
        Else
            CustoConsig3.Caption = Format(0, FORMATO_CUSTO)
        End If
                
        If objRegInventario.dQuantDemo3 > QTDE_ESTOQUE_DELTA Then
            CustoDemo3.Caption = Format(objRegInventario.dCustoDemo3 / objRegInventario.dQuantDemo3, FORMATO_CUSTO)
        Else
            CustoDemo3.Caption = Format(0, FORMATO_CUSTO)
        End If
                
        If objRegInventario.dQuantOutras3 > QTDE_ESTOQUE_DELTA Then
            CustoOutras3.Caption = Format(objRegInventario.dCustoOutras3 / objRegInventario.dQuantOutras3, FORMATO_CUSTO)
        Else
            CustoOutras3.Caption = Format(0, FORMATO_CUSTO)
        End If
                
        If objRegInventario.dQuantBenef3 > QTDE_ESTOQUE_DELTA Then
           CustoBenef3.Caption = Format(objRegInventario.dCustoBenef3 / objRegInventario.dQuantBenef3, FORMATO_CUSTO)
        Else
           CustoBenef3.Caption = Format(0, FORMATO_CUSTO)
        End If
        
        If objRegInventario.dQuantConserto > QTDE_ESTOQUE_DELTA Then
            CustoConserto.Caption = Format(objRegInventario.dCustoConserto / objRegInventario.dQuantConserto, FORMATO_CUSTO)
        Else
            CustoConserto.Caption = Format(0, FORMATO_CUSTO)
        End If
                
        If objRegInventario.dQuantConsig > QTDE_ESTOQUE_DELTA Then
            CustoConsig.Caption = Format(objRegInventario.dCustoConsig / objRegInventario.dQuantConsig, FORMATO_CUSTO)
        Else
            CustoConsig.Caption = Format(0, FORMATO_CUSTO)
        End If
                
        If objRegInventario.dQuantDemo > QTDE_ESTOQUE_DELTA Then
            CustoDemo.Caption = Format(objRegInventario.dCustoDemo / objRegInventario.dQuantDemo, FORMATO_CUSTO)
        Else
            CustoDemo.Caption = Format(0, FORMATO_CUSTO)
        End If
                
        If objRegInventario.dQuantOutras > QTDE_ESTOQUE_DELTA Then
            CustoOutras.Caption = Format(objRegInventario.dCustoOutras / objRegInventario.dQuantOutras, FORMATO_CUSTO)
        Else
            CustoOutras.Caption = Format(0, FORMATO_CUSTO)
        End If
                
        If objRegInventario.dQuantBenef > QTDE_ESTOQUE_DELTA Then
           CustoBenef.Caption = Format(objRegInventario.dCustoBenef / objRegInventario.dQuantBenef, FORMATO_CUSTO)
        Else
           CustoBenef.Caption = Format(0, FORMATO_CUSTO)
        End If
    
        Quantidade.Caption = Format(objRegInventario.dQuantidadeUMEstoque, "Standard")
        ValorTotal.Caption = Format(objRegInventario.dValorUnitario, "standard")
        
        QuantConserto.Caption = Formata_Estoque(objRegInventario.dQuantConserto)
        ValorConserto.Caption = Format(objRegInventario.dCustoConserto, "standard")
        
        QuantConserto3.Caption = Formata_Estoque(objRegInventario.dQuantConserto3)
        ValorConserto3.Caption = Format(objRegInventario.dCustoConserto3, "standard")
        
        QuantConsig.Caption = Formata_Estoque(objRegInventario.dQuantConsig)
        ValorConsig.Caption = Format(objRegInventario.dCustoConsig, "standard")
        
        QuantConsig3.Caption = Formata_Estoque(objRegInventario.dQuantConsig3)
        ValorConsig3.Caption = Format(objRegInventario.dCustoConsig3, "standard")
        
        QuantDemo.Caption = Formata_Estoque(objRegInventario.dQuantDemo)
        ValorDemo.Caption = Format(objRegInventario.dCustoDemo, "standard")
        
        QuantDemo3.Caption = Formata_Estoque(objRegInventario.dQuantDemo3)
        ValorDemo3.Caption = Format(objRegInventario.dCustoDemo3, "standard")
        
        QuantOutras.Caption = Formata_Estoque(objRegInventario.dQuantOutras)
        ValorOutras.Caption = Format(objRegInventario.dCustoOutras, "standard")
        
        QuantOutras3.Caption = Formata_Estoque(objRegInventario.dQuantOutras3)
        ValorOutras3.Caption = Format(objRegInventario.dCustoOutras3, "standard")
        
        QuantBenef.Caption = Formata_Estoque(objRegInventario.dQuantBenef)
        ValorBenef.Caption = Format(objRegInventario.dCustoBenef, "standard")
        
        QuantBenef3.Caption = Formata_Estoque(objRegInventario.dQuantBenef3)
        ValorBenef3.Caption = Format(objRegInventario.dCustoBenef3, "standard")
    
    End If
    
    iAlterado = 0

    Traz_RegInventario_Tela = SUCESSO

    Exit Function

Erro_Traz_RegInventario_Tela:

    Traz_RegInventario_Tela = gErr

    Select Case gErr

        Case 70269, 70541

        Case 70542
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO", gErr, objAlmoxarifado.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159279)

    End Select

    Exit Function

End Function

Function Move_Tela_Memoria(objRegInventario As ClassRegInventario) As Long
'Move dados da tela para a memória

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_Move_Tela_Memoria

    objRegInventario.iFilialEmpresa = giFilialEmpresa
    objRegInventario.dtData = StrParaDate(Data.Text)

    'Formata o Produto
    lErro = CF("Produto_Formata", Produto.Caption, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 70270

    objRegInventario.sProduto = sProdutoFormatado
    objRegInventario.sDescricao = Descricao.Caption
    objRegInventario.sModelo = Modelo.Caption
    objRegInventario.sIPICodigo = IPICodigo.Caption

    'Se a natureza foi preenchida
    If NaturezaProduto.ListIndex <> -1 Then
        objRegInventario.iNatureza = NaturezaProduto.ItemData(NaturezaProduto.ListIndex)
    End If

    objRegInventario.sSiglaUMEstoque = UM.Caption
    objRegInventario.dQuantidadeUMEstoque = StrParaDbl(Quantidade.Caption)
    objRegInventario.dValorUnitario = StrParaDbl(ValorUnitario.Caption)
    objRegInventario.sObservacoes = Observacoes.Text

    'Se o Almoxarifado foi preenchido
    If Len(Trim(AlmoxarifadoLabel.Caption)) > 0 Then

        objAlmoxarifado.sNomeReduzido = AlmoxarifadoLabel.Caption
        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25060 Then gError 70939

        'Se não encontrou o Almoxarifado, erro
        If lErro = 25060 Then gError 70940

        objRegInventario.iAlmoxarifado = objAlmoxarifado.iCodigo

    End If

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Select Case gErr

        Case 70270, 70939

        Case 70940
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO1", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159280)

    End Select

    Exit Function

End Function

Private Sub BotaoRegCadastrado_Click()

Dim colSelecao As New Collection
Dim objRegInventario As ClassRegInventario

    Call Chama_Tela("RegInventarioLista", colSelecao, objRegInventario, objEventoBotaoInventario)

End Sub

Private Sub objEventoBotaoInventario_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRegInventario As ClassRegInventario

On Error GoTo Erro_objEventoBotaoInventario_evSelecao

    Set objRegInventario = obj1

    'Limpa a tela
    Call Limpa_Tela_RegInventario

    'Coloca Data do Inventário na tela
    If objRegInventario.dtData <> DATA_NULA Then
        Call DateParaMasked(Data, objRegInventario.dtData)
    End If

    Me.Show

    Exit Sub

Erro_objEventoBotaoInventario_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159281)

    End Select

    Exit Sub

End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    'Se a Data está preenchida
    If Len(Trim(Data.ClipText)) > 0 Then

        'Critica seu formato
        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then gError 70271

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case 70271

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159282)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRegInventario As New ClassRegInventario
Dim objRegInventarioAlmox As New ClassRegInventarioAlmox

On Error GoTo Erro_objEventoBotaoInventario_evSelecao

    Set objRegInventario = obj1
    
    If Almoxarifado.ListIndex <> -1 Then
        
        objRegInventarioAlmox.sProduto = objRegInventario.sProduto
        objRegInventarioAlmox.dtData = objRegInventario.dtData
        objRegInventarioAlmox.iAlmoxarifado = Almoxarifado.ItemData(Almoxarifado.ListIndex)
        
        lErro = CF("RegInventarioAlmox_Le", objRegInventarioAlmox)
        If lErro <> SUCESSO And lErro <> 69885 Then gError 69886
            
    Else
        
        objRegInventario.iFilialEmpresa = giFilialEmpresa
    
        lErro = CF("RegInventario_Le", objRegInventario)
        If lErro <> SUCESSO Then gError 81872
            
    End If
    
    'Traz o Registro de inventário para a tela
    lErro = Traz_RegInventario_Tela(objRegInventario, objRegInventarioAlmox)
    If lErro <> SUCESSO Then gError 70540

    Me.Show

    Exit Sub

Erro_objEventoBotaoInventario_evSelecao:

    Select Case gErr

        Case 70540, 69886, 81872

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159283)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoFinal_Change()

    iFrameFiltroAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProdutoInicial_Change()

    iFrameFiltroAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    'Se a data está preenchida
    If Len(Trim(Data.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 70272

    End If

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 70272

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159284)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    'Se a data está preenchida
    If Len(Trim(Data.ClipText)) > 0 Then

        'Aumenta a data em um dia
        lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 70273

    End If

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 70273

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159285)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String

On Error GoTo Erro_ProdutoInicial_Validate

    'Se o código do produto não foi preenchido, sai da rotina
    If Len(Trim(ProdutoInicial.ClipText)) = 0 Then
        DescProdInic.Caption = ""
        Exit Sub
    End If

    'Formata o Produto
    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 70280

    objProduto.sCodigo = sProdutoFormatado

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 70281

    'Se o Produto não está cadastrado, pergunta se deseja criar
    If lErro = 28030 Then gError 70282

    DescProdInic.Caption = objProduto.sDescricao

    'Se o código do produto inicial é maior que o final, erro
    If ProdutoInicial.Text > ProdutoFinal.Text And Len(Trim(ProdutoFinal.ClipText)) > 0 Then gError 70283

    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 70280, 70281

        Case 70282
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", objProduto.sCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("Produto", objProduto)
            End If

        Case 70283
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTODE_MAIOR_PRODUTOATE", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159286)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String

On Error GoTo Erro_ProdutoFinal_Validate

    'Se o código do produto não foi preenchido, sai da rotina
    If Len(Trim(ProdutoFinal.ClipText)) = 0 Then
        DescProdFim.Caption = ""
        Exit Sub
    End If

    'Formata o Produto
    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 70285

    objProduto.sCodigo = sProdutoFormatado

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 70286

    'Se o Produto não está cadastrado, pergunta se deseja criar
    If lErro = 28030 Then gError 70287

    DescProdFim.Caption = objProduto.sDescricao

    'Se o código do produto inicial é maior que o final, erro
    If ProdutoInicial.Text > ProdutoFinal.Text And Len(Trim(ProdutoInicial.ClipText)) > 0 Then gError 70289

    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 70285, 70286

        Case 70287
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", objProduto.sCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("Produto", objProduto)
            End If

        Case 70289
           lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTODE_MAIOR_PRODUTOATE", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159287)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoLabel_Click()

Dim colSelecao As New Collection
Dim objRegInventario As New ClassRegInventario
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long
Dim sSelecao As String

On Error GoTo Erro_ProdutoLabel_Click

    'Verifica se a data está preenchida
    If Len(Trim(Data.ClipText)) = 0 Then gError 70545

    'Verifica se existe um Registro de Inventário com a data passada
    objRegInventario.dtData = Data.Text
    objRegInventario.iFilialEmpresa = giFilialEmpresa
    lErro = CF("RegInventario_Le_Data", objRegInventario)
    If lErro <> SUCESSO And lErro <> 70237 Then gError 70546

    'Se não encontrou, erro
    If lErro = 70237 Then gError 70547

    'Verifica se Produto está preenchido
    If Len(Trim(Produto.Caption)) > 0 Then

        'Critica o formato do Produto
        lErro = CF("Produto_Formata", Produto.Caption, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 70301

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objRegInventario.sProduto = sProdutoFormatado
        Else
            objRegInventario.sProduto = ""
        End If

    End If

    'Trata filtros definidos no frame de Filtro de Produtos
    lErro = Trata_Filtro_Browse(colSelecao, sSelecao)
    If lErro <> SUCESSO Then gError 70284
    
    'Inicio Daniel
    colSelecao.Add objRegInventario.dtData
    'Fim Daniel 27/08/2001

    'Chama a tela de browse RegInventarioProdutosLista passando como parâmetro a seleção do Filtro de Produtos (sSelecao)
    Call Chama_Tela("RegInventarioProdutosLista", colSelecao, objRegInventario, objEventoProduto, sSelecao)

    Exit Sub

Erro_ProdutoLabel_Click:

    Select Case gErr

        Case 70284, 70301, 70546

        Case 70545
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)

        Case 70547
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGINVENTARIO_NAO_CADASTRADO1", gErr, objRegInventario.dtData)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159288)

    End Select

    Exit Sub

End Sub

Function Trata_Filtro_Browse(colSelecao As Collection, sSelecao As String) As Long

Dim lErro As Long
Dim iCampoPreenchido As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sContaFormatada As String
Dim iContaPreenchida As Integer

On Error GoTo Erro_Trata_Filtro_Browse

    'Se o Produto Inicial foi preenchido
    If Len(Trim(ProdutoInicial.ClipText)) > 0 Then

        'Formata o Produto
        lErro = CF("Produto_Formata", ProdutoInicial.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 70332

        sSelecao = sSelecao & "Produto >= ?"
        iCampoPreenchido = 1

        'Adiciona em no Filtro, Produto > ProdutoInicial
        colSelecao.Add sProdutoFormatado
    End If

    'Se o Produto Final foi preenchido
    If Len(Trim(ProdutoFinal.ClipText)) > 0 Then

        'Formata o Produto
        lErro = CF("Produto_Formata", ProdutoFinal.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 70344

        If iCampoPreenchido = 1 Then
            sSelecao = sSelecao & " AND Produto <= ?"
        Else
            sSelecao = sSelecao & " Produto <= ?"
            iCampoPreenchido = 1
        End If

        'Adiciona em no Filtro, Produto < ProdutoFinal
        colSelecao.Add sProdutoFormatado

    End If

    'Se selecionou apenas um almoxarifado e ele está preenchido
    If OptionUmTipo.Value = True And Almoxarifado.ListIndex <> -1 Then

        If iCampoPreenchido = 1 Then
            sSelecao = sSelecao & " AND Almoxarifado = ?"
        Else
            sSelecao = sSelecao & " Almoxarifado = ?"
            iCampoPreenchido = 1
        End If

        'Adiciona no Filtro código do Almoxarifado
        colSelecao.Add Almoxarifado.ItemData(Almoxarifado.ListIndex)
    End If

    'Se selecionou apenas uma Conta Contábil e ela foi preenchida
    If OptionContaUma.Value = True And Len(Trim(ContaContabil.ClipText)) > 0 Then

        'Formata a Conta Contábil
        lErro = CF("Conta_Formata", ContaContabil.Text, sContaFormatada, iContaPreenchida)
        If lErro <> SUCESSO Then gError 70538

        If iCampoPreenchido = 1 Then
            sSelecao = sSelecao & " AND ContaContabil = ?"
        Else
            sSelecao = sSelecao & " ContaContabil = ?"
            iCampoPreenchido = 1
        End If

        'Adiciona Conta Contábil no Sistema de Setas
        colSelecao.Add sContaFormatada

    End If

    Trata_Filtro_Browse = SUCESSO

    Exit Function

Erro_Trata_Filtro_Browse:

    Trata_Filtro_Browse = gErr

    Select Case gErr

        Case 70332, 70344, 70538

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159289)

    End Select

    Exit Function

End Function

Private Sub Data_GotFocus()

    Call MaskEdBox_TrataGotFocus(Data, iAlterado)

End Sub

Private Sub Data_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Observacoes_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub OptionTodosTipos_Click()

    Almoxarifado.Enabled = False
    Almoxarifado.ListIndex = -1
    iFrameFiltroAlterado = REGISTRO_ALTERADO

End Sub

Private Sub OptionUmTipo_Click()

    Almoxarifado.Enabled = True
    iFrameFiltroAlterado = REGISTRO_ALTERADO

End Sub

Private Sub OptionContaTodas_Click()

    ContaContabil.Enabled = False
    ContaContabil.PromptInclude = False
    ContaContabil.Text = ""
    ContaContabil.PromptInclude = True
    iFrameFiltroAlterado = REGISTRO_ALTERADO

End Sub

Private Sub OptionContaUma_Click()

    ContaContabil.Enabled = True
    iFrameFiltroAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TabStrip1_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

       If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Se o frame atual é o de Edição e o TAB de Filtros foi alterado
        If (TabStrip1.SelectedItem.Index = TAB_EDICAO Or TabStrip1.SelectedItem.Index = TAB_TERCEIROS) And iFrameFiltroAlterado = REGISTRO_ALTERADO Then

            'Limpa o Frame de Edição
            Call Limpa_Tela_Edicao

        End If

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index

    End If

End Sub

Private Sub ProdutoInicialLabel_Click()

Dim colSelecao As New Collection
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_ProdutoInicialLabel_Click

    'Verifica se Produto está preenchido
    If Len(Trim(ProdutoInicial.ClipText)) > 0 Then

        'Critica o formato do Produto
        lErro = CF("Produto_Formata", ProdutoInicial.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 70290

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objProduto.sCodigo = sProdutoFormatado
        Else
            objProduto.sCodigo = ""
        End If

    End If

    'Chama a tela de browse
    Call Chama_Tela("ProdutoLista_Consulta", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_ProdutoInicialLabel_Click:

    Select Case gErr

        Case 70290

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159290)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoDe_evSelecao

    Set objProduto = obj1

    ProdutoInicial.PromptInclude = False
    ProdutoInicial.Text = objProduto.sCodigo
    ProdutoInicial.PromptInclude = True

    DescProdInic.Caption = objProduto.sDescricao

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159291)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoFinalLabel_Click()

Dim colSelecao As New Collection
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_ProdutoFinalLabel_Click

    'Verifica se Produto está preenchido
    If Len(Trim(ProdutoFinal.ClipText)) > 0 Then

        'Critica o formato do Produto
        lErro = CF("Produto_Formata", ProdutoFinal.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 70291

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objProduto.sCodigo = sProdutoFormatado
        Else
            objProduto.sCodigo = ""
        End If

    End If

    'Chama a tela de browse
    Call Chama_Tela("ProdutoLista_Consulta", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_ProdutoFinalLabel_Click:

    Select Case gErr

        Case 70291

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159292)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    ProdutoFinal.PromptInclude = False
    ProdutoFinal.Text = objProduto.sCodigo
    ProdutoFinal.PromptInclude = True

    DescProdFim.Caption = objProduto.sDescricao

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159293)

    End Select

    Exit Sub

End Sub

Private Sub BotaoContaContabil_Click()
'Chama o browser de plano de contas

Dim lErro As Long
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim iContaPreenchida As Integer
Dim sConta As String

On Error GoTo Erro_BotaoContaContabil_Click

    'Se foram selecionadas todas as contas, sai da rotina
    If OptionContaTodas.Value = True Then Exit Sub

    sConta = String(STRING_CONTA, 0)

    'Formata a Conta
    lErro = CF("Conta_Formata", ContaContabil.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then gError 70292

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    Call Chama_Tela("PlanoContaLista", colSelecao, objPlanoConta, objEventoContaContabil)

    Exit Sub

Erro_BotaoContaContabil_Click:

    Select Case gErr

        Case 70292

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159294)

    End Select

    Exit Sub

End Sub

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
        If lErro <> SUCESSO Then gError 70293

        ContaContabil.Text = sContaEnxuta
        ContaContabil.PromptInclude = True
    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case gErr

        Case 70293
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159295)

    End Select

    Exit Sub

End Sub

Private Sub ContaContabil_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaFormatada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim sContaMascarada As String

On Error GoTo Erro_ContaContabil_Validate

    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica_Modulo", ContaContabil.Text, ContaContabil.ClipText, objPlanoConta, MODULO_LIVROSFISCAIS)
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 70294

    If lErro = SUCESSO Then

        sContaFormatada = objPlanoConta.sConta

        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then gError 70295

        ContaContabil.PromptInclude = False
        ContaContabil.Text = sContaMascarada
        ContaContabil.PromptInclude = True


    'Se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then

        'Critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
        lErro = CF("Conta_Critica", ContaContabil.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
        If lErro <> SUCESSO And lErro <> 5700 Then gError 70296

        'Conta não cadastrada
        If lErro = 5700 Then gError 70297

    End If

    Exit Sub

Erro_ContaContabil_Validate:

    Cancel = True

    Select Case gErr

        Case 70294, 70296

        Case 70295
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case 70297
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", gErr, ContaContabil.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159296)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava dados do Registro de Inventário
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 70298

    'Limpa a tela
    Call Limpa_Tela_RegInventario

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 70298

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159297)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objRegInventario As New ClassRegInventario

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se a data foi preenchida
    If Len(Trim(Data.ClipText)) = 0 Then gError 70299

    'Verifica se o Produto Foi preenchido
    If Len(Trim(Produto.Caption)) = 0 Then gError 70300

    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objRegInventario)
    If lErro <> SUCESSO Then gError 70302

    'Grava um Registro de Inventário
    lErro = CF("RegInventario_Atualiza", objRegInventario)
    If lErro <> SUCESSO Then gError 70303

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 70302, 70303

        Case 70299
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)

        Case 70300
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159298)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objRegInventario As New ClassRegInventario

On Error GoTo Erro_BotaoExcluir_Click

    'Verifica se a data foi preenchida
    If Len(Trim(Data.ClipText)) = 0 Then gError 70318

    'Verifica se o Produto Foi preenchido
    If Len(Trim(Produto.Caption)) = 0 Then gError 70319

    'Guarda dados do Registro de Inventário
    lErro = Move_Tela_Memoria(objRegInventario)
    If lErro <> SUCESSO Then gError 70304

    'Lê o Registro de Inventário a partir da Data, Produto e FilialEmpresa
    lErro = CF("RegInventario_Le", objRegInventario)
    If lErro <> SUCESSO And lErro <> 70308 Then gError 70309

    'Se não encontrou, erro
    If lErro = 70309 Then gError 70310

    'Pede a confirmação da exclusão do Registro de Inventário
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_REGIVENTARIO", objRegInventario.sProduto, objRegInventario.dtData, objRegInventario.iFilialEmpresa)
    If vbMsgRes = vbNo Then Exit Sub

    GL_objMDIForm.MousePointer = vbHourglass

    'Exclui Registro de Inventário
    lErro = CF("RegInventario_Exclui", objRegInventario)
    If lErro <> SUCESSO Then gError 70311

    'Limpa a tela
    Call Limpa_Tela_RegInventario

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 70304, 70309, 70311

        Case 70310
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGINVENTARIO_NAO_CADASTRADO", gErr, objRegInventario.sProduto, objRegInventario.dtData, objRegInventario.iFilialEmpresa)

        Case 70318
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)

        Case 70319
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159299)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 70312

    'Limpa a tela
    Call Limpa_Tela_RegInventario

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 70312

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159300)

    End Select

    Exit Sub

End Sub

Sub Limpa_Tela_RegInventario()

    'Função Genérica que limpa a tela
    Call Limpa_Tela(Me)

    'Limpa restante dos campos
    DescProdInic.Caption = ""
    DescProdFim.Caption = ""
    OptionContaTodas.Value = True
    OptionTodosTipos.Value = True

    'Limpa a parte de edição
    Call Limpa_Tela_Edicao

End Sub

Sub Limpa_Tela_Edicao()

    Produto.Caption = ""
    ProdutoTerc.Caption = ""
    Descricao.Caption = ""
    Modelo.Caption = ""
    IPICodigo.Caption = ""
    NaturezaProduto.ListIndex = -1
    UM.Caption = ""
    UnidMedLabel.Caption = ""
    Quantidade.Caption = ""
    ValorUnitario.Caption = ""
    AlmoxarifadoLabel.Caption = ""
    AlmoxarifadoCaption.Caption = ""
    Observacoes.Text = ""
    
    QuantConserto.Caption = ""
    QuantBenef.Caption = ""
    QuantBenef3.Caption = ""
    QuantConserto3.Caption = ""
    QuantConsig.Caption = ""
    QuantConsig3.Caption = ""
    QuantDemo.Caption = ""
    QuantDemo3.Caption = ""
    QuantOutras.Caption = ""
    QuantOutras3.Caption = ""
    QuantBenef.Caption = ""
    QuantBenef3.Caption = ""
    ValorTotal.Caption = ""
    ValorConserto.Caption = ""
    ValorConserto3.Caption = ""
    ValorConsig.Caption = ""
    ValorConsig3.Caption = ""
    ValorDemo.Caption = ""
    ValorDemo3.Caption = ""
    ValorOutras.Caption = ""
    ValorOutras3.Caption = ""
    ValorBenef.Caption = ""
    ValorBenef3.Caption = ""
    CustoConserto.Caption = ""
    CustoConserto3.Caption = ""
    CustoConsig.Caption = ""
    CustoConsig3.Caption = ""
    CustoDemo.Caption = ""
    CustoDemo3.Caption = ""
    CustoOutras.Caption = ""
    CustoOutras3.Caption = ""
    CustoBenef.Caption = ""
    CustoBenef3.Caption = ""
    
    iFrameFiltroAlterado = 0

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then

        If Me.ActiveControl Is Data Then
            Call BotaoRegCadastrado_Click
        ElseIf Me.ActiveControl Is ProdutoInicial Then
            Call ProdutoInicialLabel_Click
        ElseIf Me.ActiveControl Is ProdutoFinal Then
            Call ProdutoFinalLabel_Click
        ElseIf Me.ActiveControl Is ContaContabil Then
            Call BotaoContaContabil_Click
        End If

    End If

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Edição de Registro de Inventário"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "EdicaoRegInventario"

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

'**** fim do trecho a ser copiado *****

'
''se ainda nao existir reg em reginventario entao obter dado do sistema p/as qtdes.
'
''Tabela: RegInventário
'
''TAB DE IDENTIFICACAO
'
''1 - Data -  Identificacao do Inventario, quando o usuario
''preencher uma data será verificado se existe um inventario com
''a data, se não existir será pesquisado no sistema o inventario
''na Data, agora se já existir não faz nada.
'
''Disparar Batch na criação do inventário.
'
''TAB DE FILTRO
''Tendo Preenchido a Identificação ele poderá filtrar os produtos
''por Produto de e até, por Conta Contabil, por Almoxarifado,
''para poder identificar melhor o Produto.
'
''TAB DE IDENTIFICACAO
''O Unico campo alteravel é o campo de Observação.

'******** Vai para o BatchFIs, apagar depois ********************
Public Function Rotina_Geracao_RegInventario(dtData As Date) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Rotina_Geracao_RegInventario

    'Abre Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 70241

    'Gera Registro de Inventário por produto
    lErro = CF("RegInventario_Geracao_Trans", dtData)
    If lErro <> SUCESSO Then gError 70242

    'Gera Registro de Inventário por Almoxarifado
    lErro = CF("RegInventarioAlmox_Geracao_Trans", dtData)
    If lErro <> SUCESSO Then gError 70243

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 70244

    Rotina_Geracao_RegInventario = SUCESSO

    Exit Function

    Rotina_Geracao_RegInventario = gErr

Erro_Rotina_Geracao_RegInventario:

    Select Case gErr

        Case 70241
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 70242, 70243

        Case 70244
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159301)

    End Select

    Call Transacao_Rollback
    
    Exit Function

End Function



Private Sub Label2_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label2(Index), Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2(Index), Button, Shift, X, Y)
End Sub


Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub AlmoxarifadoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(AlmoxarifadoLabel, Source, X, Y)
End Sub

Private Sub AlmoxarifadoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(AlmoxarifadoLabel, Button, Shift, X, Y)
End Sub

Private Sub ValorUnitario_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorUnitario, Source, X, Y)
End Sub

Private Sub ValorUnitario_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorUnitario, Button, Shift, X, Y)
End Sub

Private Sub Quantidade_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Quantidade, Source, X, Y)
End Sub

Private Sub Quantidade_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Quantidade, Button, Shift, X, Y)
End Sub

Private Sub IPICodigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(IPICodigo, Source, X, Y)
End Sub

Private Sub IPICodigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(IPICodigo, Button, Shift, X, Y)
End Sub

Private Sub Modelo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Modelo, Source, X, Y)
End Sub

Private Sub Modelo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Modelo, Button, Shift, X, Y)
End Sub

Private Sub Descricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Descricao, Source, X, Y)
End Sub

Private Sub Descricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Descricao, Button, Shift, X, Y)
End Sub

Private Sub Produto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Produto, Source, X, Y)
End Sub

Private Sub Produto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Produto, Button, Shift, X, Y)
End Sub

Private Sub UM_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(UM, Source, X, Y)
End Sub

Private Sub UM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(UM, Button, Shift, X, Y)
End Sub

Private Sub Label48_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label48, Source, X, Y)
End Sub

Private Sub Label48_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label48, Button, Shift, X, Y)
End Sub

Private Sub LblUMEstoque_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblUMEstoque, Source, X, Y)
End Sub

Private Sub LblUMEstoque_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblUMEstoque, Button, Shift, X, Y)
End Sub

Private Sub ProdutoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ProdutoLabel, Source, X, Y)
End Sub

Private Sub ProdutoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ProdutoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub DescProdFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescProdFim, Source, X, Y)
End Sub

Private Sub DescProdFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescProdFim, Button, Shift, X, Y)
End Sub

Private Sub DescProdInic_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescProdInic, Source, X, Y)
End Sub

Private Sub DescProdInic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescProdInic, Button, Shift, X, Y)
End Sub

Private Sub ProdutoInicialLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ProdutoInicialLabel, Source, X, Y)
End Sub

Private Sub ProdutoInicialLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ProdutoInicialLabel, Button, Shift, X, Y)
End Sub

Private Sub ProdutoFinalLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ProdutoFinalLabel, Source, X, Y)
End Sub

Private Sub ProdutoFinalLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ProdutoFinalLabel, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub
