VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl GeracaoOP 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5460
      Index           =   3
      Left            =   0
      TabIndex        =   28
      Top             =   495
      Visible         =   0   'False
      Width           =   9435
      Begin VB.Frame FrameItemPV 
         Caption         =   "Pedidos de Venda associados ao item ()"
         Height          =   1455
         Left            =   30
         TabIndex        =   82
         Top             =   3960
         Width           =   9390
         Begin VB.TextBox ItemPVUM 
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            Enabled         =   0   'False
            Height          =   225
            Left            =   2370
            TabIndex        =   89
            Text            =   "UM"
            Top             =   345
            Width           =   510
         End
         Begin VB.TextBox ItemPVFilialPV 
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            Enabled         =   0   'False
            Height          =   225
            Left            =   2745
            TabIndex        =   88
            Text            =   "Filial do PV"
            Top             =   1125
            Width           =   1170
         End
         Begin VB.TextBox ItemPVDataEmissao 
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            Enabled         =   0   'False
            Height          =   225
            Left            =   4335
            TabIndex        =   87
            Text            =   "Emissão"
            Top             =   645
            Width           =   1200
         End
         Begin VB.TextBox ItemPVFilial 
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            Enabled         =   0   'False
            Height          =   225
            Left            =   3270
            TabIndex        =   86
            Text            =   "Filial"
            Top             =   675
            Width           =   870
         End
         Begin VB.TextBox ItemPVCliente 
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            Enabled         =   0   'False
            Height          =   225
            Left            =   1815
            TabIndex        =   85
            Text            =   "Cliente"
            Top             =   645
            Width           =   1440
         End
         Begin VB.TextBox ItemPVPedido 
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            Enabled         =   0   'False
            Height          =   225
            Left            =   4530
            TabIndex        =   84
            Text            =   "Pedido"
            Top             =   1110
            Width           =   840
         End
         Begin VB.TextBox ItemPVDataEntrega 
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            Enabled         =   0   'False
            Height          =   225
            Left            =   4320
            TabIndex        =   83
            Text            =   "Entrega"
            Top             =   330
            Width           =   1200
         End
         Begin MSMask.MaskEdBox ItemPVQuantidade 
            Height          =   225
            Left            =   3030
            TabIndex        =   90
            Top             =   345
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
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
         Begin MSFlexGridLib.MSFlexGrid GridItemPV 
            Height          =   1140
            Left            =   45
            TabIndex        =   45
            Top             =   240
            Width           =   9300
            _ExtentX        =   16404
            _ExtentY        =   2011
            _Version        =   393216
            Rows            =   21
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
            AllowUserResizing=   1
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Geração automática"
         Height          =   1245
         Left            =   30
         TabIndex        =   78
         Top             =   -30
         Width           =   9420
         Begin VB.CheckBox GeraOpSub 
            Caption         =   "Gera O.P. para subprodutos"
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
            TabIndex        =   34
            Top             =   615
            Width           =   2790
         End
         Begin VB.CheckBox GeraReqCompra 
            Caption         =   "Gera Req. de Compras para matérias-primas"
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
            TabIndex        =   37
            Top             =   900
            Width           =   4065
         End
         Begin VB.TextBox CodOPGerada 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5655
            MaxLength       =   9
            TabIndex        =   35
            Top             =   615
            Width           =   1305
         End
         Begin VB.CheckBox GeraOPs 
            Caption         =   "Gera uma O.P. para cada subproduto"
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
            Height          =   270
            Left            =   4770
            TabIndex        =   39
            Top             =   930
            Width           =   4170
         End
         Begin VB.CommandButton botaoProxNum2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6960
            Picture         =   "GeracaoOP.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Numeração Automática"
            Top             =   615
            Width           =   300
         End
         Begin VB.CommandButton BotaoOPRC 
            Caption         =   "..."
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
            Height          =   300
            Left            =   4200
            TabIndex        =   38
            Top             =   885
            Width           =   360
         End
         Begin VB.CommandButton BotaoGerarOP 
            Caption         =   "Gerar Ordem de Produção"
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
            Height          =   330
            Left            =   5655
            TabIndex        =   40
            Top             =   240
            Width           =   2715
         End
         Begin VB.TextBox Codigo 
            Height          =   285
            Left            =   1320
            MaxLength       =   9
            TabIndex        =   30
            Top             =   240
            Width           =   1350
         End
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   2670
            Picture         =   "GeracaoOP.ctx":00EA
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Numeração Automática"
            Top             =   255
            Width           =   300
         End
         Begin MSComCtl2.UpDown UpDownData 
            Height          =   300
            Left            =   5025
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox Data 
            Height          =   300
            Left            =   3990
            TabIndex        =   32
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label LabelCodOPGerada 
            AutoSize        =   -1  'True
            Caption         =   "Código O.P. a ser gerada:"
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
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   3390
            TabIndex        =   81
            Top             =   660
            Width           =   2220
         End
         Begin VB.Label LabelCodigo 
            AutoSize        =   -1  'True
            Caption         =   "Código O.P.:"
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
            TabIndex        =   80
            Top             =   255
            Width           =   1095
         End
         Begin VB.Label Label2 
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
            Left            =   3435
            TabIndex        =   79
            Top             =   255
            Width           =   480
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Material a ser produzido"
         Height          =   2730
         Left            =   30
         TabIndex        =   65
         Top             =   1200
         Width           =   9420
         Begin MSMask.MaskEdBox QuantOP 
            Height          =   255
            Left            =   1110
            TabIndex        =   77
            Top             =   615
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
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
         Begin VB.CommandButton BotaoEstoque 
            Caption         =   "Estoque"
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
            Left            =   4185
            TabIndex        =   43
            Top             =   2295
            Width           =   1815
         End
         Begin VB.CommandButton BotaoCcls 
            Caption         =   "Centros de Custo"
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
            Left            =   7470
            TabIndex        =   44
            Top             =   2295
            Width           =   1815
         End
         Begin VB.CommandButton BotaoImprimirPrevia 
            Caption         =   "Prévia de Subprodutos e Compras"
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
            Left            =   60
            TabIndex        =   42
            Top             =   2310
            Width           =   3660
         End
         Begin VB.TextBox DescricaoItemOP 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   4320
            MaxLength       =   50
            TabIndex        =   68
            Top             =   630
            Width           =   2010
         End
         Begin VB.TextBox UMOP 
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            Enabled         =   0   'False
            Height          =   255
            Left            =   1800
            TabIndex        =   67
            Text            =   "UM"
            Top             =   990
            Width           =   552
         End
         Begin VB.TextBox FilialPVOP 
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            Enabled         =   0   'False
            Height          =   255
            Left            =   3270
            TabIndex        =   66
            Top             =   1425
            Width           =   1170
         End
         Begin MSMask.MaskEdBox Prioridade 
            Height          =   255
            Left            =   5835
            TabIndex        =   69
            Top             =   1425
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   3
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantFalta 
            Height          =   255
            Left            =   3465
            TabIndex        =   70
            Top             =   1020
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
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
         Begin MSMask.MaskEdBox PedidoDeVendaOP 
            Height          =   255
            Left            =   4425
            TabIndex        =   71
            Top             =   1425
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdutoOP 
            Height          =   255
            Left            =   285
            TabIndex        =   72
            Top             =   1140
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Almoxarifado 
            Height          =   255
            Left            =   4545
            TabIndex        =   73
            Top             =   1050
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Ccl 
            Height          =   255
            Left            =   5940
            TabIndex        =   74
            Top             =   1005
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            AllowPrompt     =   -1  'True
            MaxLength       =   10
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
         Begin MSMask.MaskEdBox DataPrevisaoFim 
            Height          =   255
            Left            =   2205
            TabIndex        =   75
            Top             =   1440
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataPrevisaoInicio 
            Height          =   255
            Left            =   990
            TabIndex        =   76
            Top             =   1455
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridOP 
            Height          =   1965
            Left            =   45
            TabIndex        =   41
            Top             =   255
            Width           =   9300
            _ExtentX        =   16404
            _ExtentY        =   3466
            _Version        =   393216
            Rows            =   21
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
            AllowUserResizing=   1
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5415
      Index           =   2
      Left            =   0
      TabIndex        =   11
      Top             =   465
      Visible         =   0   'False
      Width           =   9450
      Begin VB.CommandButton BotaoEntrega 
         Caption         =   "Datas de Entrega"
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
         Left            =   6690
         TabIndex        =   27
         Top             =   4845
         Width           =   1650
      End
      Begin VB.ComboBox Ordenados 
         Height          =   315
         ItemData        =   "GeracaoOP.ctx":01D4
         Left            =   2550
         List            =   "GeracaoOP.ctx":01D6
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   135
         Width           =   3480
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
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
         Left            =   2910
         Picture         =   "GeracaoOP.ctx":01D8
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   4845
         Width           =   1650
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
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
         Left            =   4800
         Picture         =   "GeracaoOP.ctx":11F2
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   4845
         Width           =   1650
      End
      Begin VB.TextBox DataEntrega 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   4500
         TabIndex        =   17
         Text            =   "Entrega"
         Top             =   2595
         Width           =   1200
      End
      Begin VB.CheckBox GeraOP 
         DragMode        =   1  'Automatic
         Height          =   270
         Left            =   480
         TabIndex        =   13
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox Pedido 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   4710
         TabIndex        =   23
         Text            =   "Pedido"
         Top             =   3375
         Width           =   885
      End
      Begin VB.TextBox Cliente 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   1992
         TabIndex        =   19
         Text            =   "Cliente"
         Top             =   2904
         Width           =   1260
      End
      Begin VB.TextBox Filial 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   3450
         TabIndex        =   20
         Text            =   "Filial"
         Top             =   2940
         Width           =   870
      End
      Begin VB.TextBox DataEmissao 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   4512
         TabIndex        =   21
         Text            =   "Emissão"
         Top             =   2904
         Width           =   1200
      End
      Begin VB.TextBox FilialPV 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   2925
         TabIndex        =   22
         Text            =   "Filial do PV"
         Top             =   3384
         Width           =   1170
      End
      Begin VB.TextBox UM 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   2550
         TabIndex        =   15
         Text            =   "UM"
         Top             =   2610
         Width           =   510
      End
      Begin VB.TextBox DescricaoItem 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   3060
         MaxLength       =   50
         TabIndex        =   18
         Top             =   1680
         Width           =   1965
      End
      Begin MSMask.MaskEdBox Produto 
         Height          =   225
         Left            =   1245
         TabIndex        =   14
         Top             =   2640
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Quantidade 
         Height          =   225
         Left            =   3210
         TabIndex        =   16
         Top             =   2610
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
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
      Begin MSFlexGridLib.MSFlexGrid GridPV 
         Height          =   2940
         Left            =   45
         TabIndex        =   24
         Top             =   540
         Width           =   9390
         _ExtentX        =   16563
         _ExtentY        =   5186
         _Version        =   393216
         Rows            =   15
         Cols            =   7
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
         AllowUserResizing=   1
      End
      Begin VB.Label ItensSelecionados 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1980
         TabIndex        =   48
         Top             =   4950
         Width           =   795
      End
      Begin VB.Label Label4 
         Caption         =   "Ordenados por:"
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
         Left            =   1005
         TabIndex        =   47
         Top             =   180
         Width           =   1410
      End
      Begin VB.Label LabelItensSelecionados 
         Caption         =   "Itens Selecionados:"
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
         Left            =   255
         TabIndex        =   46
         Top             =   4980
         Width           =   1800
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5385
      Index           =   1
      Left            =   45
      TabIndex        =   0
      Top             =   510
      Width           =   9420
      Begin VB.Frame Frame2 
         Caption         =   "Exibe Pedidos de Venda"
         Height          =   4632
         Left            =   1110
         TabIndex        =   49
         Top             =   375
         Width           =   7095
         Begin VB.Frame Frame5 
            Caption         =   "Data Entrega"
            Height          =   795
            Left            =   795
            TabIndex        =   59
            Top             =   2400
            Width           =   5520
            Begin MSComCtl2.UpDown UpDownEntregaAte 
               Height          =   300
               Left            =   4575
               TabIndex        =   60
               TabStop         =   0   'False
               Top             =   330
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown UpDownEntregaDe 
               Height          =   300
               Left            =   1920
               TabIndex        =   61
               TabStop         =   0   'False
               Top             =   330
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataEntregaDe 
               Height          =   300
               Left            =   780
               TabIndex        =   6
               Top             =   330
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox DataEntregaAte 
               Height          =   300
               Left            =   3435
               TabIndex        =   7
               Top             =   330
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin VB.Label Label8 
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
               Height          =   195
               Left            =   3000
               TabIndex        =   63
               Top             =   390
               Width           =   360
            End
            Begin VB.Label Label7 
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
               Height          =   195
               Left            =   375
               TabIndex        =   62
               Top             =   375
               Width           =   315
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Clientes"
            Height          =   795
            Left            =   795
            TabIndex        =   56
            Top             =   1530
            Width           =   5520
            Begin MSMask.MaskEdBox ClienteDe 
               Height          =   300
               Left            =   810
               TabIndex        =   4
               Top             =   315
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ClienteAte 
               Height          =   300
               Left            =   3465
               TabIndex        =   5
               Top             =   330
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
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
               Left            =   420
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   58
               Top             =   345
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
               Height          =   195
               Left            =   3000
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   57
               Top             =   375
               Width           =   360
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Pedidos"
            Height          =   795
            Left            =   795
            TabIndex        =   53
            Top             =   660
            Width           =   5520
            Begin MSMask.MaskEdBox PedidoInicial 
               Height          =   300
               Left            =   825
               TabIndex        =   2
               Top             =   300
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox PedidoFinal 
               Height          =   300
               Left            =   3465
               TabIndex        =   3
               Top             =   315
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin VB.Label LabelPedidoAte 
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
               Height          =   195
               Left            =   2985
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   55
               Top             =   360
               Width           =   360
            End
            Begin VB.Label LabelPedidoDe 
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
               Left            =   405
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   54
               Top             =   345
               Width           =   315
            End
         End
         Begin VB.CheckBox ExibeTodos 
            Caption         =   "Exibe Todos os Pedidos"
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
            Left            =   2280
            TabIndex        =   1
            Top             =   345
            Width           =   2430
         End
         Begin VB.Frame Frame7 
            Caption         =   "Produtos"
            Height          =   795
            Left            =   825
            TabIndex        =   50
            Top             =   3270
            Width           =   5520
            Begin MSMask.MaskEdBox ProdutoDe 
               Height          =   300
               Left            =   735
               TabIndex        =   8
               Top             =   315
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ProdutoAte 
               Height          =   300
               Left            =   3450
               TabIndex        =   9
               Top             =   285
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin VB.Label LabelProdutoAte 
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
               Height          =   195
               Left            =   3030
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   52
               Top             =   345
               Width           =   360
            End
            Begin VB.Label LabelProdutoDe 
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
               Height          =   195
               Left            =   345
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   51
               Top             =   360
               Width           =   315
            End
         End
         Begin VB.CheckBox CheckExibePedidoOP 
            Caption         =   "Inclui Pedidos com Ordens de Produção"
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
            Left            =   1710
            TabIndex        =   10
            Top             =   4260
            Width           =   3750
         End
      End
   End
   Begin VB.CommandButton BotaoFechar 
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
      Left            =   8265
      Picture         =   "GeracaoOP.ctx":23D4
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Fechar"
      Top             =   15
      Width           =   1230
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5865
      Left            =   -45
      TabIndex        =   64
      Top             =   120
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   10345
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pedidos de Venda"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ordens de Produção a Serem Geradas"
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
Attribute VB_Name = "GeracaoOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

Dim gobjGeracaoOP As New ClassGeracaoOP
Dim iTabPrincipalAlterado As Integer
Dim iTabPedidoAlterado As Integer
Dim iAlterado As Integer
Dim iFrameAtual As Integer
Dim iSelecionadosPV As Integer

Dim gobjOP As ClassOrdemDeProducao
Dim gbTrazendoDados As Boolean

Dim asOrdenacao(3) As String
Dim asOrdenacaoString(3) As String

Dim objGridOP As AdmGrid
Const GRIDOP_PRODUTO_COL = 1
Const GRIDOP_DESCRICAOITEMOP_COL = 2
Const GRIDOP_UM_COL = 3
Const GRIDOP_QUANT_COL = 4
Const GRIDOP_QUANTFALTA_COL = 5
Const GRIDOP_ALMOXARIFADO_COL = 6
Const GRIDOP_CCL_COL = 7
Const GRIDOP_DATAPREVISAOINICIO_COL = 8
Const GRIDOP_DATAPREVISAOFIM_COL = 9
'Const GRIDOP_PEDIDODEVENDA_COL = 10
'Const GRIDOP_FILIALPV_COL = 11
Const GRIDOP_PRIORIDADE_COL = 10

Dim objGridPV As AdmGrid
Const GRIDPV_GERAOP_COL = 1
Const GRIDPV_PRODUTO_COL = 2
Const GRIDPV_DESCRICAOITEM_COL = 3
Const GRIDPV_UM_COL = 4
Const GRIDPV_QUANTIDADE_COL = 5
Const GRIDPV_CLIENTE_COL = 6
Const GRIDPV_FILIAL_COL = 7
Const GRIDPV_DATAEMISSAO_COL = 8
Const GRIDPV_DATAENTREGA_COL = 9
Const GRIDPV_PEDIDO_COL = 10
Const GRIDPV_FILIALPV_COL = 11

Dim objGridItemPV As AdmGrid
Const GRIDITEMPV_PEDIDO_COL = 1
Const GRIDITEMPV_FILIALPV_COL = 2
Const GRIDITEMPV_QUANTIDADE_COL = 3
Const GRIDITEMPV_UM_COL = 4
Const GRIDITEMPV_CLIENTE_COL = 5
Const GRIDITEMPV_FILIAL_COL = 6
Const GRIDITEMPV_DATAEMISSAO_COL = 7
Const GRIDITEMPV_DATAENTREGA_COL = 8

Const TABSELECAO = 1
Const TABPV = 2
Const TABOP = 3

'Eventos de Browse
Private WithEvents objEventoPedidoDe As AdmEvento
Attribute objEventoPedidoDe.VB_VarHelpID = -1
Private WithEvents objEventoPedidoAte As AdmEvento
Attribute objEventoPedidoAte.VB_VarHelpID = -1
Private WithEvents objEventoClienteDe As AdmEvento
Attribute objEventoClienteDe.VB_VarHelpID = -1
Private WithEvents objEventoClienteAte As AdmEvento
Attribute objEventoClienteAte.VB_VarHelpID = -1
Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1
Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1
Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoEstoque As AdmEvento
Attribute objEventoEstoque.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_Selecao = 1
Private Const TAB_PV = 2
Private Const TAB_OP_GERADA = 3

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iChamada As Integer)

Dim lErro As Long
Dim sCodProduto As String
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    If (GridPV.Row <> 0) And (iFrameAtual = TABPV) Then

            'Verifica se produto está preenchido
            sCodProduto = GridPV.TextMatrix(GridPV.Row, GRIDPV_PRODUTO_COL)

            lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then Error 41542

            'Pesquisa o controle da coluna em questão
            Select Case objControl.Name

                Case GeraOP.Name

                    'GeraOP
                    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                        objControl.Enabled = True
                    Else
                        objControl.Enabled = False
                    End If

            End Select

    ElseIf (GridOP.Row <> 0) And (iFrameAtual = TABOP) Then

            'Verifica se produto está preenchido
            sCodProduto = GridOP.TextMatrix(GridOP.Row, GRIDOP_PRODUTO_COL)

            lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then Error 41543

            'Pesquisa a controle da coluna em questão
            Select Case objControl.Name

                'Quantidade
                Case QuantOP.Name, Almoxarifado.Name, Ccl.Name, DataPrevisaoInicio.Name, DataPrevisaoFim.Name, Prioridade.Name

                    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                        objControl.Enabled = True
                    Else
                        objControl.Enabled = False
                    End If

            End Select

    End If

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case Err

        Case 41542, 41543

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160896)

    End Select

    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'Se for o GridParcelas
            Case GridOP.Name

                lErro = Saida_Celula_OP(objGridInt)
                If lErro <> SUCESSO Then Error 41544

            'Se for o GridDescontos
            Case GridPV.Name

                lErro = Saida_Celula_PV(objGridInt)
                If lErro <> SUCESSO Then Error 41545

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 41546
    
    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 41544 To 41545

        Case 41546
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160897)

    End Select

    Exit Function

End Function

Private Sub BotaoEstoque_Click()
'Informa se produto é estocado em algum almoxarifado

Dim lErro As Long
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoEstoque_Click

    If GridOP.Row = 0 Then Error 52185

    sCodProduto = GridOP.TextMatrix(GridOP.Row, GRIDOP_PRODUTO_COL)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then Error 52186

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        colSelecao.Add sProdutoFormatado
        'chama a tela de lista de estoque do produto corrente
        Call Chama_Tela("EstoqueProdutoFilialLista", colSelecao, objEstoqueProduto, objEventoEstoque)
    Else
        Error 52187
    End If

    Exit Sub

Erro_BotaoEstoque_Click:

    Select Case Err

        Case 52185
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", Err)
        
        Case 52186
        
        Case 52187
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 160898)

    End Select

    Exit Sub

End Sub

Private Sub BotaoImprimirPrevia_Click()

Dim lErro As Long
Dim objCalcNecesProd As New ClassCalcNecesProd
Dim iIndice As Integer
Dim iProdutoPreenchido As Integer
Dim sProduto As String, sCcl As String
Dim sProdutoFormatado As String, bAchou As Boolean
Dim objNecesProdInfo As ClassNecesProdInfo

On Error GoTo Erro_BotaoImprimirPrevia_Click

    bAchou = False
    
    'Para cada item do grid, guarda em um objeto os dados do grid
    For iIndice = 1 To objGridOP.iLinhasExistentes

        Set objNecesProdInfo = New ClassNecesProdInfo

        sProduto = GridOP.TextMatrix(iIndice, GRIDOP_PRODUTO_COL)

        'Critica o formato do Produto
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 124215

        objNecesProdInfo.sProduto = sProdutoFormatado

        objNecesProdInfo.sUMNecesInfo = GridOP.TextMatrix(iIndice, GRIDOP_UM_COL)

        If Len(Trim(GridOP.TextMatrix(iIndice, GRIDOP_QUANT_COL))) > 0 Then
            objNecesProdInfo.dQuantNecesInfo = CDbl(GridOP.TextMatrix(iIndice, GRIDOP_QUANT_COL))
            If objNecesProdInfo.dQuantNecesInfo > 0 Then bAchou = True
        Else
            objNecesProdInfo.dQuantNecesInfo = 0
        End If
    
        Set objNecesProdInfo.objItemOP = gobjOP.colItens(iIndice)
    
        objCalcNecesProd.colNecesInfProd.Add objNecesProdInfo
    
    Next
    
    If bAchou = False Then gError 124216
    
    objCalcNecesProd.iFilialEmpresa = giFilialEmpresa
    
    lErro = CF("Producao_Calcula_Necessidades", objCalcNecesProd)
    If lErro <> SUCESSO Then gError 124217
    
    lErro = CF("Rel_Producao_Calcula_Necessidades", objCalcNecesProd)
    If lErro <> SUCESSO Then gError 124217
    
    Exit Sub
    
Erro_BotaoImprimirPrevia_Click:

    Select Case gErr
        
        Case 124215, 124217
        
        Case 124216
            Call Rotina_Erro(vbOKOnly, "ERRO_NAO_DEFINIU_QTDE_PROD", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160899)
     
    End Select
     
    Exit Sub

End Sub

Private Sub ClienteAte_GotFocus()

Dim iTabAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    Call MaskEdBox_TrataGotFocus(ClienteAte, iAlterado)
    iTabPrincipalAlterado = iTabAux

End Sub

Private Sub ClienteDe_GotFocus()

Dim iTabAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    Call MaskEdBox_TrataGotFocus(ClienteDe, iAlterado)
    iTabPrincipalAlterado = iTabAux
    
End Sub

Private Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)

End Sub

Private Sub DataEntregaAte_GotFocus()

Dim iTabAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    Call MaskEdBox_TrataGotFocus(DataEntregaAte, iAlterado)
    iTabPrincipalAlterado = iTabAux

End Sub

Private Sub DataEntregaDe_GotFocus()

Dim iTabAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    Call MaskEdBox_TrataGotFocus(DataEntregaDe, iAlterado)
    iTabPrincipalAlterado = iTabAux

End Sub

Private Sub objEventoEstoque_evselecao(obj1 As Object)

Dim objEstoqueProduto As New ClassEstoqueProduto
Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sCodProduto As String

On Error GoTo Erro_objEventoEstoque_evselecao

    If GridOP.Row <> 0 Then

        Set objEstoqueProduto = obj1

        sCodProduto = GridOP.TextMatrix(GridOP.Row, GRIDOP_PRODUTO_COL)

        lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then Error 52189

        'Verifica se o produto está preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            'Preenche o Nome do Almoxarifado
            GridOP.TextMatrix(GridOP.Row, GRIDOP_ALMOXARIFADO_COL) = objEstoqueProduto.sAlmoxarifadoNomeReduzido

            Almoxarifado.Text = objEstoqueProduto.sAlmoxarifadoNomeReduzido

        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoEstoque_evselecao:

    Select Case Err

        Case 52189

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 160900)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objOrdemDeProducao As New ClassOrdemDeProducao

On Error GoTo Erro_Codigo_Validate

    If Len(Trim(Codigo.Text)) = 0 Then Exit Sub
    
    objOrdemDeProducao.sCodigo = Codigo.Text
    objOrdemDeProducao.iFilialEmpresa = giFilialEmpresa
    
    'tenta ler a OP desejada
    lErro = CF("OrdemProducao_Le", objOrdemDeProducao)
    If lErro <> SUCESSO And lErro <> 30368 And lErro <> 55316 Then Error 55877

    'ordem de producao baixada
    If lErro = 55316 Then Error 55878
    
    'se já existir uma ordem de producao com este codigo
    If lErro = SUCESSO Then Error 55879
    
    Exit Sub
    
Erro_Codigo_Validate:

    Cancel = True


    Select Case Err
    
        Case 46603, 55877
        
        Case 55878
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_BAIXADA", Err, Codigo.Text)
        
        Case 55879
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_JA_CADASTRADA", Err, Codigo.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160901)
            
    End Select
        
    Exit Sub

End Sub

Private Sub GridOP_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridOP, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridOP, iAlterado)
    End If

End Sub

Private Sub GridOP_EnterCell()
    Call Grid_Entrada_Celula(objGridOP, iAlterado)
End Sub

Private Sub GridOP_GotFocus()
    Call Grid_Recebe_Foco(objGridOP)
End Sub

Private Sub GridOP_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridOP)
End Sub

Private Sub GridOP_KeyPress(KeyAscii As Integer)
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridOP, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridOP, iAlterado)
    End If

End Sub

Private Sub GridOP_LeaveCell()
    Call Saida_Celula(objGridOP)
End Sub

Private Sub GridOP_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridOP)
End Sub

Private Sub GridOP_RowColChange()
    Call Grid_RowColChange(objGridOP)
    If Not gbTrazendoDados Then
        If objGridOP.iLinhaAntiga <> objGridOP.objGrid.Row Then
            Call Trata_ItemPV
            objGridOP.iLinhaAntiga = objGridOP.objGrid.Row
        End If
    End If
End Sub

Private Sub GridOP_Scroll()
    Call Grid_Scroll(objGridOP)
End Sub

Private Function Inicializa_GridOP(objGridInt As AdmGrid) As Long
'Inicializa o GridOP
Dim lErro As Long

On Error GoTo Erro_Inicializa_GridOP

    'Form do GridOP
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quant. OP")
    objGridInt.colColuna.Add ("Quant. Falta")
    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("Ccl")
    objGridInt.colColuna.Add ("Previsão Início")
    objGridInt.colColuna.Add ("Previsão Fim")
    'objGridInt.colColuna.Add ("Pedido de Venda")
    'objGridInt.colColuna.Add ("Filial do Pedido")
    objGridInt.colColuna.Add ("Prioridade")

    'Controles que participam do GridPV
    objGridInt.colCampo.Add (ProdutoOP.Name)
    objGridInt.colCampo.Add (DescricaoItemOP.Name)
    objGridInt.colCampo.Add (UMOP.Name)
    objGridInt.colCampo.Add (QuantOP.Name)
    objGridInt.colCampo.Add (QuantFalta.Name)
    objGridInt.colCampo.Add (Almoxarifado.Name)
    objGridInt.colCampo.Add (Ccl.Name)
    objGridInt.colCampo.Add (DataPrevisaoInicio.Name)
    objGridInt.colCampo.Add (DataPrevisaoFim.Name)
    'objGridInt.colCampo.Add (PedidoDeVendaOP.Name)
    'objGridInt.colCampo.Add (FilialPVOP.Name)
    objGridInt.colCampo.Add (Prioridade.Name)
    
    PedidoDeVendaOP.left = POSICAO_FORA_TELA
    FilialPVOP.left = POSICAO_FORA_TELA

    objGridInt.objGrid = GridOP

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE + 1

    'linhas visiveis do grid sem contar com as linhas fixas
    objGridInt.iLinhasVisiveis = 5

    'Largura da primeira coluna
    GridOP.ColWidth(0) = 400

    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    Call Grid_Inicializa(objGridInt)

    Inicializa_GridOP = SUCESSO

    Exit Function

Erro_Inicializa_GridOP:

    Inicializa_GridOP = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160902)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_OP(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_OP

    Select Case GridOP.Col

        Case GRIDOP_QUANT_COL
            lErro = Saida_Celula_QuantOP(objGridInt)
            If lErro <> SUCESSO Then Error 41547

        Case GRIDOP_ALMOXARIFADO_COL
            lErro = Saida_Celula_Almoxarifado(objGridInt)
            If lErro <> SUCESSO Then Error 41548

        Case GRIDOP_CCL_COL
            lErro = Saida_Celula_Ccl(objGridInt)
            If lErro <> SUCESSO Then Error 41549

        Case GRIDOP_DATAPREVISAOINICIO_COL
            lErro = Saida_Celula_DataPrevInicio(objGridInt)
            If lErro <> SUCESSO Then Error 41550

        Case GRIDOP_DATAPREVISAOFIM_COL
            lErro = Saida_Celula_DataPrevFim(objGridInt)
            If lErro <> SUCESSO Then Error 41551

        Case GRIDOP_PRIORIDADE_COL
            lErro = Saida_Celula_Prioridade(objGridInt)
            If lErro <> SUCESSO Then Error 41552

    End Select

    Saida_Celula_OP = SUCESSO

    Exit Function

Erro_Saida_Celula_OP:

    Saida_Celula_OP = Err

    Select Case Err

        Case 41547 To 41552

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160903)

    End Select

    Exit Function

End Function

Private Sub CheckExibePedidoOP_Click()
    iTabPrincipalAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PedidoFinal_GotFocus()

Dim iTabAux As Integer

    iTabAux = iTabPrincipalAlterado
    Call MaskEdBox_TrataGotFocus(PedidoFinal, iAlterado)
    iTabPrincipalAlterado = iTabAux
    
End Sub

Private Sub PedidoInicial_GotFocus()

Dim iTabAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    Call MaskEdBox_TrataGotFocus(PedidoInicial, iAlterado)
    iTabPrincipalAlterado = iTabAux

End Sub

Private Sub ProdutoOP_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridOP)
End Sub

Private Sub ProdutoOP_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOP)
End Sub

Private Sub ProdutoOP_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOP.objControle = ProdutoOP
    lErro = Grid_Campo_Libera_Foco(objGridOP)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer

On Error GoTo Erro_TabStrip1_BeforeClick

    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
        
    If (Cancel = False) Then
    
        'Se Frame atual é a seleção
        If TabStrip1.SelectedItem.Index = TABSELECAO Then

            If iTabPrincipalAlterado = REGISTRO_ALTERADO Then
                
                Call Grid_Limpa(objGridPV)
                
                lErro = Trata_TabPedidos()
                If lErro <> SUCESSO Then Error 41455

            End If
        
        End If

        'se o frame atual é a seleção ou o pedido
        If TabStrip1.SelectedItem.Index = TABSELECAO Or TabStrip1.SelectedItem.Index = TABPV Then

            If iTabPrincipalAlterado = REGISTRO_ALTERADO Or iTabPedidoAlterado = REGISTRO_ALTERADO Then
   
'                For iIndice = 1 To objGridPV.iLinhasExistentes
'
'                    If GridPV.TextMatrix(iIndice, GRIDPV_GERAOP_COL) = 1 Then
'
'                        For iIndice1 = iIndice + 1 To objGridPV.iLinhasExistentes
'
'                            'se existe mais de um item marcado com o mesmo produto ==> erro
'                            If GridPV.TextMatrix(iIndice1, GRIDPV_GERAOP_COL) = 1 And GridPV.TextMatrix(iIndice, GRIDPV_PRODUTO_COL) = GridPV.TextMatrix(iIndice1, GRIDPV_PRODUTO_COL) Then Error 46419
'                        Next
'                    End If
'                Next

                gbTrazendoDados = True
                        
                Call Grid_Limpa(objGridOP)
                
                lErro = Trata_TabOP()
                If lErro <> SUCESSO Then Error 41588
            
                gbTrazendoDados = False
            
                iTabPrincipalAlterado = 0
                iTabPedidoAlterado = 0
                
            End If

        End If
        
    End If
        
    Exit Sub
        
Erro_TabStrip1_BeforeClick:

    Cancel = True

    Select Case Err

        Case 41455, 41588

        Case 46419
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_REPETIDO", Err, GridPV.TextMatrix(iIndice, GRIDPV_PRODUTO_COL))

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160904)

    End Select

    Exit Sub

End Sub

Private Sub UMOP_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridOP)
End Sub

Private Sub UMOP_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOP)
End Sub

Private Sub UMOP_Validate(Cancel As Boolean)
    
Dim lErro As Long

    Set objGridOP.objControle = UMOP
    lErro = Grid_Campo_Libera_Foco(objGridOP)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub QuantOP_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub QuantOP_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridOP)
End Sub

Private Sub QuantOP_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOP)
End Sub

Private Sub QuantOP_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOP.objControle = QuantOP
    lErro = Grid_Campo_Libera_Foco(objGridOP)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantFalta_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridOP)
End Sub

Private Sub QuantFalta_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOP)
End Sub

Private Sub QuantFalta_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOP.objControle = QuantFalta
    lErro = Grid_Campo_Libera_Foco(objGridOP)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Almoxarifado_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Almoxarifado_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridOP)
End Sub

Private Sub Almoxarifado_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOP)
End Sub

Private Sub Almoxarifado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOP.objControle = Almoxarifado
    lErro = Grid_Campo_Libera_Foco(objGridOP)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Ccl_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Ccl_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridOP)
End Sub

Private Sub Ccl_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOP)
End Sub

Private Sub Ccl_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOP.objControle = Ccl
    lErro = Grid_Campo_Libera_Foco(objGridOP)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DescricaoItemOP_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridOP)
End Sub

Private Sub DescricaoItemOP_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOP)
End Sub

Private Sub DescricaoItemOP_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOP.objControle = DescricaoItemOP
    lErro = Grid_Campo_Libera_Foco(objGridOP)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DataPrevisaoInicio_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataPrevisaoInicio_GotFocus()
    Set objGridOP.objControle = DataPrevisaoInicio
    Call Grid_Campo_Recebe_Foco(objGridOP)
End Sub

Private Sub DataPrevisaoInicio_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOP)
End Sub

Private Sub DataPrevisaoInicio_Validate(Cancel As Boolean)

Dim lErro As Long

    lErro = Grid_Campo_Libera_Foco(objGridOP)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DataPrevisaoFim_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataPrevisaoFim_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridOP)
End Sub

Private Sub DataPrevisaoFim_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOP)
End Sub

Private Sub DataPrevisaoFim_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOP.objControle = DataPrevisaoFim
    lErro = Grid_Campo_Libera_Foco(objGridOP)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub FilialPVOP_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridOP)
End Sub

Private Sub FilialPVOP_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOP)
End Sub

Private Sub FilialPVOP_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOP.objControle = FilialPVOP
    lErro = Grid_Campo_Libera_Foco(objGridOP)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub PedidoDeVendaOP_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridOP)
End Sub

Private Sub PedidoDeVendaOP_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOP)
End Sub

Private Sub PedidoDeVendaOP_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOP.objControle = PedidoDeVendaOP
    lErro = Grid_Campo_Libera_Foco(objGridOP)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Prioridade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Prioridade_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridOP)
End Sub

Private Sub Prioridade_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOP)
End Sub

Private Sub Prioridade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOP.objControle = Prioridade
    lErro = Grid_Campo_Libera_Foco(objGridOP)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_QuantOP(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_QuantOP

    Set objGridInt.objControle = QuantOP

    'se a quantidade foi preenchida
    If Len(Trim(QuantOP.ClipText)) > 0 Then

        lErro = Valor_Positivo_Critica(QuantOP.Text)
        If lErro <> SUCESSO Then Error 41556

        'If CDbl(QuantOP.Text) > CDbl(GridOP.TextMatrix(GridOP.Row, GRIDOP_QUANTFALTA_COL)) Then Error 41643
        
        QuantOP.Text = Formata_Estoque(QuantOP.Text)

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 41557

    Saida_Celula_QuantOP = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantOP:

    Saida_Celula_QuantOP = Err

    Select Case Err

        Case 41556
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 41557
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 41643
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTOP_MAIOR_QUANTFALTA", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160905)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Almoxarifado(objGridInt As AdmGrid) As Long
'faz a critica da celula de produto do grid que está deixando de ser a corrente
Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Almoxarifado

    Set objGridInt.objControle = Almoxarifado

    If Len(Trim(Almoxarifado.ClipText)) <> 0 Then

        lErro = CF("Produto_Formata", GridOP.TextMatrix(GridOP.Row, GRIDOP_PRODUTO_COL), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then Error 41558

        'se produto estiver preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            'verifica almoxarifado
            lErro = TP_Almoxarifado_Filial_Produto_Grid(sProdutoFormatado, Almoxarifado, objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25157 And lErro <> 25162 Then Error 41559

            If lErro = 25157 Then Error 41560

            If lErro = 25162 Then Error 41561

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 41562

    Saida_Celula_Almoxarifado = SUCESSO

    Exit Function

Erro_Saida_Celula_Almoxarifado:

    Saida_Celula_Almoxarifado = Err

    Select Case Err

        Case 41558, 41559, 41562
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 41560

            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_INEXISTENTE", Almoxarifado.Text)

            If vbMsg = vbYes Then

                objAlmoxarifado.sNomeReduzido = Almoxarifado.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 41561

            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_INEXISTENTE1", CInt(Almoxarifado.Text))

            If vbMsg = vbYes Then

                objAlmoxarifado.iCodigo = CInt(Almoxarifado.Text)

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)


            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160906)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Ccl(objGridInt As AdmGrid) As Long
'faz a critica da celula de produto do grid que está deixando de ser a corrente

Dim lErro As Long, sCclFormatada As String
Dim objCcl As New ClassCcl

On Error GoTo Erro_Saida_Celula_Ccl

    Set objGridInt.objControle = Ccl

    If Len(Trim(Ccl.ClipText)) > 0 Then

        lErro = CF("Ccl_Critica", Ccl.Text, sCclFormatada, objCcl)
        If lErro <> SUCESSO And lErro <> 5703 Then Error 41563

        If lErro = 5703 Then Error 41564

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 41565

    Saida_Celula_Ccl = SUCESSO

    Exit Function

Erro_Saida_Celula_Ccl:

    Saida_Celula_Ccl = Err

    Select Case Err

        Case 41563, 41565
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 41564
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", Err, Ccl.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160907)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataPrevInicio(objGridInt As AdmGrid) As Long
'faz a critica da celula de DataPrevisaoInicio do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DataPrevInicio

    Set objGridInt.objControle = DataPrevisaoInicio

    'verifica se a data está preenchida
    If Len(Trim(DataPrevisaoInicio.ClipText)) > 0 Then

        'verifica se a data é válida
        lErro = Data_Critica(DataPrevisaoInicio.Text)
        If lErro <> SUCESSO Then Error 41566

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 41567

    Saida_Celula_DataPrevInicio = SUCESSO

    Exit Function

Erro_Saida_Celula_DataPrevInicio:

    Saida_Celula_DataPrevInicio = Err

    Select Case Err

        Case 41566, 41567
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160908)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataPrevFim(objGridInt As AdmGrid) As Long
'faz a critica da celula de dataprevisaoFim do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DataPrevFim

    Set objGridInt.objControle = DataPrevisaoFim

    'verifica se a data está preenchida
    If Len(Trim(DataPrevisaoFim.ClipText)) > 0 Then

        'verifica se a data é válida
        lErro = Data_Critica(DataPrevisaoFim.Text)
        If lErro <> SUCESSO Then Error 41568

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 41569

    Saida_Celula_DataPrevFim = SUCESSO

    Exit Function

Erro_Saida_Celula_DataPrevFim:

    Saida_Celula_DataPrevFim = Err

    Select Case Err

        Case 41568, 41569
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160909)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Prioridade(objGridInt As AdmGrid) As Long
'faz a critica da celula de Prioridade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Prioridade

    Set objGridInt.objControle = Prioridade

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 41570

    Saida_Celula_Prioridade = SUCESSO

    Exit Function

Erro_Saida_Celula_Prioridade:

    Saida_Celula_Prioridade = Err

    Select Case Err

        Case 41570
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160910)

    End Select

    Exit Function

End Function

Private Sub GridPV_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridPV, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPV, iAlterado)
    End If

End Sub

Private Sub GridPV_EnterCell()
    Call Grid_Entrada_Celula(objGridPV, iAlterado)
End Sub

Private Sub GridPV_GotFocus()
    Call Grid_Recebe_Foco(objGridPV)
End Sub

Private Sub GridPV_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridPV)
End Sub

Private Sub GridPV_KeyPress(KeyAscii As Integer)
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridPV, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPV, iAlterado)
    End If

End Sub

Private Sub GridPV_LeaveCell()
    Call Saida_Celula(objGridPV)
End Sub

Private Sub GridPV_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridPV)
    
End Sub

Private Sub GridPV_RowColChange()
    Call Grid_RowColChange(objGridPV)
End Sub

Private Sub GridPV_Scroll()
    Call Grid_Scroll(objGridPV)
End Sub

Private Function Inicializa_GridPV(objGridInt As AdmGrid) As Long
'Inicializa o GridPV
Dim lErro  As Long

On Error GoTo Erro_Inicializa_GridPV

    'Form do GridPV
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Gera OP")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Data Emissão")
    objGridInt.colColuna.Add ("Data Entrega")
    objGridInt.colColuna.Add ("PV")
    objGridInt.colColuna.Add ("Filial do Pedido")

    'Controles que participam do GridPV
    objGridInt.colCampo.Add (GeraOP.Name)
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoItem.Name)
    objGridInt.colCampo.Add (UM.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (Cliente.Name)
    objGridInt.colCampo.Add (Filial.Name)
    objGridInt.colCampo.Add (DataEmissao.Name)
    objGridInt.colCampo.Add (DataEntrega.Name)
    objGridInt.colCampo.Add (Pedido.Name)
    objGridInt.colCampo.Add (FilialPV.Name)

    'Grid do GridInterno
    objGridInt.objGrid = GridPV

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 12

    'Largura da primeira coluna
    GridPV.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridPV = SUCESSO

    Exit Function

Erro_Inicializa_GridPV:

    Inicializa_GridPV = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160911)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_PV(objGridInt As AdmGrid) As Long
''Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_PV

    Select Case GridPV.Col

        Case GRIDPV_GERAOP_COL

            lErro = Saida_Celula_GeraOP(objGridInt)
            If lErro <> SUCESSO Then Error 41571

        Case GRIDPV_QUANTIDADE_COL

            lErro = Saida_Celula_Quantidade(objGridInt)
            If lErro <> SUCESSO Then Error 41572

    End Select

    Saida_Celula_PV = SUCESSO

    Exit Function

Erro_Saida_Celula_PV:

    Saida_Celula_PV = Err

    Select Case Err

        Case 41571 To 41572
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160912)

    End Select

    Exit Function

End Function

Private Sub GeraOP_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPV)
End Sub

Private Sub GeraOP_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPV)
End Sub

Private Sub GeraOP_Validate(Cancel As Boolean)
    
Dim lErro As Long

    Set objGridPV.objControle = GeraOP
    lErro = Grid_Campo_Libera_Foco(objGridPV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Produto_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPV)
End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPV)
End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridPV.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGridPV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UM_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPV)
End Sub

Private Sub UM_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPV)
End Sub

Private Sub UM_Validate(Cancel As Boolean)
    
Dim lErro As Long

    Set objGridPV.objControle = UM
    lErro = Grid_Campo_Libera_Foco(objGridPV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Quantidade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Quantidade_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPV)
End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPV)
End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridPV.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGridPV)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub DescricaoItem_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPV)
End Sub

Private Sub DescricaoItem_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPV)
End Sub

Private Sub DescricaoItem_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridPV.objControle = DescricaoItem
    lErro = Grid_Campo_Libera_Foco(objGridPV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Cliente_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPV)
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPV)
End Sub

Private Sub Cliente_Validate(Cancel As Boolean)
    
Dim lErro As Long

    Set objGridPV.objControle = Cliente
    lErro = Grid_Campo_Libera_Foco(objGridPV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Filial_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPV)
End Sub

Private Sub Filial_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPV)
End Sub

Private Sub Filial_Validate(Cancel As Boolean)
    
Dim lErro As Long

    Set objGridPV.objControle = Filial
    lErro = Grid_Campo_Libera_Foco(objGridPV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DataEmissao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPV)
End Sub

Private Sub DataEmissao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPV)
End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)
    
Dim lErro As Long

    Set objGridPV.objControle = DataEmissao
    lErro = Grid_Campo_Libera_Foco(objGridPV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DataEntrega_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPV)
End Sub

Private Sub DataEntrega_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPV)
End Sub

Private Sub DataEntrega_Validate(Cancel As Boolean)
    
Dim lErro As Long

    Set objGridPV.objControle = DataEntrega
    lErro = Grid_Campo_Libera_Foco(objGridPV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub FilialPV_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPV)
End Sub

Private Sub FilialPV_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPV)
End Sub

Private Sub FilialPV_Validate(Cancel As Boolean)
    
Dim lErro As Long

    Set objGridPV.objControle = FilialPV
    lErro = Grid_Campo_Libera_Foco(objGridPV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Pedido_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPV)
End Sub

Private Sub Pedido_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPV)
End Sub

Private Sub Pedido_Validate(Cancel As Boolean)
    
Dim lErro As Long

    Set objGridPV.objControle = Pedido
    lErro = Grid_Campo_Libera_Foco(objGridPV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_GeraOP(objGridInt As AdmGrid) As Long
'faz a critica da celula GeraOP do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_GeraOP

    Set objGridInt.objControle = GeraOP
        
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 41576

    Saida_Celula_GeraOP = SUCESSO

    Exit Function

Erro_Saida_Celula_GeraOP:

    Saida_Celula_GeraOP = Err

    Select Case Err

        Case 41576
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160913)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantTotal As Double

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    'se a quantidade foi preenchida
    If Len(Trim(Quantidade.ClipText)) > 0 Then

        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then Error 41577

        Quantidade.Text = Formata_Estoque(Quantidade.Text)

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 41578

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = Err

    Select Case Err

        Case 41577
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 41578
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160914)

    End Select

    Exit Function

End Function

Private Sub GeraOP_Click()

Dim iClick As Integer
Dim lErro As Long

On Error GoTo Erro_GeraOP_Click

    iAlterado = REGISTRO_ALTERADO
    iTabPedidoAlterado = REGISTRO_ALTERADO
    
    'Verifica se é alguma linha válida
    If GridPV.Row > objGridPV.iLinhasExistentes Then Exit Sub

    'Verifica se está selecionando ou desselecionando
    If Len(Trim(GridPV.TextMatrix(GridPV.Row, GRIDPV_GERAOP_COL))) > 0 Then
        iClick = CInt(GridPV.TextMatrix(GridPV.Row, GRIDPV_GERAOP_COL)) = 1
    End If

    If iClick = True Then
        gobjGeracaoOP.colItens(GridPV.Row).iSelecionado = 1
        iSelecionadosPV = iSelecionadosPV + 1
    Else
        gobjGeracaoOP.colItens(GridPV.Row).iSelecionado = 0
        iSelecionadosPV = iSelecionadosPV - 1
    End If

    ItensSelecionados.Caption = CStr(iSelecionadosPV)

    gobjGeracaoOP.iNumItensSelecionados = iSelecionadosPV

    lErro = Grid_Refresh_Checkbox(objGridPV)
    If lErro <> SUCESSO Then Error 41579

    Exit Sub

Erro_GeraOP_Click:

    Select Case Err

        Case 41579

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160915)

    End Select

    Exit Sub

End Sub

Private Sub Troca_Enabled(Value As Boolean)

    PedidoInicial.Enabled = Value
    PedidoFinal.Enabled = Value
    ClienteDe.Enabled = Value
    ClienteAte.Enabled = Value
    DataEntregaDe.Enabled = Value
    DataEntregaAte.Enabled = Value
    ProdutoDe.Enabled = Value
    ProdutoAte.Enabled = Value
    UpDownEntregaDe.Enabled = Value
    UpDownEntregaAte.Enabled = Value

End Sub

Private Sub Inicializa_Campos()

    PedidoInicial.Text = ""
    PedidoFinal.Text = ""
    ClienteDe.Text = ""
    ClienteAte.Text = ""

    DataEntregaDe.PromptInclude = False
    DataEntregaDe.Text = ""
    DataEntregaDe.PromptInclude = True

    DataEntregaAte.PromptInclude = False
    DataEntregaAte.Text = ""
    DataEntregaAte.PromptInclude = True

    ProdutoDe.PromptInclude = False
    ProdutoDe.Text = ""
    ProdutoDe.PromptInclude = True

    ProdutoAte.PromptInclude = False
    ProdutoAte.Text = ""
    ProdutoAte.PromptInclude = True

End Sub

Private Sub BotaoCcls_Click()
'chama tela de Lista de Ccl

Dim objCcls As New ClassCcl
Dim colSelecao As New Collection

    Call Chama_Tela("CclLista", colSelecao, objCcls, objEventoCcl)

End Sub

Private Sub objEventoCcl_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclMascarado As String
Dim sCclFormatada As String

On Error GoTo Erro_objEventoCcl_evSelecao

    Set objCcl = obj1

    'Se o produto da linha corrente estiver preenchido e Linha corrente diferente da Linha fixa
    If Len(Trim(GridOP.TextMatrix(GridOP.Row, GRIDOP_PRODUTO_COL))) <> 0 And GridOP.Row <> 0 Then

        sCclMascarado = String(STRING_CCL, 0)

        lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then Error 41580

        'Coloca o valor do Ccl na coluna correspondente
        GridOP.TextMatrix(GridOP.Row, GRIDOP_CCL_COL) = sCclMascarado

        Ccl.PromptInclude = False
        Ccl.Text = sCclMascarado
        Ccl.PromptInclude = True

    End If

    Me.Show

    Exit Sub

Erro_objEventoCcl_evSelecao:

    Select Case Err

        Case 41580

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160916)

    End Select

    Exit Sub

End Sub

Private Sub ClienteAte_Change()
    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Data_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Data_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    'Se a Data está preenchida
    If Len(Data.ClipText) = 0 Then Exit Sub

    'Verifica se a Data é válida
    lErro = Data_Critica(Data.Text)
    If lErro <> SUCESSO Then Error 41581

    Exit Sub

Erro_Data_Validate:

    Cancel = True


    Select Case Err

        Case 41581

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160917)

    End Select

    Exit Sub

End Sub

Private Sub DataEntregaAte_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub DataEntregaAte_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataEntregaAte_Validate

    'Se a DataEmissaoDe está preenchida
    If Len(DataEntregaAte.ClipText) = 0 Then Exit Sub

    'Verifica se a DataEmissaoDe é válida
    lErro = Data_Critica(DataEntregaAte.Text)
    If lErro <> SUCESSO Then Error 41582

    Exit Sub

Erro_DataEntregaAte_Validate:

    Cancel = True


    Select Case Err

        Case 41582

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160918)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoAte_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProdutoAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sProdutoAteFormatado As String
Dim iProdutoAtePreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_ProdutoDe_Validate

    lErro = CF("Produto_Formata", ProdutoAte.Text, sProdutoAteFormatado, iProdutoAtePreenchido)
    If lErro <> SUCESSO Then Error 41447

    If iProdutoAtePreenchido = PRODUTO_PREENCHIDO Then

        objProduto.sCodigo = sProdutoAteFormatado

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 41584

        If lErro = 28030 Then Error 41585

    End If

    Exit Sub

Erro_ProdutoDe_Validate:

    Cancel = True


    Select Case Err

        Case 41584

        Case 41585
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)

        Case 41447

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160919)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoDe_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProdutoDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sProdutoDeFormatado As String, sProdutoAteFormatado As String
Dim iProdutoDePreenchido As Integer, iProdutoAtePreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_ProdutoDe_Validate

    lErro = CF("Produto_Formata", ProdutoDe.Text, sProdutoDeFormatado, iProdutoDePreenchido)
    If lErro <> SUCESSO Then Error 41443

    If iProdutoDePreenchido = PRODUTO_PREENCHIDO Then

        objProduto.sCodigo = sProdutoDeFormatado

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 41586

        If lErro = 28030 Then Error 41587

    End If

    Exit Sub

Erro_ProdutoDe_Validate:

    Cancel = True


    Select Case Err

        Case 41443

        Case 41586

        Case 41587
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160920)

    End Select

    Exit Sub

End Sub

Private Sub ClienteDe_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub PedidoFinal_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PedidoInicial_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEntregaDe_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEntregaDe_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataEntregaDe_Validate

    'Se a DataEmissaoDe está preenchida
    If Len(DataEntregaDe.ClipText) = 0 Then Exit Sub

    'Verifica se a DataEmissaoDe é válida
    lErro = Data_Critica(DataEntregaDe.Text)
    If lErro <> SUCESSO Then Error 41438

    Exit Sub

Erro_DataEntregaDe_Validate:

    Cancel = True


    Select Case Err

        Case 41438

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160921)

    End Select

    Exit Sub

End Sub

Private Sub LabelPedidoDe_Click()

Dim objPedidoDeVenda As New ClassPedidoDeVenda
Dim colSelecao As New Collection

    'Preenche PedidoDe com o pedido da tela
    If Len(Trim(PedidoInicial.ClipText)) > 0 Then objPedidoDeVenda.lCodigo = CLng(PedidoInicial.Text)

    objPedidoDeVenda.iFilialEmpresa = giFilialEmpresa

    'Chama Tela PedidoVendaLista
    Call Chama_Tela("PedidoVendaLista", colSelecao, objPedidoDeVenda, objEventoPedidoDe)

End Sub

Private Sub objEventoPedidoDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPedidoDeVenda As ClassPedidoDeVenda

On Error GoTo Erro_objEventoPedidoDe_evSelecao

    Set objPedidoDeVenda = obj1

    If ExibeTodos.Value = False Then

        PedidoInicial.Text = CStr(objPedidoDeVenda.lCodigo)

    End If

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoPedidoDe_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160922)

    End Select

    Exit Sub

End Sub

Private Sub LabelPedidoAte_Click()

Dim objPedidoDeVenda As New ClassPedidoDeVenda
Dim colSelecao As Collection

    'Preenche PedidoAte com o pedido da tela
    If Len(Trim(PedidoFinal.ClipText)) > 0 Then objPedidoDeVenda.lCodigo = CLng(PedidoFinal.Text)

    objPedidoDeVenda.iFilialEmpresa = giFilialEmpresa

    'Chama Tela PedidoVendaLista
    Call Chama_Tela("PedidoVendaLista", colSelecao, objPedidoDeVenda, objEventoPedidoAte)

End Sub

Private Sub objEventoPedidoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPedidoDeVenda As ClassPedidoDeVenda

On Error GoTo Erro_objEventoPedidoAte_evSelecao

    Set objPedidoDeVenda = obj1

    If ExibeTodos.Value = False Then

        PedidoFinal.Text = CStr(objPedidoDeVenda.lCodigo)

    End If

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoPedidoAte_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160923)

    End Select

    Exit Sub

End Sub

Private Sub LabelClienteDe_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    'Preenche ClienteDe com o cliente da tela
    If Len(Trim(ClienteDe.ClipText)) > 0 Then objCliente.lCodigo = CLng(ClienteDe.Text)

    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteDe)

End Sub

Private Sub objEventoClienteDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCliente As ClassCliente

On Error GoTo Erro_objEventoClienteDe_evSelecao

    Set objCliente = obj1

    If ExibeTodos.Value = False Then

        ClienteDe.Text = CStr(objCliente.lCodigo)

    End If

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoClienteDe_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160924)

    End Select

    Exit Sub

End Sub

Private Sub LabelClienteAte_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    'Preenche ClienteAte com o cliente da tela
    If Len(Trim(ClienteAte.ClipText)) > 0 Then objCliente.lCodigo = CLng(ClienteAte.Text)

    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteAte)

End Sub

Private Sub objEventoClienteAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCliente As ClassCliente

On Error GoTo Erro_objEventoClienteAte_evSelecao

    Set objCliente = obj1

    If ExibeTodos.Value = False Then

        ClienteAte.Text = CStr(objCliente.lCodigo)

    End If

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoClienteAte_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160925)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    'Aumenta a DataEntregaDe em 1 dia
    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 41458

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case Err

        Case 41458

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160926)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    'Aumenta a DataEntregaDe em 1 dia
    lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 41459

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case Err

        Case 41459

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160927)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEntregaAte_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownEntregaAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEntregaAte_DownClick

    'Diminui a DataEntregaDe em 1 dia
    lErro = Data_Up_Down_Click(DataEntregaAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 41456

    Exit Sub

Erro_UpDownEntregaAte_DownClick:

    Select Case Err

        Case 41456

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160928)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEntregaAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEntregaDe_UpClick

    'Aumenta a DataEntregaDe em 1 dia
    lErro = Data_Up_Down_Click(DataEntregaAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 41457

    Exit Sub

Erro_UpDownEntregaDe_UpClick:

    Select Case Err

        Case 41457

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160929)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEntregaDe_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownEntregaDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEntregaDe_DownClick

    'Diminui a DataEntregaDe em 1 dia
    lErro = Data_Up_Down_Click(DataEntregaDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 41449

    Exit Sub

Erro_UpDownEntregaDe_DownClick:

    Select Case Err

        Case 41449

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160930)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEntregaDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEntregaDe_UpClick

    'Aumenta a DataEntregaDe em 1 dia
    lErro = Data_Up_Down_Click(DataEntregaDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 41450

    Exit Sub

Erro_UpDownEntregaDe_UpClick:

    Select Case Err

        Case 41450

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160931)

    End Select

    Exit Sub

End Sub

Private Sub LabelProdutoDe_Click()

Dim objProduto As New ClassProduto
Dim colSelecao As Collection
Dim sProdutoDe As String
Dim iProdutoDePreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_LabelProdutoDe_Click

    lErro = CF("Produto_Formata", ProdutoDe.Text, sProdutoDe, iProdutoDePreenchido)
    If lErro <> SUCESSO Then Error 41451

    'Preenche Produtocom o codigo da tela
    If iProdutoDePreenchido = PRODUTO_PREENCHIDO Then objProduto.sCodigo = sProdutoDe

    'Chama Tela ProdutoLista
    Call Chama_Tela("ProdutoLista_Consulta", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case Err

        Case 41451

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160932)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoDe_evSelecao(obj1 As Object)

Dim objProduto As ClassProduto
Dim lErro As Long
Dim sProduto As String

On Error GoTo Erro_objEventoProdutoDe_evSelecao

    Set objProduto = obj1

    lErro = Mascara_MascararProduto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then Error 41452

    If ExibeTodos.Value = False Then
        
        ProdutoDe.Text = sProduto
        Call ProdutoDe_Validate(bSGECancelDummy)
        
    End If

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case Err

        Case 41452

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160933)

    End Select

    Exit Sub

End Sub

Private Sub LabelProdutoAte_Click()

Dim objProduto As New ClassProduto
Dim colSelecao As Collection
Dim sProdutoAte As String
Dim iProdutoAtePreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_LabelProdutoDe_Click

    lErro = CF("Produto_Formata", ProdutoAte.Text, sProdutoAte, iProdutoAtePreenchido)
    If lErro <> SUCESSO Then Error 41453

    'Preenche Produtocom o codigo da tela
    If iProdutoAtePreenchido = PRODUTO_PREENCHIDO Then objProduto.sCodigo = sProdutoAte

    'Chama Tela ProdutoLista
    Call Chama_Tela("ProdutoLista_Consulta", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case Err

        Case 41453

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160934)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim objProduto As ClassProduto
Dim lErro As Long
Dim sProduto As String

On Error GoTo Erro_objEventoProdutoDe_evSelecao

    Set objProduto = obj1

    lErro = Mascara_MascararProduto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then Error 41454

    If ExibeTodos.Value = False Then
        
        ProdutoAte.Text = sProduto
        Call ProdutoAte_Validate(bSGECancelDummy)
        
    End If

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case Err

        Case 41454

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160935)

    End Select

    Exit Sub

End Sub

Private Sub Testa_Selecionados()

    If ExibeTodos.Value = False Then

        If Len(Trim(PedidoInicial.ClipText)) = 0 And Len(Trim(PedidoFinal.ClipText)) = 0 Then
            If Len(Trim(ClienteDe.ClipText)) = 0 And Len(Trim(ClienteAte.ClipText)) = 0 Then
                If Len(Trim(DataEntregaDe.ClipText)) = 0 And Len(Trim(DataEntregaAte.ClipText)) = 0 Then
                    iTabPrincipalAlterado = 0
                    iSelecionadosPV = 0
                End If
            End If
        End If

    Else
        iTabPrincipalAlterado = REGISTRO_ALTERADO
    End If

End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click

    'Se Frame atual não corresponde ao Tab clicado
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub


        'Torna Frame atual invisível
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index
        'Torna Frame de Pedido visível
        Frame1(TabStrip1.SelectedItem.Index).Visible = True

        Select Case iFrameAtual
        
            Case TAB_Selecao
                Parent.HelpContextID = IDH_GERACAO_OP_SELECAO
                
            Case TAB_PV
                Parent.HelpContextID = IDH_GERACAO_OP_PV
                        
            Case TAB_OP_GERADA
                Parent.HelpContextID = IDH_GERACAO_OP_GERADA
                
        End Select
    
    
    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160936)

    End Select

    Exit Sub

End Sub

Private Sub BotaoMarcarTodos_Click()
'Marca todos os pedidos do Grid

Dim iLinha As Integer

    iAlterado = REGISTRO_ALTERADO
    iTabPedidoAlterado = REGISTRO_ALTERADO

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridPV.iLinhasExistentes

        'Marca na tela o pedido em questão
        GridPV.TextMatrix(iLinha, GRIDPV_GERAOP_COL) = S_MARCADO

        gobjGeracaoOP.colItens.Item(iLinha).iSelecionado = S_MARCADO

    Next

    iSelecionadosPV = iLinha - 1

    ItensSelecionados.Caption = CStr(iSelecionadosPV)

    gobjGeracaoOP.iNumItensSelecionados = iSelecionadosPV

    'Atualiza na tela a checkbox marcada
    Call Grid_Refresh_Checkbox(objGridPV)

End Sub

Private Sub BotaoDesmarcarTodos_Click()
'Desmarca todos os pedidos do Grid

Dim iLinha As Integer

    iAlterado = REGISTRO_ALTERADO
    iTabPedidoAlterado = REGISTRO_ALTERADO

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridPV.iLinhasExistentes

        'Desmarca na tela o pedido em questão
        GridPV.TextMatrix(iLinha, GRIDPV_GERAOP_COL) = S_DESMARCADO

        'Desmarca no Obj o pedido em questão
        gobjGeracaoOP.colItens.Item(iLinha).iSelecionado = S_DESMARCADO

    Next

    ItensSelecionados.Caption = "0"

    iSelecionadosPV = 0

    gobjGeracaoOP.iNumItensSelecionados = 0

    'Atualiza na tela a checkbox desmarcada
    Call Grid_Refresh_Checkbox(objGridPV)

End Sub

Private Sub BotaoFechar_Click()

    'Fecha a tela
    Unload Me

End Sub

Private Sub Ordenados_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Ordenados_Click()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Ordenados_Click

    If Ordenados.ListIndex = -1 Then Exit Sub

    'Verifica se a coleção de OP está vazia
    If gobjGeracaoOP.colItens.Count = 0 Then Exit Sub

    'Passa a Ordenaçao escolhida para o Obj
    gobjGeracaoOP.sOrdenacao = asOrdenacao(Ordenados.ListIndex)

    'Limpa a coleção
    Set gobjGeracaoOP = New ClassGeracaoOP

    lErro = Move_TabSelecao_Memoria()
    If lErro <> SUCESSO Then Error 55875

    'Preenche a Coleção
    lErro = CF("GeracaoOP_SelecionarItens", gobjGeracaoOP)
    If lErro <> SUCESSO Then Error 41589

    'Limpa o GridPedido
    Call Grid_Limpa(objGridPV)

    'Preenche o GridPedido
    Call GridPV_Preenche(gobjGeracaoOP.colItens)

    Exit Sub

Erro_Ordenados_Click:

    Select Case Err

        Case 41589, 55875

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160937)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objGridPV = Nothing
    Set objGridItemPV = Nothing
    Set objGridOP = Nothing

    Set objEventoPedidoDe = Nothing
    Set objEventoPedidoAte = Nothing
    Set objEventoClienteDe = Nothing
    Set objEventoClienteAte = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    Set objEventoCcl = Nothing
    Set objEventoCodigo = Nothing
    Set objEventoEstoque = Nothing
    Set gobjGeracaoOP = Nothing
    Set gobjOP = Nothing


End Sub

Private Sub GridPV_Preenche(colItens As Collection)
'Preenche o Grid Pedido com os dados de colItens

Dim iLinha As Integer
Dim objGeracaoOPItem As ClassGeracaoOPItem
Dim sProdutoMascarado As String
Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim objFilialCliente As New ClassFilialCliente
Dim sNomeRedCliente As String
Dim dFator As Double

On Error GoTo Erro_GridPV_Preenche

    'Se o número de itens for maior que o número de linhas do Grid
    
    If colItens.Count + 1 > GridPV.Rows Then
    
        'Altera o número de linhas do Grid de acordo com o número de itens
        GridPV.Rows = colItens.Count + 1

        'Chama rotina de Inicialização do Grid
        Call Grid_Inicializa(objGridPV)

    End If

    iLinha = 0

    'Percorre todas as OP da Coleção
    For Each objGeracaoOPItem In colItens

        iLinha = iLinha + 1

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        lErro = Mascara_MascararProduto(objGeracaoOPItem.objItemOP.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 41543

        'Passa para a tela os dados da OP em questão
        GridPV.TextMatrix(iLinha, GRIDPV_GERAOP_COL) = CStr(objGeracaoOPItem.iSelecionado)
        GridPV.TextMatrix(iLinha, GRIDPV_PRODUTO_COL) = sProdutoMascarado
        GridPV.TextMatrix(iLinha, GRIDPV_DESCRICAOITEM_COL) = objGeracaoOPItem.sProdutoDescricao
        GridPV.TextMatrix(iLinha, GRIDPV_UM_COL) = objGeracaoOPItem.objItemOP.sSiglaUMEstoque
        
        lErro = CF("UM_Conversao", objGeracaoOPItem.objItemOP.iClasseUM, objGeracaoOPItem.objItemOP.sSiglaUM, objGeracaoOPItem.objItemOP.sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 90766
         
        GridPV.TextMatrix(iLinha, GRIDPV_QUANTIDADE_COL) = Formata_Estoque(objGeracaoOPItem.dQtdeFalta * dFator)
        GridPV.TextMatrix(iLinha, GRIDPV_CLIENTE_COL) = objGeracaoOPItem.sClienteReduzido
        
        objFilialCliente.iCodFilial = objGeracaoOPItem.iFilialCliente
        sNomeRedCliente = objGeracaoOPItem.sClienteReduzido
        
        'le nome da filial cliente
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sNomeRedCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 49993
        
        If lErro = 17660 Then gError 49994
        
        'PREENCHE CODIGO-NOME DA FILIAL
        GridPV.TextMatrix(iLinha, GRIDPV_FILIAL_COL) = CStr(objFilialCliente.iCodFilial) & SEPARADOR & objFilialCliente.sNome
        
        If objGeracaoOPItem.dtEmissaoPV <> DATA_NULA And objGeracaoOPItem.dtEmissaoPV <> 0 Then GridPV.TextMatrix(iLinha, GRIDPV_DATAEMISSAO_COL) = Format(objGeracaoOPItem.dtEmissaoPV, "dd/mm/yyyy")
        If objGeracaoOPItem.dtEntregaItemPV <> DATA_NULA And objGeracaoOPItem.dtEntregaItemPV <> 0 Then GridPV.TextMatrix(iLinha, GRIDPV_DATAENTREGA_COL) = Format(objGeracaoOPItem.dtEntregaItemPV, "dd/mm/yyyy")
        
        'preenche o objFilialEmpresa
        objFilialEmpresa.iCodFilial = objGeracaoOPItem.objItemOP.iFilialPedido
        
        'Le o nome da filial
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 49985
        
        'se não encontrou a filial ----- ERRO
        If lErro = 27378 Then gError 49992
        
        'CODIGO-NOME
        GridPV.TextMatrix(iLinha, GRIDPV_FILIALPV_COL) = CStr(objFilialEmpresa.iCodFilial) & SEPARADOR & objFilialEmpresa.sNome
                        
        GridPV.TextMatrix(iLinha, GRIDPV_PEDIDO_COL) = CStr(objGeracaoOPItem.objItemOP.lCodPedido)

    Next

    Call Grid_Refresh_Checkbox(objGridPV)

    'Passa para o Obj o número de itens passados pela Coleção
    objGridPV.iLinhasExistentes = colItens.Count

    Call Totaliza_Selecionados(colItens.Count)

    Exit Sub

Erro_GridPV_Preenche:

    Select Case gErr

        Case 41543, 49985, 49993, 90766
        
        Case 49992
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)
            
        Case 49994
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA2", gErr, sNomeRedCliente, objFilialCliente.iCodFilial)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160938)

    End Select

    Exit Sub

End Sub

Private Sub Totaliza_Selecionados(iNumLinhas As Integer)

Dim iLinha As Integer
Dim iTotal As Integer

    iTotal = 0

    For iLinha = 1 To iNumLinhas

        If GridPV.TextMatrix(iLinha, GRIDPV_GERAOP_COL) = "1" Then iTotal = iTotal + 1

    Next

    iSelecionadosPV = iTotal

    ItensSelecionados.Caption = CStr(iTotal)

End Sub

Private Sub GridOP_Preenche(colItens As Collection)

Dim iLinha As Integer
Dim objGeracaoOPItem As ClassGeracaoOPItem
Dim sProdutoMascarado As String
Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim objProdutoFilial As New ClassProdutoFilial
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim dFator As Double
Dim objItemRomaneiGrade As ClassItemRomaneioGrade
Dim objItemRomaneiGradeAux As ClassItemRomaneioGrade
Dim objItemOP As ClassItemOP
Dim objOP As New ClassOrdemDeProducao
Dim bAchou As Boolean
Dim objItemOPItemPV As ClassItemOPItemPV
Dim iFilialEmpresaAnt As Integer
Dim objFilialCliente As ClassFilialCliente
Dim objItemPV As New ClassItemPedido
Dim colSaida As Collection
Dim colCampos As Collection

On Error GoTo Erro_GridOP_Preenche

    'Se o número de op's for maior que o número de linhas do Grid
    If colItens.Count + 1 > GridOP.Rows Then
        'Altera o número de linhas do Grid de acordo com o número de OP's
        GridOP.Rows = colItens.Count + 1

        'Chama rotina de Inicialização do Grid
        Call Grid_Inicializa(objGridOP)

    End If

    iLinha = 0
    
    'Percorre todas as OP da Coleção e agrupa por produtos as OP selecionadas
    For Each objGeracaoOPItem In colItens

        If objGeracaoOPItem.iSelecionado = MARCADO Then
        
            Set objFilialCliente = New ClassFilialCliente
              
            objFilialCliente.iCodFilial = objGeracaoOPItem.iFilialCliente
            
            'le nome da filial cliente
            lErro = CF("FilialCliente_Le_NomeRed_CodFilial", objGeracaoOPItem.sClienteReduzido, objFilialCliente)
            If lErro <> SUCESSO And lErro <> 17660 Then gError 49993
        
            lErro = CF("UM_Conversao", objGeracaoOPItem.objItemOP.iClasseUM, objGeracaoOPItem.objItemOP.sSiglaUM, objGeracaoOPItem.objItemOP.sSiglaUMEstoque, dFator)
            If lErro <> SUCESSO Then gError 90768
    
            bAchou = False
            For Each objItemOP In objOP.colItens
                'Se já tem o produto complementa a quantidade, senão insere
                If objItemOP.sProduto = objGeracaoOPItem.objItemOP.sProduto Then
                    objItemOP.dQuantidade = objItemOP.dQuantidade + (objGeracaoOPItem.dQtdeFalta * dFator)
                    bAchou = True
                    Exit For
                End If
            Next
            
            If objGeracaoOPItem.objItemOP.iFilialPedido <> iFilialEmpresaAnt Then
                
                Set objFilialEmpresa = New AdmFiliais
                
                iFilialEmpresaAnt = objGeracaoOPItem.objItemOP.iFilialPedido
            
                'preenche o objFilialEmpresa
                objFilialEmpresa.iCodFilial = objGeracaoOPItem.objItemOP.iFilialPedido
            
                'Le o nome da filial
                lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
                If lErro <> SUCESSO And lErro <> 27378 Then gError 49995
                
            End If
            
            'Se ainda não tinha esse item na OP adiciona
            If Not bAchou Then

                Set objItemOP = New ClassItemOP
                
                objItemOP.dQuantidade = objGeracaoOPItem.dQtdeFalta
                objItemOP.iFilialEmpresa = objGeracaoOPItem.objItemOP.iFilialPedido
                objItemOP.iItem = objOP.colItens.Count + 1
                objItemOP.iClasseUM = objGeracaoOPItem.objItemOP.iClasseUM
                objItemOP.iPrioridade = objGeracaoOPItem.iPrioridade
                objItemOP.dQuantidade = (objGeracaoOPItem.dQtdeFalta * dFator)
                objItemOP.sSiglaUM = objGeracaoOPItem.objItemOP.sSiglaUM
                objItemOP.sSiglaUMEstoque = objGeracaoOPItem.objItemOP.sSiglaUMEstoque
                objItemOP.sDescricao = objGeracaoOPItem.sProdutoDescricao
                objItemOP.sProduto = objGeracaoOPItem.objItemOP.sProduto
                
                objOP.colItens.Add objItemOP
            End If

            'Adiciona a informação dos pedidos de venda no item da OP
            'ou no romaneio
            Set objItemOPItemPV = New ClassItemOPItemPV
            
            objItemOPItemPV.dQuantidade = objGeracaoOPItem.dQtdeFalta * dFator
            objItemOPItemPV.dtDataEmissao = objGeracaoOPItem.dtEmissaoPV
            objItemOPItemPV.dtDataEntrega = objGeracaoOPItem.dtEntregaItemPV
            objItemOPItemPV.iFilial = objGeracaoOPItem.iFilialCliente
            objItemOPItemPV.iFilialEmpresa = objGeracaoOPItem.objItemOP.iFilialPedido
            objItemOPItemPV.lCliente = objFilialCliente.lCodCliente
            objItemOPItemPV.lCodigoPV = objGeracaoOPItem.objItemOP.lCodPedido
            objItemOPItemPV.sCodigoOP = objGeracaoOPItem.objItemOP.sCodigo
            objItemOPItemPV.sDescProduto = objGeracaoOPItem.sProdutoDescricao
            objItemOPItemPV.sNomeCliente = objGeracaoOPItem.sClienteReduzido
            objItemOPItemPV.sNomeFilialCliente = objFilialCliente.sNome
            objItemOPItemPV.sNomeFilialEmpresa = objFilialEmpresa.sNome
            objItemOPItemPV.sProduto = objGeracaoOPItem.objItemOP.sProduto
            objItemOPItemPV.sUM = objGeracaoOPItem.objItemOP.sSiglaUMEstoque
            objItemOPItemPV.iPrioridade = objGeracaoOPItem.iPrioridade
            
            'Le o item do pedido para pegar o NumIntDoc
            Set objItemPV = New ClassItemPedido
            objItemPV.lCodPedido = objItemOPItemPV.lCodigoPV
            objItemPV.iFilialEmpresa = objItemOPItemPV.iFilialEmpresa
            objItemPV.sProduto = objItemOPItemPV.sProduto
            
            lErro = CF("ItemPedido_Le", objItemPV)
            If lErro <> SUCESSO And lErro <> 23971 Then gError 49995
            
            objItemOPItemPV.lNumIntDocItemPV = objItemPV.lNumIntDoc
            
            objItemOP.colItemPV.Add objItemOPItemPV
                
            If objGeracaoOPItem.objItemOP.colItensRomaneioGrade.Count <> 0 Then
                'Busca para ver se o produto já está no romaneio
                For Each objItemRomaneiGrade In objGeracaoOPItem.objItemOP.colItensRomaneioGrade
                    bAchou = False
                    For Each objItemRomaneiGradeAux In objItemOP.colItensRomaneioGrade
                        If objItemRomaneiGrade.sProduto = objItemRomaneiGradeAux.sProduto Then
                            objItemRomaneiGradeAux.dQuantidade = objItemRomaneiGrade.dQuantidade + objItemRomaneiGradeAux.dQuantidade
                            objItemRomaneiGradeAux.dQuantOP = objItemRomaneiGrade.dQuantOP + objItemRomaneiGradeAux.dQuantOP
                            objItemRomaneiGradeAux.dQuantPV = objItemRomaneiGrade.dQuantPV + objItemRomaneiGradeAux.dQuantPV
                            objItemRomaneiGradeAux.dQuantReservada = objItemRomaneiGrade.dQuantReservada + objItemRomaneiGradeAux.dQuantReservada
                            objItemRomaneiGradeAux.dQuantSC = objItemRomaneiGrade.dQuantSC + objItemRomaneiGradeAux.dQuantSC
                            objItemRomaneiGradeAux.dQuantFaturada = objItemRomaneiGrade.dQuantFaturada + objItemRomaneiGradeAux.dQuantFaturada
                            objItemRomaneiGradeAux.dQuantCancelada = objItemRomaneiGrade.dQuantCancelada + objItemRomaneiGradeAux.dQuantCancelada
                            objItemRomaneiGradeAux.dQuantAFaturar = objItemRomaneiGrade.dQuantAFaturar + objItemRomaneiGradeAux.dQuantAFaturar
                            bAchou = True
                        End If
                    Next
                    'Se não achou insere o novo item de romaneio
                    If Not bAchou Then
                        
                        Set objItemRomaneiGradeAux = New ClassItemRomaneioGrade
                        
                        objItemRomaneiGradeAux.dQuantidade = objItemRomaneiGrade.dQuantidade
                        objItemRomaneiGradeAux.dQuantOP = objItemRomaneiGrade.dQuantOP
                        objItemRomaneiGradeAux.dQuantPV = objItemRomaneiGrade.dQuantPV
                        objItemRomaneiGradeAux.dQuantReservada = objItemRomaneiGrade.dQuantReservada
                        objItemRomaneiGradeAux.dQuantSC = objItemRomaneiGrade.dQuantSC
                        objItemRomaneiGradeAux.dQuantFaturada = objItemRomaneiGrade.dQuantFaturada
                        objItemRomaneiGradeAux.dQuantCancelada = objItemRomaneiGrade.dQuantCancelada
                        objItemRomaneiGradeAux.dQuantAFaturar = objItemRomaneiGrade.dQuantAFaturar
                        objItemRomaneiGradeAux.sDescricao = objItemRomaneiGrade.sDescricao
                        objItemRomaneiGradeAux.sProduto = objItemRomaneiGrade.sProduto
                        objItemRomaneiGradeAux.sUMEstoque = objItemRomaneiGrade.sUMEstoque
                        
                        objItemRomaneiGradeAux.iAlmoxarifado = objItemRomaneiGrade.iAlmoxarifado
                        objItemRomaneiGradeAux.sVersao = objItemRomaneiGrade.sVersao
                        
                        objItemOP.colItensRomaneioGrade.Add objItemRomaneiGradeAux
                    End If
                Next
            
            End If

        End If

    Next

    'Define a ordem de prioridade para distribuição da quanbtidade a ser produzida
    'e da producao entrada
    Set colCampos = New Collection
    colCampos.Add "iPrioridade"
    colCampos.Add "dtDataEntrega"
    colCampos.Add "dQuantidade"
    colCampos.Add "dtDataEmissao"

    For Each objItemOP In objOP.colItens
        Set colSaida = New Collection
        Call Ordena_Colecao(objItemOP.colItemPV, colSaida, colCampos)
        For iLinha = objItemOP.colItemPV.Count To 1 Step -1
            objItemOP.colItemPV.Remove iLinha
        Next
        For Each objItemOPItemPV In colSaida
            objItemOP.colItemPV.Add objItemOPItemPV
        Next
    Next

    'Percorre todas as OP da Coleção
    'For Each objGeracaoOPItem In colItens
    iLinha = 0
    For Each objItemOP In objOP.colItens

'        If True Or objItemOP.colItensRomaneioGrade.Count = 0 Then

            iLinha = iLinha + 1

            'Passa para a tela os dados da OP em questão
            sProdutoMascarado = String(STRING_PRODUTO, 0)

            'Mascara produto
            lErro = Mascara_MascararProduto(objItemOP.sProduto, sProdutoMascarado)
            If lErro <> SUCESSO Then Error 41542

            GridOP.TextMatrix(iLinha, GRIDOP_PRODUTO_COL) = sProdutoMascarado
            GridOP.TextMatrix(iLinha, GRIDOP_DESCRICAOITEMOP_COL) = objItemOP.sDescricao
            GridOP.TextMatrix(iLinha, GRIDOP_UM_COL) = objItemOP.sSiglaUMEstoque
            GridOP.TextMatrix(iLinha, GRIDOP_QUANT_COL) = Formata_Estoque(objItemOP.dQuantidade)
            GridOP.TextMatrix(iLinha, GRIDOP_QUANTFALTA_COL) = Formata_Estoque(objItemOP.dQuantidade)
             
            If objItemOP.colItensRomaneioGrade.Count = 0 Then
                        
                'Preneche o objProduto Filial para poder ler o Almoxarifado Padrao
                objProdutoFilial.sProduto = objItemOP.sProduto
                objProdutoFilial.iFilialEmpresa = giFilialEmpresa
                
                'Le o AlmoxarifadoPadrao
                lErro = CF("ProdutoFilial_Le", objProdutoFilial)
                If lErro <> SUCESSO And lErro <> 28261 Then Error 52181
                
                If lErro = 28261 Then Error 52182
                
                'Prenche para que seja Lido o Nome Reduzido do Almoxarifado Padrão
                objAlmoxarifado.iCodigo = objProdutoFilial.iAlmoxarifado
                
                lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25056 Then Error 52183
                            
                If lErro = 25056 Then Error 52184
                
                'Prenche o Almoxarifado com o Almoxarifado Padrão
                GridOP.TextMatrix(iLinha, GRIDOP_ALMOXARIFADO_COL) = objAlmoxarifado.sNomeReduzido
            
            Else
            
                GridOP.TextMatrix(iLinha, GRIDOP_ALMOXARIFADO_COL) = ""
                
            End If
            
            GridOP.TextMatrix(iLinha, GRIDOP_CCL_COL) = ""
            GridOP.TextMatrix(iLinha, GRIDOP_PRIORIDADE_COL) = CStr(objItemOP.iPrioridade)
            GridOP.TextMatrix(iLinha, GRIDOP_DATAPREVISAOINICIO_COL) = Format(gdtDataAtual, "dd/mm/yyyy")
            GridOP.TextMatrix(iLinha, GRIDOP_DATAPREVISAOFIM_COL) = Format(gdtDataAtual, "dd/mm/yyyy")
                                
'            Else
'
'            For Each objItemRomaneiGrade In objItemOP.colItensRomaneioGrade
'
'                iLinha = iLinha + 1
'
'                'Passa para a tela os dados da OP em questão
'                sProdutoMascarado = String(STRING_PRODUTO, 0)
'
'                'Mascara produto
'                lErro = Mascara_MascararProduto(objItemRomaneiGrade.sProduto, sProdutoMascarado)
'                If lErro <> SUCESSO Then Error 41542
'
'                GridOP.TextMatrix(iLinha, GRIDOP_PRODUTO_COL) = sProdutoMascarado
'                GridOP.TextMatrix(iLinha, GRIDOP_DESCRICAOITEMOP_COL) = objItemRomaneiGrade.sDescricao
'                GridOP.TextMatrix(iLinha, GRIDOP_UM_COL) = objItemRomaneiGrade.sUMEstoque
'
'                objItemRomaneiGrade.dQuantAFaturar = objItemRomaneiGrade.dQuantidade - objItemRomaneiGrade.dQuantFaturada - objItemRomaneiGrade.dQuantCancelada - objItemRomaneiGrade.dQuantReservada - objItemRomaneiGrade.dQuantOP
'
'                GridOP.TextMatrix(iLinha, GRIDOP_QUANT_COL) = Formata_Estoque(objItemRomaneiGrade.dQuantAFaturar)
'                GridOP.TextMatrix(iLinha, GRIDOP_QUANTFALTA_COL) = Formata_Estoque(objItemRomaneiGrade.dQuantAFaturar)
'
'                'Preneche o objProduto Filial para poder ler o Almoxarifado Padrao
'                objProdutoFilial.sProduto = objItemRomaneiGrade.sProduto
'                objProdutoFilial.iFilialEmpresa = giFilialEmpresa
'
'                'Le o AlmoxarifadoPadrao
'                lErro = CF("ProdutoFilial_Le", objProdutoFilial)
'                If lErro <> SUCESSO And lErro <> 28261 Then Error 52181
'
'                If lErro = 28261 Then Error 52182
'
'                'Prenche para que seja Lido o Nome Reduzido do Almoxarifado Padrão
'                objAlmoxarifado.iCodigo = objProdutoFilial.iAlmoxarifado
'
'                lErro = CF("Almoxarifado_Le", objAlmoxarifado)
'                If lErro <> SUCESSO And lErro <> 25056 Then Error 52183
'
'                If lErro = 25056 Then Error 52184
'
'                'Prenche o Almoxarifado com o Almoxarifado Padrão
'                GridOP.TextMatrix(iLinha, GRIDOP_ALMOXARIFADO_COL) = objAlmoxarifado.sNomeReduzido
'                GridOP.TextMatrix(iLinha, GRIDOP_CCL_COL) = ""
'                GridOP.TextMatrix(iLinha, GRIDOP_PRIORIDADE_COL) = CStr(objItemOP.iPrioridade)
'                GridOP.TextMatrix(iLinha, GRIDOP_DATAPREVISAOINICIO_COL) = Format(gdtDataAtual, "dd/mm/yyyy")
'                GridOP.TextMatrix(iLinha, GRIDOP_DATAPREVISAOFIM_COL) = Format(gdtDataAtual, "dd/mm/yyyy")
'
'            Next
'
'        End If

    Next
    
    Set gobjOP = objOP

    Call Grid_Refresh_Checkbox(objGridOP)

    'Passa para o Obj o número de OP's passados pela Coleção
    objGridOP.iLinhasExistentes = iLinha

    gbTrazendoDados = False

    Call Trata_ItemPV(1)

    Exit Sub

Erro_GridOP_Preenche:

    Select Case Err

        Case 41542, 49995, 90768
        
        Case 49996
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err, objFilialEmpresa.iCodFilial)
        
        Case 52181, 52183
        
        Case 52182
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTOFILIAL_INEXISTENTE", Err, objProdutoFilial.sProduto, objProdutoFilial.iFilialEmpresa)
        
        Case 52184
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", Err, objProdutoFilial.iAlmoxarifado)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160939)

    End Select

    Exit Sub

End Sub

Private Function Trata_TabOP() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_TabOP

    If iSelecionadosPV > 0 Then

        BotaoGerarOP.Enabled = True
        
        Call GridOP_Preenche(gobjGeracaoOP.colItens)

    End If

    Exit Function

Erro_Trata_TabOP:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160940)

    End Select

    Exit Function

End Function

Private Function Trata_TabPedidos() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_TabPedidos

    If Ordenados.ListIndex = -1 Then

        Ordenados.ListIndex = 0

    Else

        Set gobjGeracaoOP = New ClassGeracaoOP

        lErro = Move_TabSelecao_Memoria()
        If lErro <> SUCESSO Then Error 55876

        'Preenche a Coleção de Pedidos
        lErro = CF("GeracaoOP_SelecionarItens", gobjGeracaoOP)
        If lErro <> SUCESSO Then Error 41590

        'Preenche o GridPedido
        Call GridPV_Preenche(gobjGeracaoOP.colItens)

    End If

    Exit Function

Erro_Trata_TabPedidos:

    Trata_TabPedidos = Err

    Select Case Err

        Case 41590, 55876

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160941)

    End Select

    Exit Function

End Function

Private Function Critica_Campos() As Long

Dim lErro As Long
Dim sProdutoDe As String, sProdutoAte As String
Dim iPreenchidoDe As Integer, iPreenchidoAte As Integer

On Error GoTo Erro_Critica_Campos

    'Se PedidoInicial e PedidoFinal estão preenchidos
    If Len(Trim(PedidoInicial.ClipText)) > 0 And Len(Trim(PedidoFinal.ClipText)) > 0 Then

        'Verifica se ClienteAte é maior ou igual que ClienteDe
        If CLng(Trim(PedidoFinal.Text)) < CLng(Trim(PedidoInicial.Text)) Then Error 41442

    End If
    
       'Se ClienteAte e ClienteDe estão preenchidos
    If Len(Trim(ClienteDe.Text)) > 0 And Len(Trim(ClienteAte.Text)) > 0 Then

        'Verifica se ClienteAte é maior ou igual que ClienteDe
        If CLng(Trim(ClienteAte.Text)) < CLng(Trim(ClienteDe.Text)) Then Error 41437

    End If

    If Len(Trim(DataEntregaDe.ClipText)) > 0 And Len(Trim(DataEntregaAte.ClipText)) > 0 Then

        If CDate(DataEntregaDe.Text) > CDate(DataEntregaAte.Text) Then Error 41583
    
    End If

    lErro = CF("Produto_Formata", ProdutoDe.Text, sProdutoDe, iPreenchidoDe)
    If lErro <> SUCESSO Then Error 41446
    
    lErro = CF("Produto_Formata", ProdutoAte.Text, sProdutoAte, iPreenchidoAte)
    If lErro <> SUCESSO Then Error 41641
    
    If (iPreenchidoDe = PRODUTO_PREENCHIDO) And (iPreenchidoAte = PRODUTO_PREENCHIDO) Then
        If sProdutoDe > sProdutoAte Then Error 41448
    End If
    
    Critica_Campos = SUCESSO
    
    Exit Function
    
Erro_Critica_Campos:
    
    Critica_Campos = Err
    
    Select Case Err
        
        Case 41446, 41641
        
        Case 41448
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTODE_MAIOR_PRODUTOATE", Err)
            ProdutoAte.SetFocus
        
        Case 41442
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOINICIAL_MAIOR_PEDIDOFINAL", Err)
            PedidoFinal.SetFocus
        
        Case 41437
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTEDE_MAIOR_CLIENTEATE", Err)
            ClienteAte.SetFocus
        
        Case 41583
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)
            DataEntregaAte.SetFocus

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160942)
    
    End Select
    
    Exit Function

End Function

Private Function Move_TabSelecao_Memoria() As Long

Dim lErro As Long
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Move_TabSelecao_Memoria

    lErro = Critica_Campos()
    If lErro <> SUCESSO Then Error 41642
    
    gobjGeracaoOP.iTodos = ExibeTodos.Value

    'Se PedidoFinal e PedidoInicial estão preenchidos
    If Len(Trim(PedidoInicial.ClipText)) > 0 Then
        gobjGeracaoOP.lPVDe = CLng(PedidoInicial.Text)
    Else
        gobjGeracaoOP.lPVDe = 0
    End If

    If Len(Trim(PedidoFinal.ClipText)) > 0 Then
        gobjGeracaoOP.lPVAte = CLng(PedidoFinal.Text)
    Else
        gobjGeracaoOP.lPVAte = 0
    End If

    'Se ClienteAté e ClienteDe estão preenchidos
    If Len(Trim(ClienteDe.ClipText)) > 0 Then
        gobjGeracaoOP.lClientesDe = CLng(ClienteDe.Text)
    Else
        gobjGeracaoOP.lClientesDe = 0
    End If

    If Len(Trim(ClienteAte.ClipText)) > 0 Then
        gobjGeracaoOP.lClientesAte = CLng(ClienteAte.Text)
    Else
        gobjGeracaoOP.lClientesAte = 0
    End If

    'Se a DataEntregaDe está preenchida
    If Len(DataEntregaDe.ClipText) > 0 Then
        gobjGeracaoOP.dtEntregaDe = CDate(DataEntregaDe.Text)
    'Se a DataEntregaDe não está preenchida
    Else
        gobjGeracaoOP.dtEntregaDe = DATA_NULA
    End If

    'Se a DataEntregaAté está preenchida
    If Len(DataEntregaAte.ClipText) > 0 Then
        gobjGeracaoOP.dtEntregaAte = CDate(DataEntregaAte.Text)
    'Se a DataEntregaAté não está preenchida
    Else
        gobjGeracaoOP.dtEntregaAte = DATA_NULA
    End If

    sProduto = ProdutoDe.Text
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then Error 41591

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        gobjGeracaoOP.sProdDe = sProdutoFormatado
    Else
        gobjGeracaoOP.sProdDe = ""
    End If

    sProduto = ProdutoAte.Text
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then Error 41592

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        gobjGeracaoOP.sProdAte = sProdutoFormatado
    Else
        gobjGeracaoOP.sProdAte = ""
    End If

    If CheckExibePedidoOP.Value = MARCADO Then
        gobjGeracaoOP.iIncluiPVcomOP = MARCADO
    Else
        gobjGeracaoOP.iIncluiPVcomOP = DESMARCADO
    End If

    gobjGeracaoOP.iFilialEmpresa = giFilialEmpresa

    gobjGeracaoOP.sOrdenacao = asOrdenacao(Ordenados.ListIndex)

    Move_TabSelecao_Memoria = SUCESSO

    Exit Function

Erro_Move_TabSelecao_Memoria:

    Move_TabSelecao_Memoria = Err
    
    Select Case Err

        Case 41591, 41592, 41642

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160943)

    End Select

    Exit Function

End Function

Private Sub ExibeTodos_Click()

    iTabPrincipalAlterado = REGISTRO_ALTERADO

    'Limpa os campos da tela
    Call Inicializa_Campos

    'Se marcar ExibeTodos, exibe todos os pedidos
    If ExibeTodos.Value = MARCADO Then
        Call Troca_Enabled(False)
        CheckExibePedidoOP.Value = MARCADO
    Else
        Call Troca_Enabled(True)
    End If

End Sub

Private Function Inicializa_Mascaras() As Long

Dim lErro As Long
Dim sMascaraCclPadrao As String

On Error GoTo Erro_Inicializa_Mascaras

    'Inicializa máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoDe)
    If lErro <> SUCESSO Then Error 41593

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoAte)
    If lErro <> SUCESSO Then Error 41594

    'Inicializa Máscara de Ccl
    sMascaraCclPadrao = String(STRING_CCL, 0)

    lErro = MascaraCcl(sMascaraCclPadrao)
    If lErro <> SUCESSO Then Error 41595

    Ccl.Mask = sMascaraCclPadrao

    Quantidade.Format = FORMATO_ESTOQUE
    QuantOP.Format = FORMATO_ESTOQUE
    QuantFalta.Format = FORMATO_ESTOQUE
    
    Inicializa_Mascaras = SUCESSO

    Exit Function

Erro_Inicializa_Mascaras:

    Inicializa_Mascaras = Err

    Select Case Err

        Case 41593 To 41595

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160944)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    Set objGridPV = New AdmGrid
    Set objGridOP = New AdmGrid
    Set objGridItemPV = New AdmGrid
    
    Set gobjOP = New ClassOrdemDeProducao
    
    gbTrazendoDados = True

    Set objEventoPedidoDe = New AdmEvento
    Set objEventoPedidoAte = New AdmEvento
    Set objEventoClienteDe = New AdmEvento
    Set objEventoClienteAte = New AdmEvento
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    Set objEventoCcl = New AdmEvento
    Set objEventoCodigo = New AdmEvento
    Set objEventoEstoque = New AdmEvento

    asOrdenacao(0) = "ItensPedidoDeVenda.Produto, ItensPedidoDeVenda.DataEntrega"
    asOrdenacao(1) = "ItensPedidoDeVenda.DataEntrega, ItensPedidoDeVenda.FilialEmpresa, ItensPedidoDeVenda.CodPedido"
    asOrdenacao(2) = "ItensPedidoDeVenda.FilialEmpresa, ItensPedidoDeVenda.Produto, ItensPedidoDeVenda.DataEntrega"
    asOrdenacao(3) = "ItensPedidoDeVenda.DataEntrega, ItensPedidoDeVenda.Produto"

    asOrdenacaoString(0) = "Produto + Entrega"
    asOrdenacaoString(1) = "Entrega + Filial do Pedido + Pedido"
    asOrdenacaoString(2) = "Filial do Pedido + Produto + Entrega"
    asOrdenacaoString(3) = "Entrega + Produto"

    For iIndice = 0 To 3
        Ordenados.AddItem asOrdenacaoString(iIndice)
    Next

    'no frame 1:
        'selecionar o "Nao" (o "todos" comeca desmarcado e os outros campos "limpos")
    ExibeTodos.Value = DESMARCADO
    CheckExibePedidoOP.Value = DESMARCADO
    Call Inicializa_Campos

    'no frame 2:
        'selecionar a ordenacao "produto+entrega"
    Ordenados.ListIndex = 0
    iSelecionadosPV = 0

    'no frame 3:

    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True

    'Inicilaiza mascaras
    lErro = Inicializa_Mascaras()
    If lErro <> SUCESSO Then Error 41596

    lErro = Inicializa_GridOP(objGridOP)
    If lErro <> SUCESSO Then Error 41597

    lErro = Inicializa_GridPV(objGridPV)
    If lErro <> SUCESSO Then Error 41598

    lErro = Inicializa_GridItemPV(objGridItemPV)
    If lErro <> SUCESSO Then Error 41598

    If gobjMAT.iOPDetalhamentoRCs = MARCADO Then
        BotaoOPRC.Visible = True
    Else
        BotaoOPRC.Visible = False
    End If
    
    iFrameAtual = 1

    iAlterado = 0
    iTabPrincipalAlterado = 0
    iTabPedidoAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 41596 To 41598

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160945)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Limpa_Tela_GeracaoOP() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_GeracaoOP

    Call Limpa_Tela(Me)

    Call Grid_Limpa(objGridPV)
    Call Grid_Limpa(objGridOP)

    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True

    ExibeTodos.Value = DESMARCADO
    CheckExibePedidoOP = DESMARCADO

    'Torna Frame atual invisível
    Frame1(TabStrip1.SelectedItem.Index).Visible = False
    iFrameAtual = 1
    'Torna Frame atual visível
    Frame1(iFrameAtual).Visible = True
    TabStrip1.Tabs.Item(iFrameAtual).Selected = True
    
    iSelecionadosPV = 0
    iTabPrincipalAlterado = 0
    iTabPedidoAlterado = 0
    iAlterado = 0
    
    BotaoGerarOP.Enabled = False

    GeraOpSub.Value = vbUnchecked
    GeraReqCompra.Value = vbUnchecked
    GeraOPs.Value = vbUnchecked 'Inserido por Wagner

    Limpa_Tela_GeracaoOP = SUCESSO

    Exit Function

Erro_Limpa_Tela_GeracaoOP:

    Limpa_Tela_GeracaoOP = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160946)

    End Select

    Exit Function

End Function

Private Sub BotaoGerarOP_Click()
'codigo da OP tem que estar preenchido, bem como a data e gobjGeracaoOP.colItens tem que ter itens selecionados

Dim lErro As Long

On Error GoTo Erro_BotaoGerarOP_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 36610

    lErro = Limpa_Tela_GeracaoOP
    If lErro <> SUCESSO Then Error 36611

    Exit Sub

Erro_BotaoGerarOP_Click:

    Select Case Err

        Case 36610, 36611

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160947)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim objOrdemDeProducao As ClassOrdemDeProducao
Dim lErro As Long

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objOrdemDeProducao = obj1

    Codigo.Text = objOrdemDeProducao.sCodigo

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160948)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim objOrdemDeProducao As New ClassOrdemDeProducao
Dim colSelecao As Collection

    If Len(Trim(Codigo.Text)) > 0 Then objOrdemDeProducao.sCodigo = Codigo.Text

    'chama a tela ordemDeProducao
    Call Chama_Tela("OrdemProducaoLista", colSelecao, objOrdemDeProducao, objEventoCodigo)

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objOrdemDeProducao As New ClassOrdemDeProducao

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'verifica se o codigo foi preenchido
    If Len(Codigo.Text) = 0 Then gError 36612

    'Verifica se a Data foi preenchida
    If Len(Trim(Data.ClipText)) = 0 Then gError 36613

    'verifica se tem pelo menos 1 item na ordem de produção a ser gerada
    If objGridOP.iLinhasExistentes = 0 Then gError 36614

    'Loop de Validação dos dados do Grid
    For iIndice = 1 To objGridOP.iLinhasExistentes

        'Verifica se a quantidade foi digitada
        If Len(Trim(GridOP.TextMatrix(iIndice, GRIDOP_QUANT_COL))) = 0 Then gError 36615

        'Verifica se o almoxarifado foi informado
        'If Len(Trim(GridOP.TextMatrix(iIndice, GRIDOP_ALMOXARIFADO_COL))) = 0 Then gError 36616

        If Len(Trim(GridOP.TextMatrix(iIndice, GRIDOP_DATAPREVISAOINICIO_COL))) = 0 Then gError 41601

        If Len(Trim(GridOP.TextMatrix(iIndice, GRIDOP_DATAPREVISAOFIM_COL))) = 0 Then gError 41602

    Next

    If GeraOpSub.Value = vbChecked Then
        If Len(Trim(CodOPGerada.Text)) = 0 Then gError 126478
        
        
        If UCase(Trim(Codigo.Text)) = UCase(Trim(CodOPGerada.Text)) Then gError 180396
    End If

    lErro = Move_Tela_Memoria(objOrdemDeProducao)
    If lErro <> SUCESSO Then gError 36617

    lErro = CF("GeracaoOP_Grava", objOrdemDeProducao)
    If lErro <> SUCESSO Then gError 36618

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 36612
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGOOP_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus
 
        Case 36613
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)

        Case 36614
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NENHUM_ITEMOP_INFORMADO", gErr)

        Case 36615
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA", gErr, iIndice)

        Case 36616
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_PREENCHIDO", gErr, iIndice)

        Case 36617, 36618
        
        Case 41601
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAPREVISAOINICIO_NAO_INFORMADO", gErr, iIndice)

        Case 41602
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAPREVISAOFIM_NAO_INFORMADO", gErr, iIndice)

        Case 126478
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_OP_SUBPRODUTO_NAO_INF", gErr)
            
        Case 180396
            Call Rotina_Erro(vbOKOnly, "ERRO_CODOP_IGUAL_CODSUBOP", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160949)

    End Select

    Exit Function

End Function

Function Move_Tela_Memoria(objOrdemDeProducao As ClassOrdemDeProducao) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objOrdemDeProducao.sCodigo = Codigo.Text
    objOrdemDeProducao.dtDataEmissao = CDate(Data.Text)
    objOrdemDeProducao.iFilialEmpresa = giFilialEmpresa
    objOrdemDeProducao.iGeraOP = GeraOpSub.Value
    objOrdemDeProducao.iGeraOPsArvore = GeraOPs.Value 'Inserido por Wagner
    objOrdemDeProducao.iGeraReqCompra = GeraReqCompra.Value
    
    Set objOrdemDeProducao.objTela = Me
    
    If objOrdemDeProducao.iGeraOP = vbChecked Then
        objOrdemDeProducao.sCodOPGerar = CodOPGerada
    End If

    lErro = Move_Grid_Memoria(objOrdemDeProducao)
    If lErro <> SUCESSO Then Error 36619

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case 36619

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160950)

    End Select

    Exit Function

End Function

Private Sub GeraOPSub_Click()

    iAlterado = REGISTRO_ALTERADO
    
    If GeraOpSub.Value = vbChecked Then
        
        LabelCodOPGerada.Enabled = True
        CodOPGerada.Enabled = True
        '##############
        'INSERIDO POR WAGNER
        GeraOPs.Enabled = True
        '##############
        botaoProxNum2.Enabled = True 'Inserido por Wagner
    Else
        LabelCodOPGerada.Enabled = False
        CodOPGerada.Enabled = False
        CodOPGerada.Text = ""
        '##############
        'INSERIDO POR WAGNER
        GeraOPs.Enabled = False
        GeraOPs.Value = vbUnchecked
        '##############
        botaoProxNum2.Enabled = False 'Inserido por Wagner
    End If
    

End Sub

Private Sub GeraReqCompra_Click()

    iAlterado = REGISTRO_ALTERADO

    If gobjMAT.iOPDetalhamentoRCs = MARCADO Then
        If GeraReqCompra.Value = MARCADO Then
            BotaoOPRC.Enabled = True
        Else
            BotaoOPRC.Enabled = False
        End If
    End If
    
End Sub

Function Move_Grid_Memoria(objOrdemDeProducao As ClassOrdemDeProducao) As Long
'move os dados do grid de itens a serem produzidos para a memoria

Dim lErro As Long
Dim iIndice As Integer, iCount As Integer
Dim iProdutoPreenchido As Integer
Dim sProduto As String, sCcl As String, sCclFormatada As String, iCclPreenchida As Integer
Dim sProdutoFormatado As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objItemOP As ClassItemOP, objItemOPAux As ClassItemOP
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome

On Error GoTo Erro_Move_Grid_Memoria

    objOrdemDeProducao.iNumItens = 0

    'Para cada item do grid, guarda em um objeto os dados do grid
    For iIndice = 1 To objGridOP.iLinhasExistentes

        Set objItemOP = New ClassItemOP

        objItemOP.sCodigo = objOrdemDeProducao.sCodigo
        objItemOP.iFilialEmpresa = objOrdemDeProducao.iFilialEmpresa

        sProduto = GridOP.TextMatrix(iIndice, GRIDOP_PRODUTO_COL)

        'Critica o formato do Produto
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then Error 36620

        objItemOP.sProduto = sProdutoFormatado

        objItemOP.sSiglaUM = GridOP.TextMatrix(iIndice, GRIDOP_UM_COL)

        If Len(Trim(GridOP.TextMatrix(iIndice, GRIDOP_QUANT_COL))) > 0 Then
            objItemOP.dQuantidade = CDbl(GridOP.TextMatrix(iIndice, GRIDOP_QUANT_COL))
        Else
            objItemOP.dQuantidade = 0
        End If

        objAlmoxarifado.sNomeReduzido = GridOP.TextMatrix(iIndice, GRIDOP_ALMOXARIFADO_COL)

        If Len(Trim(objAlmoxarifado.sNomeReduzido)) <> 0 Then

            If colCodigoNome.Count > 0 Then
                
                For Each objCodigoNome In colCodigoNome
                    If objCodigoNome.sNome = objAlmoxarifado.sNomeReduzido Then
                        objItemOP.iAlmoxarifado = objCodigoNome.iCodigo
                        Exit For
                    End If
                Next
            
            End If
                    
            If objItemOP.iAlmoxarifado = 0 Then
    
                lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25056 Then Error 36621
    
                'trata se o almoxarifado não existir
                If lErro = 25056 Then Error 36622
    
                objItemOP.iAlmoxarifado = objAlmoxarifado.iCodigo
        
                colCodigoNome.Add objAlmoxarifado.iCodigo, objAlmoxarifado.sNomeReduzido
                
            End If
        
        End If
        
        If gobjOP.colItens.Count >= iIndice Then
            Set objItemOPAux = gobjOP.colItens.Item(iIndice)
            
            If objItemOPAux.colItensRomaneioGrade.Count <> 0 Then
            
                Set objItemOP.colItensRomaneioGrade = objItemOPAux.colItensRomaneioGrade
                objItemOP.iPossuiGrade = MARCADO
            
            End If
            
        End If
        
        sCcl = GridOP.TextMatrix(iIndice, GRIDOP_CCL_COL)

        If Len(Trim(sCcl)) <> 0 Then

            'Formata Ccl para BD
            lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
            If lErro <> SUCESSO Then Error 36623

        Else

            sCclFormatada = ""

        End If

        objItemOP.sCcl = sCclFormatada

        objItemOP.iItem = iIndice

        'continuação de Move_Grid_Memoria
        lErro = Move_Grid_Memoria1(objItemOP, iIndice)
        If lErro <> SUCESSO Then Error 36626

        Set objItemOP.colProdutoKitInfo = gobjOP.colItens.Item(iIndice).colProdutoKitInfo
        If gobjOP.colItens.Item(iIndice).colItemPV.Count = 1 Then
            objItemOP.lCodPedido = gobjOP.colItens.Item(iIndice).colItemPV.Item(1).lCodigoPV
            objItemOP.iFilialPedido = gobjOP.colItens.Item(iIndice).colItemPV.Item(1).iFilialEmpresa
        Else
            Set objItemOP.colItemPV = gobjOP.colItens.Item(iIndice).colItemPV
        End If

        objOrdemDeProducao.colItens.Add objItemOP

        objOrdemDeProducao.iNumItens = objOrdemDeProducao.iNumItens + 1

    Next

    Move_Grid_Memoria = SUCESSO

    Exit Function

Erro_Move_Grid_Memoria:

    Move_Grid_Memoria = Err

    Select Case Err

        Case 36620, 36621, 36623, 36626

        Case 36622
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", Err, objAlmoxarifado.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160951)

    End Select

    Exit Function

End Function

Function Move_Grid_Memoria1(objItemOP As ClassItemOP, ByVal iIndice As Integer) As Long
'continuação de Move_Grid_Memoria

Dim sItem As String
Dim objFiliais As AdmFiliais
Dim lErro As Long
Dim objKit As New ClassKit

On Error GoTo Erro_Move_Grid_Memoria1

    If Len(GridOP.TextMatrix(iIndice, GRIDOP_DATAPREVISAOINICIO_COL)) > 0 Then
        objItemOP.dtDataInicioProd = CDate(GridOP.TextMatrix(iIndice, GRIDOP_DATAPREVISAOINICIO_COL))
    Else
        objItemOP.dtDataInicioProd = DATA_NULA
    End If

    If Len(GridOP.TextMatrix(iIndice, GRIDOP_DATAPREVISAOFIM_COL)) > 0 Then
        objItemOP.dtDataFimProd = CDate(GridOP.TextMatrix(iIndice, GRIDOP_DATAPREVISAOFIM_COL))
    Else
        objItemOP.dtDataFimProd = DATA_NULA
    End If

    'verifica se a data fim é maior ou igual a data inicio
    If objItemOP.dtDataInicioProd <> DATA_NULA And objItemOP.dtDataFimProd <> DATA_NULA And objItemOP.dtDataFimProd < objItemOP.dtDataInicioProd Then Error 36624

    objItemOP.iSituacao = ITEMOP_SITUACAO_NORMAL
    objItemOP.iDestinacao = ITEMOP_DESTINACAO_PV

'    'descobre o identificador da filial do pedido
'    For Each objFiliais In gcolFiliais
'        If objFiliais.iCodFilial <> 0 Then
'            If objFiliais.iCodFilial = CInt(Codigo_Extrai(GridOP.TextMatrix(iIndice, GRIDOP_FILIALPV_COL))) Then
'                objItemOP.iFilialPedido = objFiliais.iCodFilial
'                Exit For
'            End If
'        End If
'    Next
'
'    objItemOP.lCodPedido = CLng(GridOP.TextMatrix(iIndice, GRIDOP_PEDIDODEVENDA_COL))

    If Len(Trim(GridOP.TextMatrix(iIndice, GRIDOP_PRIORIDADE_COL))) > 0 Then objItemOP.iPrioridade = CInt(GridOP.TextMatrix(iIndice, GRIDOP_PRIORIDADE_COL))

    'Armazena o Produto Raiz do kit
    objKit.sProdutoRaiz = objItemOP.sProduto
    
    'Le as Versoes Ativas e a Padrao
    lErro = CF("Kit_Le_Padrao", objKit)
    If lErro <> SUCESSO And lErro <> 106304 Then gError 126480

    objItemOP.sVersao = objKit.sVersao

    Move_Grid_Memoria1 = SUCESSO

    Exit Function

Erro_Move_Grid_Memoria1:

    Move_Grid_Memoria1 = Err

    Select Case Err

        Case 36624
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAFIMOP_ANTERIOR_DATAINICIOOP", Err, objItemOP.dtDataFimProd, objItemOP.dtDataInicioProd, iIndice)

        Case 126480

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160952)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_GERACAO_OP_SELECAO
    Set Form_Load_Ocx = Me
    Caption = "Geração Automática de Ordens de Produção"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "GeracaoOP"
    
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

Public Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO
    
End Function

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is PedidoInicial Then
            Call LabelPedidoDe_Click
        ElseIf Me.ActiveControl Is PedidoFinal Then
            Call LabelPedidoAte_Click
        ElseIf Me.ActiveControl Is ClienteDe Then
            Call LabelClienteDe_Click
        ElseIf Me.ActiveControl Is ClienteAte Then
            Call LabelClienteAte_Click
        ElseIf Me.ActiveControl Is ProdutoDe Then
            Call LabelProdutoDe_Click
        ElseIf Me.ActiveControl Is ProdutoAte Then
            Call LabelProdutoAte_Click
        ElseIf Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        ElseIf Me.ActiveControl Is Almoxarifado Then
            Call BotaoEstoque_Click
        ElseIf Me.ActiveControl Is Ccl Then
            Call BotaoCcls_Click
        End If
    End If

End Sub


Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub LabelClienteDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteDe, Source, X, Y)
End Sub

Private Sub LabelClienteDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteDe, Button, Shift, X, Y)
End Sub

Private Sub LabelClienteAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteAte, Source, X, Y)
End Sub

Private Sub LabelClienteAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteAte, Button, Shift, X, Y)
End Sub

Private Sub LabelPedidoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelPedidoAte, Source, X, Y)
End Sub

Private Sub LabelPedidoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelPedidoAte, Button, Shift, X, Y)
End Sub

Private Sub LabelPedidoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelPedidoDe, Source, X, Y)
End Sub

Private Sub LabelPedidoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelPedidoDe, Button, Shift, X, Y)
End Sub

Private Sub LabelProdutoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProdutoAte, Source, X, Y)
End Sub

Private Sub LabelProdutoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProdutoAte, Button, Shift, X, Y)
End Sub

Private Sub LabelProdutoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProdutoDe, Source, X, Y)
End Sub

Private Sub LabelProdutoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProdutoDe, Button, Shift, X, Y)
End Sub

'Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label3, Source, X, Y)
'End Sub
'
'Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
'End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigo, Source, X, Y)
End Sub

Private Sub LabelCodigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigo, Button, Shift, X, Y)
End Sub

Private Sub ItensSelecionados_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ItensSelecionados, Source, X, Y)
End Sub

Private Sub ItensSelecionados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ItensSelecionados, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub LabelItensSelecionados_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelItensSelecionados, Source, X, Y)
End Sub

Private Sub LabelItensSelecionados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelItensSelecionados, Button, Shift, X, Y)
End Sub

'###################################################
'Inserido por Wagner
Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim sCodigoOP As String

On Error GoTo Erro_BotaoProxNum_Click

    'Mostra número do proximo lote disponível
    lErro = CF("OrdemProducao_Automatico", sCodigoOP, giFilialEmpresa)
    If lErro <> SUCESSO Then gError 131860

    Codigo.Text = sCodigoOP

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 131860
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160953)
    
    End Select

    Exit Sub

End Sub

Private Sub botaoProxNum2_Click()

Dim lErro As Long
Dim sCodigoOP As String

On Error GoTo Erro_BotaoProxNum2_Click

    'Mostra número do proximo lote disponível
    lErro = CF("OrdemProducao_Automatico", sCodigoOP, giFilialEmpresa)
    If lErro <> SUCESSO Then gError 131875

    CodOPGerada.Text = sCodigoOP

    Exit Sub

Erro_BotaoProxNum2_Click:

    Select Case gErr

        Case 131875
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160954)
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoOPRC_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objItemOP As ClassItemOP

On Error GoTo Erro_BotaoOPRC_Click:

    gobjOP.sCodigo = Codigo.Text
    
'    For iIndice = 1 To objGridOP.iLinhasExistentes
'
'        Set objItemOP = gobjOP.colItens.Item(iIndice)
'
'        objItemOP.sVersao = GridOP.TextMatrix(iIndice, iGrid_Versao_Col)
'
'    Next

    'Se o item tiver como originador o Orcamento de Venda, busca a versao
    lErro = CF("OrdemProducao_Le_Versao_OV", gobjOP)
    If lErro <> SUCESSO Then gError 181895
                
    Call Chama_Tela_Modal("OrdemProducaoRC", gobjOP)

    Exit Sub

Erro_BotaoOPRC_Click:

    Select Case gErr
    
        Case 181895

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174307)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoEntrega_Click()

Dim lErro As Long
Dim objPV As New ClassPedidoDeVenda
Dim objItemPV As ClassItemPedido
Dim sProdutoTela As String
Dim sProdutoBD As String
Dim iStatus As Integer
Dim bAchou As Boolean
Dim iProdutoPreenchido As Integer
Dim dQuantidade As Double
Dim objDataEntrega As ClassDataEntrega
Dim iIndice As Integer

On Error GoTo Erro_BotaoEntrega_Click

    If GridPV.Row = 0 Then gError 182841

    dQuantidade = StrParaDbl(GridPV.TextMatrix(GridPV.Row, GRIDPV_QUANTIDADE_COL))
    sProdutoTela = GridPV.TextMatrix(GridPV.Row, GRIDPV_PRODUTO_COL)
    sProdutoBD = String(STRING_PRODUTO, 0)
   
    'Critica o formato do Produto
    lErro = CF("Produto_Formata", sProdutoTela, sProdutoBD, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 182842
    
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 182843
        
    objPV.iFilialEmpresa = giFilialEmpresa
    objPV.lCodigo = StrParaLong(GridPV.TextMatrix(GridPV.Row, GRIDPV_PEDIDO_COL))
    
    'Lê os dados do Pedido de Venda
    lErro = CF("PedidoDeVenda_Le_Todos_Completo", objPV, iStatus)
    If lErro <> SUCESSO Then gError 182844
    
    bAchou = False
    For Each objItemPV In objPV.colItensPedido
        If objItemPV.sProduto = sProdutoBD Then
            bAchou = True
            Exit For
        End If
    Next
    
    If Not bAchou Then gError 182845

    lErro = CF("NFiscal_Le_DatasEntrega", objPV)
    If lErro <> SUCESSO Then gError 199583

    For iIndice = objItemPV.colItemPVEntrega.Count To 1 Step -1
        
        Set objDataEntrega = objItemPV.colItemPVEntrega.Item(iIndice)
            
        If objDataEntrega.dQuantidadeEntregue >= objDataEntrega.dQuantidade Then
            objItemPV.colItemPVEntrega.Remove (iIndice)
        Else
            objDataEntrega.dQuantidade = objDataEntrega.dQuantidade - objDataEntrega.dQuantidadeEntregue
        End If
    
    Next
        
    Call Chama_Tela_Modal("DataEntrega", sProdutoTela, dQuantidade, objItemPV.colItemPVEntrega, DESMARCADO)

    Exit Sub

Erro_BotaoEntrega_Click:

    Select Case gErr

        Case 182841
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 182842, 182844, 182845, 199583

        Case 182843
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182846)

    End Select

    Exit Sub
    
End Sub

Private Function Inicializa_GridItemPV(objGridInt As AdmGrid) As Long
'Inicializa o GridItemPV
Dim lErro  As Long

On Error GoTo Erro_Inicializa_GridItemPV

    'Form do GridItemPV
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("PV")
    objGridInt.colColuna.Add ("Filial PV")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Data Emissão")
    objGridInt.colColuna.Add ("Data Entrega")

    'Controles que participam do GridItemPV
    objGridInt.colCampo.Add (ItemPVPedido.Name)
    objGridInt.colCampo.Add (ItemPVFilialPV.Name)
    objGridInt.colCampo.Add (ItemPVQuantidade.Name)
    objGridInt.colCampo.Add (ItemPVUM.Name)
    objGridInt.colCampo.Add (ItemPVCliente.Name)
    objGridInt.colCampo.Add (ItemPVFilial.Name)
    objGridInt.colCampo.Add (ItemPVDataEmissao.Name)
    objGridInt.colCampo.Add (ItemPVDataEntrega.Name)

    'Grid do GridInterno
    objGridInt.objGrid = GridItemPV

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 3

    'Largura da primeira coluna
    GridItemPV.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridItemPV = SUCESSO

    Exit Function

Erro_Inicializa_GridItemPV:

    Inicializa_GridItemPV = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160911)

    End Select

    Exit Function

End Function

Private Function Trata_ItemPV(Optional iItem As Integer = 0) As Long

Dim lErro  As Long
Dim objItemOP As ClassItemOP
Dim objItemOPItemPV As ClassItemOPItemPV
Dim iLinha As Integer
Dim sMsgAux As String

On Error GoTo Erro_Trata_ItemPV

    If iItem = 0 Then iItem = GridOP.Row
       
    Call Grid_Limpa(objGridItemPV)
    
    If iItem <> 0 Then
    
        Set objItemOP = gobjOP.colItens.Item(iItem)

        iLinha = 0
        For Each objItemOPItemPV In objItemOP.colItemPV
            iLinha = iLinha + 1
            
            GridItemPV.TextMatrix(iLinha, GRIDITEMPV_PEDIDO_COL) = CStr(objItemOPItemPV.lCodigoPV)
            GridItemPV.TextMatrix(iLinha, GRIDITEMPV_FILIALPV_COL) = CStr(objItemOPItemPV.iFilialEmpresa) & SEPARADOR & objItemOPItemPV.sNomeFilialEmpresa
            GridItemPV.TextMatrix(iLinha, GRIDITEMPV_QUANTIDADE_COL) = Formata_Estoque(objItemOPItemPV.dQuantidade)
            GridItemPV.TextMatrix(iLinha, GRIDITEMPV_UM_COL) = objItemOPItemPV.sUM
            GridItemPV.TextMatrix(iLinha, GRIDITEMPV_CLIENTE_COL) = CStr(objItemOPItemPV.lCliente) & SEPARADOR & objItemOPItemPV.sNomeCliente
            GridItemPV.TextMatrix(iLinha, GRIDITEMPV_FILIAL_COL) = CStr(objItemOPItemPV.iFilial) & SEPARADOR & objItemOPItemPV.sNomeFilialCliente
            If objItemOPItemPV.dtDataEmissao <> DATA_NULA Then GridItemPV.TextMatrix(iLinha, GRIDITEMPV_DATAEMISSAO_COL) = Format(objItemOPItemPV.dtDataEmissao, "dd/mm/yyyy")
            If objItemOPItemPV.dtDataEntrega <> DATA_NULA Then GridItemPV.TextMatrix(iLinha, GRIDITEMPV_DATAENTREGA_COL) = Format(objItemOPItemPV.dtDataEntrega, "dd/mm/yyyy")
        
        Next
        objGridItemPV.iLinhasExistentes = objItemOP.colItemPV.Count
        
        sMsgAux = CStr(iItem) & SEPARADOR & GridOP.TextMatrix(iItem, GRIDOP_PRODUTO_COL) & SEPARADOR & GridOP.TextMatrix(iItem, GRIDOP_DESCRICAOITEMOP_COL)
        
    End If
    
    FrameItemPV.Caption = "Pedidos de Venda associados ao item (" & sMsgAux & ")"

    Trata_ItemPV = SUCESSO

    Exit Function

Erro_Trata_ItemPV:

    Trata_ItemPV = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160911)

    End Select

    Exit Function

End Function

Private Sub GridItemPV_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItemPV, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItemPV, iAlterado)
    End If

End Sub

Private Sub GridItemPV_EnterCell()
    Call Grid_Entrada_Celula(objGridItemPV, iAlterado)
End Sub

Private Sub GridItemPV_GotFocus()
    Call Grid_Recebe_Foco(objGridItemPV)
End Sub

Private Sub GridItemPV_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridItemPV)
End Sub

Private Sub GridItemPV_KeyPress(KeyAscii As Integer)
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItemPV, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItemPV, iAlterado)
    End If

End Sub

Private Sub GridItemPV_LeaveCell()
    Call Saida_Celula(objGridItemPV)
End Sub

Private Sub GridItemPV_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridItemPV)
End Sub

Private Sub GridItemPV_RowColChange()
    Call Grid_RowColChange(objGridItemPV)
End Sub

Private Sub GridItemPV_Scroll()
    Call Grid_Scroll(objGridItemPV)
End Sub
