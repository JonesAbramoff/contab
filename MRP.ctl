VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl MRP 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5235
      Index           =   4
      Left            =   150
      TabIndex        =   28
      Top             =   690
      Visible         =   0   'False
      Width           =   9180
      Begin VB.CommandButton BotaoTrazer 
         Caption         =   "Simular MRP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7395
         TabIndex        =   15
         ToolTipText     =   "Traz os Dados do Plano Mestre de Produção"
         Top             =   165
         Width           =   1665
      End
      Begin VB.ComboBox VersaoMRP 
         Height          =   315
         ItemData        =   "MRP.ctx":0000
         Left            =   885
         List            =   "MRP.ctx":0002
         TabIndex        =   14
         Top             =   165
         Width           =   2325
      End
      Begin VB.Frame Frame2 
         Caption         =   "Plano Mestre de Produção"
         Height          =   4530
         Left            =   15
         TabIndex        =   29
         Top             =   660
         Width           =   9150
         Begin MSMask.MaskEdBox CodPV 
            Height          =   315
            Left            =   4170
            TabIndex        =   100
            Top             =   750
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   6
            PromptChar      =   " "
         End
         Begin VB.CommandButton BotaoOP 
            Caption         =   "Ordem de Produção"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   1065
            TabIndex        =   19
            ToolTipText     =   "Abre a tela de Ordem de Produção"
            Top             =   3930
            Width           =   1035
         End
         Begin VB.CommandButton BotaoGrafico 
            Caption         =   "Cronograma Gráfico"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   2115
            TabIndex        =   20
            ToolTipText     =   "Abra a tela de Cronograma"
            Top             =   3930
            Width           =   1260
         End
         Begin VB.CommandButton BotaoProjecao 
            Caption         =   "Projeção de Estoque"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   3390
            TabIndex        =   21
            ToolTipText     =   "Abre o relatório de projeção de estoque"
            Top             =   3930
            Width           =   1260
         End
         Begin VB.TextBox StatusDetalhado 
            BackColor       =   &H8000000F&
            Height          =   690
            Left            =   1650
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Top             =   3120
            Width           =   7410
         End
         Begin MSMask.MaskEdBox DataProducao 
            Height          =   315
            Left            =   5115
            TabIndex        =   59
            Top             =   915
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.CommandButton BotaoErros 
            Caption         =   "Relatórios de Críticas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   4665
            TabIndex        =   22
            ToolTipText     =   "Abre o relatório de críticas"
            Top             =   3930
            Width           =   1290
         End
         Begin VB.CommandButton BotaoPreviaCarga 
            Caption         =   "Prévia de Carga nos CTs"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   5970
            TabIndex        =   23
            ToolTipText     =   "Abre o relatório de prévia de carga nos Cts"
            Top             =   3930
            Width           =   1560
         End
         Begin VB.CommandButton BotaoPreviaReqCompra 
            Caption         =   "Prévia de Req. de Compra"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   7545
            TabIndex        =   24
            ToolTipText     =   "Abre o relatório de prévia de requisições de compra"
            Top             =   3930
            Width           =   1560
         End
         Begin MSMask.MaskEdBox CodOP 
            Height          =   315
            Left            =   3150
            TabIndex        =   56
            Top             =   720
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   6
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Prioridade 
            Height          =   315
            Left            =   1440
            TabIndex        =   57
            Top             =   1260
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Status 
            Height          =   315
            Left            =   4275
            TabIndex        =   38
            Top             =   1485
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataNecess 
            Height          =   315
            Left            =   7080
            TabIndex        =   37
            Top             =   2205
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox UM 
            Height          =   315
            Left            =   4980
            TabIndex        =   36
            Top             =   2220
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Versao 
            Height          =   315
            Left            =   2250
            TabIndex        =   35
            Top             =   1635
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin VB.TextBox Descricao 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Left            =   1950
            MaxLength       =   250
            TabIndex        =   34
            Top             =   2175
            Width           =   2010
         End
         Begin MSMask.MaskEdBox Produto 
            Height          =   315
            Left            =   2385
            TabIndex        =   33
            Top             =   1125
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   315
            Left            =   5685
            TabIndex        =   31
            Top             =   2235
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Cliente 
            Height          =   315
            Left            =   1650
            TabIndex        =   32
            Top             =   675
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.CommandButton BotaoVerEtapas 
            Caption         =   "Visualizar Etapas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   15
            TabIndex        =   18
            ToolTipText     =   "Abre a tela de Etapas da Produção"
            Top             =   3930
            Width           =   1035
         End
         Begin MSFlexGridLib.MSFlexGrid GridItens 
            Height          =   2415
            Left            =   120
            TabIndex        =   16
            Top             =   315
            Width           =   8955
            _ExtentX        =   15796
            _ExtentY        =   4260
            _Version        =   393216
            Rows            =   7
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin VB.Label LabelStatusMRP 
            Caption         =   "Status detalhado:"
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
            Left            =   45
            TabIndex        =   58
            Top             =   3135
            Width           =   1560
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Versão:"
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
         Height          =   270
         Left            =   195
         TabIndex        =   30
         Top             =   210
         Width           =   690
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5040
      Index           =   3
      Left            =   120
      TabIndex        =   61
      Top             =   810
      Visible         =   0   'False
      Width           =   9225
      Begin MSMask.MaskEdBox PrioridadePV 
         Height          =   255
         Left            =   405
         TabIndex        =   99
         Top             =   825
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
         Left            =   7320
         TabIndex        =   65
         Top             =   4260
         Width           =   1800
      End
      Begin VB.TextBox DescricaoItem 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   3060
         MaxLength       =   50
         TabIndex        =   74
         Top             =   1680
         Width           =   2610
      End
      Begin VB.TextBox UMPV 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   2550
         TabIndex        =   73
         Text            =   "UM"
         Top             =   2610
         Width           =   510
      End
      Begin VB.TextBox FilialPV 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   2925
         TabIndex        =   72
         Text            =   "Filial do PV"
         Top             =   3384
         Width           =   1170
      End
      Begin VB.TextBox DataEmissao 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   4512
         TabIndex        =   71
         Text            =   "Emissão"
         Top             =   2904
         Width           =   1200
      End
      Begin VB.TextBox Filial 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   3450
         TabIndex        =   70
         Text            =   "Filial"
         Top             =   2940
         Width           =   870
      End
      Begin VB.TextBox ClientePV 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   1992
         TabIndex        =   69
         Text            =   "Cliente"
         Top             =   2904
         Width           =   1260
      End
      Begin VB.TextBox Pedido 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   4710
         TabIndex        =   68
         Text            =   "Pedido"
         Top             =   3375
         Width           =   1365
      End
      Begin VB.CheckBox GeraOP 
         DragMode        =   1  'Automatic
         Height          =   270
         Left            =   1545
         TabIndex        =   67
         Top             =   1830
         Width           =   735
      End
      Begin VB.TextBox DataEntrega 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   4500
         TabIndex        =   66
         Text            =   "Entrega"
         Top             =   2595
         Width           =   1200
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
         Left            =   5325
         Picture         =   "MRP.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   4260
         Width           =   1800
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
         Left            =   3345
         Picture         =   "MRP.ctx":11E6
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   4260
         Width           =   1800
      End
      Begin VB.ComboBox Ordenados 
         Height          =   315
         ItemData        =   "MRP.ctx":2200
         Left            =   2550
         List            =   "MRP.ctx":2202
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   210
         Width           =   3480
      End
      Begin MSMask.MaskEdBox ProdutoPV 
         Height          =   225
         Left            =   1245
         TabIndex        =   75
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
      Begin MSMask.MaskEdBox QuantidadePV 
         Height          =   225
         Left            =   3210
         TabIndex        =   76
         Top             =   2610
         Width           =   1260
         _ExtentX        =   2223
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
         Left            =   150
         TabIndex        =   77
         Top             =   690
         Width           =   9000
         _ExtentX        =   15875
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
         TabIndex        =   60
         Top             =   4395
         Width           =   1800
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
         TabIndex        =   79
         Top             =   255
         Width           =   1410
      End
      Begin VB.Label ItensSelecionados 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2100
         TabIndex        =   78
         Top             =   4365
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5040
      Index           =   2
      Left            =   120
      TabIndex        =   81
      Top             =   870
      Visible         =   0   'False
      Width           =   9225
      Begin VB.CheckBox OPMRP 
         DragMode        =   1  'Automatic
         Height          =   270
         Left            =   1980
         TabIndex        =   98
         Top             =   1725
         Width           =   735
      End
      Begin VB.TextBox OPCodigo 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   255
         Left            =   3225
         TabIndex        =   97
         Top             =   390
         Width           =   552
      End
      Begin VB.TextBox DescricaoItemOP 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   4050
         MaxLength       =   50
         TabIndex        =   89
         Top             =   0
         Width           =   2460
      End
      Begin VB.TextBox UMOP 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   255
         Left            =   1530
         TabIndex        =   88
         Text            =   "UM"
         Top             =   360
         Width           =   552
      End
      Begin VB.TextBox FilialPVOP 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   255
         Left            =   2985
         TabIndex        =   87
         Top             =   795
         Width           =   1170
      End
      Begin MSMask.MaskEdBox PrioridadeOP 
         Height          =   255
         Left            =   5550
         TabIndex        =   90
         Top             =   810
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
      Begin MSMask.MaskEdBox PedidoDeVendaOP 
         Height          =   255
         Left            =   4140
         TabIndex        =   91
         Top             =   795
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
      Begin MSMask.MaskEdBox QuantOP 
         Height          =   255
         Left            =   2145
         TabIndex        =   92
         Top             =   360
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
      Begin MSMask.MaskEdBox ProdutoOP 
         Height          =   255
         Left            =   15
         TabIndex        =   93
         Top             =   510
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
         Left            =   4260
         TabIndex        =   94
         Top             =   420
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataPrevisaoFim 
         Height          =   255
         Left            =   1920
         TabIndex        =   95
         Top             =   795
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
         Left            =   705
         TabIndex        =   96
         Top             =   825
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
      Begin VB.CommandButton BotaoMarcarOP 
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
         Left            =   3492
         Picture         =   "MRP.ctx":2204
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   4245
         Width           =   1800
      End
      Begin VB.CommandButton BotaoDesmarcarOP 
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
         Left            =   5475
         Picture         =   "MRP.ctx":321E
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   4245
         Width           =   1800
      End
      Begin MSFlexGridLib.MSFlexGrid GridOP 
         Height          =   3570
         Left            =   150
         TabIndex        =   84
         Top             =   60
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   6297
         _Version        =   393216
         Rows            =   15
         Cols            =   7
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
         AllowUserResizing=   1
      End
      Begin VB.Label ItensSelecionadosOP 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2115
         TabIndex        =   86
         Top             =   4350
         Width           =   795
      End
      Begin VB.Label Label2 
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
         TabIndex        =   85
         Top             =   4395
         Width           =   1800
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5130
      Index           =   1
      Left            =   165
      TabIndex        =   39
      Top             =   750
      Width           =   9165
      Begin VB.CheckBox TrazerPV 
         Caption         =   "Trazer Pedidos de Venda"
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
         TabIndex        =   80
         Top             =   4740
         Value           =   1  'Checked
         Width           =   3360
      End
      Begin VB.Frame FrameDataOP 
         Caption         =   "Data da O.P."
         Height          =   810
         Left            =   315
         TabIndex        =   53
         Top             =   3825
         Width           =   7950
         Begin MSMask.MaskEdBox DataOPInicial 
            Height          =   300
            Left            =   705
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
         Begin MSComCtl2.UpDown UpDownDataOPInicial 
            Height          =   300
            Left            =   1860
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   330
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataOPFinal 
            Height          =   300
            Left            =   4575
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
         Begin MSComCtl2.UpDown UpDownDataOPFinal 
            Height          =   300
            Left            =   5745
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   330
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label LabelDataAte 
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
            Index           =   1
            Left            =   4140
            TabIndex        =   55
            Top             =   345
            Width           =   360
         End
         Begin VB.Label LabelDataDe 
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
            Index           =   0
            Left            =   360
            TabIndex        =   54
            Top             =   345
            Width           =   315
         End
      End
      Begin VB.Frame FrameOP 
         Caption         =   "Ordem de Produção"
         Height          =   810
         Left            =   300
         TabIndex        =   50
         Top             =   1500
         Width           =   7950
         Begin VB.TextBox OpFinal 
            Height          =   300
            Left            =   4545
            MaxLength       =   6
            TabIndex        =   3
            Top             =   330
            Width           =   1680
         End
         Begin VB.TextBox OpInicial 
            Height          =   300
            Left            =   690
            MaxLength       =   6
            TabIndex        =   2
            Top             =   315
            Width           =   1680
         End
         Begin VB.Label LabelOpFinal 
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
            Left            =   4110
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   52
            Top             =   375
            Width           =   360
         End
         Begin VB.Label LabelOpInicial 
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
            Left            =   285
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   51
            Top             =   360
            Width           =   315
         End
      End
      Begin VB.Frame FrameCliente 
         Caption         =   "Cliente"
         Height          =   1395
         Left            =   315
         TabIndex        =   45
         Top             =   30
         Width           =   7935
         Begin MSMask.MaskEdBox ClienteInicial 
            Height          =   315
            Left            =   705
            TabIndex        =   0
            Top             =   360
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ClienteFinal 
            Height          =   315
            Left            =   705
            TabIndex        =   1
            Top             =   840
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label LabelClienteAte 
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
            TabIndex        =   49
            Top             =   885
            Width           =   435
         End
         Begin VB.Label LabelClienteDe 
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
            TabIndex        =   48
            Top             =   390
            Width           =   360
         End
         Begin VB.Label DescClienteInicial 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2295
            TabIndex        =   47
            Top             =   360
            Width           =   5535
         End
         Begin VB.Label DescClienteFinal 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2280
            TabIndex        =   46
            Top             =   840
            Width           =   5535
         End
      End
      Begin VB.Frame FrameProdutos 
         Caption         =   "Produtos"
         Height          =   1395
         Index           =   0
         Left            =   315
         TabIndex        =   40
         Top             =   2370
         Width           =   7935
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
            TabIndex        =   44
            Top             =   840
            Width           =   5535
         End
         Begin VB.Label DescProdInic 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2280
            TabIndex        =   43
            Top             =   360
            Width           =   5535
         End
         Begin VB.Label LabelProdutoDe 
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
            TabIndex        =   42
            Top             =   390
            Width           =   360
         End
         Begin VB.Label LabelProdutoAte 
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
            TabIndex        =   41
            Top             =   885
            Width           =   435
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6810
      ScaleHeight     =   495
      ScaleWidth      =   2535
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   60
      Width           =   2595
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   585
         Picture         =   "MRP.ctx":4400
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   1080
         Picture         =   "MRP.ctx":455A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1575
         Picture         =   "MRP.ctx":46E4
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGerar 
         Height          =   360
         Left            =   90
         Picture         =   "MRP.ctx":4C16
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Gerar o MRP"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   2070
         Picture         =   "MRP.ctx":5058
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5595
      Left            =   60
      TabIndex        =   8
      Top             =   360
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   9869
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ordens de Produção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pedidos de Venda"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "MRP"
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
Attribute VB_Name = "MRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjPMP As ClassPMP
Dim gobjPMPOP As ClassPMP
Dim gobjGeracaoOP As New ClassGeracaoOP

Dim sOPs As String

Dim iAlterado As Integer
Dim iAlteradoTab As Integer
Dim iAlteradoFiltro As Integer
Dim iAlteradoVersao As Integer
Dim iAlteradoOPPV As Integer

Dim iSelecionadosPV As Integer
Dim iSelecionadosOP As Integer

Dim iFrameAtual As Integer

Dim giOp_Inicial As Integer

Dim asOrdenacao(3) As String
Dim asOrdenacaoString(3) As String

'Grid de Itens
Dim objGridItens As AdmGrid
Dim iGrid_Cliente_Col As Integer
Dim iGrid_CodOP_Col As Integer
Dim iGrid_CodPV_Col As Integer
Dim iGrid_Prioridade_Col As Integer
Dim iGrid_Produto_Col As Integer
Dim iGrid_Versao_Col As Integer
Dim iGrid_Descricao_Col As Integer
Dim iGrid_UM_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_DataNecess_Col As Integer
Dim iGrid_DataProducao_Col As Integer
Dim iGrid_Status_Col As Integer

Private WithEvents objEventoClienteInic As AdmEvento
Attribute objEventoClienteInic.VB_VarHelpID = -1
Private WithEvents objEventoClienteFim As AdmEvento
Attribute objEventoClienteFim.VB_VarHelpID = -1
Private WithEvents objEventoOp As AdmEvento
Attribute objEventoOp.VB_VarHelpID = -1
Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1

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
Const GRIDPV_PRIORIDADE_COL = 12

Dim objGridOP As AdmGrid
Const GRIDOP_OPMRP_COL = 1
Const GRIDOP_CODOP_COL = 2
Const GRIDOP_PRODUTO_COL = 3
Const GRIDOP_DESCRICAOITEMOP_COL = 4
Const GRIDOP_UM_COL = 5
Const GRIDOP_QUANT_COL = 6
Const GRIDOP_ALMOXARIFADO_COL = 7
Const GRIDOP_DATAPREVISAOINICIO_COL = 8
Const GRIDOP_DATAPREVISAOFIM_COL = 9
Const GRIDOP_PEDIDODEVENDA_COL = 10
Const GRIDOP_FILIALPV_COL = 11
Const GRIDOP_PRIORIDADE_COL = 12

Const TAB_Selecao = 1
Const TAB_OP = 2
Const TAB_PV = 3
Const TAB_MRP = 4

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Grava a a simulação
    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 136460

    'Limpa Tela
    Call Limpa_Tela_MRP
    
    'Le as versões cadastradas
    lErro = CF("PMPCombo_Le_Versoes", VersaoMRP)
    If lErro <> SUCESSO Then gError 136468
    
    'fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoGravar_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 136460, 136468

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163097)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objMaquinas As New ClassMaquinas
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(VersaoMRP.Text)) = 0 Then gError 136458

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_SIMULACAO_PMP", gobjPMP.sVersao)

    If vbMsgRes = vbYes Then

        'Exclui a a simulação
        lErro = CF("PMP_Exclui_Simulacao", gobjPMP)
        If lErro <> SUCESSO Then gError 136459
    
        'Limpa Tela
        Call Limpa_Tela_MRP
    
        'Le as versões cadastradas
        lErro = CF("PMPCombo_Le_Versoes", VersaoMRP)
        If lErro <> SUCESSO Then gError 136469
    
    End If
    
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 136458
            Call Rotina_Erro(vbOKOnly, "ERRO_VERSAO_NAO_PREENCHIDO", gErr)
        
        Case 136459, 136469

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163098)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163099)

    End Select

    Exit Sub

End Sub

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "MRP"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "MRP"

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

Private Sub BotaoProjecao_Click()

Dim lErro As Long
Dim lNumIntRel As Long
Dim objRelatorio As New AdmRelatorio
    
On Error GoTo Erro_BotaoProjecao_Click
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = CF("RelEstoqueSimulacaoPMP_Prepara", gobjPMP, giFilialEmpresa, lNumIntRel)
    If lErro <> SUCESSO Then gError 138073
    
    GL_objMDIForm.MousePointer = vbDefault
    
    lErro = objRelatorio.ExecutarDireto("Projeção de Estoque", "", 0, "", "NNUMINTREL", CStr(lNumIntRel))
    If lErro <> SUCESSO Then gError 138075
    
    Exit Sub

Erro_BotaoProjecao_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 138073, 138075

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163100)

    End Select
    
    Exit Sub
    
End Sub

Private Sub ClienteInicial_Change()
    
    iAlteradoTab = REGISTRO_ALTERADO
    iAlteradoFiltro = REGISTRO_ALTERADO
    
    'por Jorge Specian
    Call Cliente_Preenche(ClienteInicial)

End Sub

Private Sub ClienteInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ClienteInicial, iAlterado)
    
End Sub

Private Sub ClienteInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteInicial_Validate

    DescClienteInicial.Caption = ""

    If Len(Trim(ClienteInicial.Text)) > 0 Then
   
        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteInicial, objCliente, 0)
        If lErro <> SUCESSO Then gError 137599
        
        ClienteInicial.Text = objCliente.sNomeReduzido
        DescClienteInicial.Caption = objCliente.sRazaoSocial

    End If
        
    Exit Sub

Erro_ClienteInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 137599
            'erro tratado na rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163101)

    End Select

End Sub

Private Sub ClienteFinal_Change()
    
    iAlteradoTab = REGISTRO_ALTERADO
    iAlteradoFiltro = REGISTRO_ALTERADO
    
    Call Cliente_Preenche(ClienteFinal)

End Sub

Private Sub ClienteFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ClienteFinal, iAlterado)
    
End Sub

Private Sub ClienteFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteFinal_Validate

    DescClienteFinal.Caption = ""
    
    If Len(Trim(ClienteFinal.Text)) > 0 Then
   
        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteFinal, objCliente, 0)
        If lErro <> SUCESSO Then gError 137600
        
        ClienteFinal.Text = objCliente.sNomeReduzido
        DescClienteFinal.Caption = objCliente.sRazaoSocial

    End If
        
    Exit Sub

Erro_ClienteFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 137600
            'erro tratado na rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163102)

    End Select

End Sub

Private Sub LabelClienteAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCliente As New ClassCliente

On Error GoTo Erro_LabelClienteAte_Click
    
    'Verifica se Cliente Final foi preenchido
    If Len(Trim(ClienteFinal.Text)) > 0 Then
    
        If IsNumeric(ClienteFinal.Text) Then

            objCliente.lCodigo = StrParaLong(ClienteFinal.Text)
            
        Else
        
            objCliente.sNomeReduzido = ClienteFinal.Text
        
        End If

    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteFim)

   Exit Sub

Erro_LabelClienteAte_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163103)

    End Select

    Exit Sub


End Sub

Private Sub LabelClienteDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCliente As New ClassCliente

On Error GoTo Erro_LabelClienteDe_Click
        
    'Verifica se Cliente Inicial foi preenchido
    If Len(Trim(ClienteInicial.Text)) > 0 Then

        If IsNumeric(ClienteInicial.Text) Then

            objCliente.lCodigo = StrParaLong(ClienteInicial.Text)

        Else

            objCliente.sNomeReduzido = ClienteInicial.Text

        End If

    End If
        
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteInic)

   Exit Sub

Erro_LabelClienteDe_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163104)

    End Select

    Exit Sub
    
End Sub

Private Sub LabelOpFinal_Click()

Dim lErro As Long
Dim colSelecao As Collection
Dim objOP As New ClassOrdemDeProducao
Dim sSelecao As String

On Error GoTo Erro_LabelOpFinal_Click

    giOp_Inicial = 0
    
    If Len(Trim(OpFinal.Text)) <> 0 Then

        objOP.sCodigo = OpFinal.Text

    End If
    sSelecao = "Tipo = 0"
    
    'lista as OP's
    Call Chama_Tela("OrdemProducaoLista", colSelecao, objOP, objEventoOp, sSelecao)
    
    Exit Sub

Erro_LabelOpFinal_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163105)

    End Select

    Exit Sub

End Sub

Private Sub LabelOpInicial_Click()

Dim lErro As Long
Dim colSelecao As Collection
Dim objOP As New ClassOrdemDeProducao
Dim sSelecao As String

On Error GoTo Erro_LabelOpInicial_Click

    giOp_Inicial = 1

    If Len(Trim(OpInicial.Text)) <> 0 Then
    
        objOP.sCodigo = OpInicial.Text

    End If

    sSelecao = "Tipo = 0"
    
    'lista as OP's
    Call Chama_Tela("OrdemProducaoLista", colSelecao, objOP, objEventoOp, sSelecao)
    
    Exit Sub

Erro_LabelOpInicial_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163106)

    End Select

    Exit Sub

End Sub

Private Sub LabelProdutoAte_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProdutoAte_Click

    'Verifica se o produto foi preenchido
    If Len(ProdutoFinal.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", ProdutoFinal.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 137601

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoProduzivelLista", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 137601
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163107)

    End Select

    Exit Sub

End Sub

Private Sub LabelProdutoDe_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProdutoDe_Click

    'Verifica se o produto foi preenchido
    If Len(ProdutoInicial.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", ProdutoInicial.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 137602

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoProduzivelLista", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 137602
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163108)

    End Select

    Exit Sub

End Sub

Private Sub objEventoClienteFim_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    ClienteFinal.Text = CStr(objCliente.lCodigo)
    Call ClienteFinal_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

End Sub

Private Sub objEventoClienteInic_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    ClienteInicial.Text = CStr(objCliente.lCodigo)
    Call ClienteInicial_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoOp_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOP As New ClassOrdemDeProducao

On Error GoTo Erro_objEventoOp_evSelecao

    Set objOP = obj1

    If giOp_Inicial = 1 Then

        OpInicial.Text = objOP.sCodigo
        
    Else

        OpFinal.Text = objOP.sCodigo

    End If

    Me.Show
    
    Exit Sub

Erro_objEventoOp_evSelecao:

    Select Case Err

       Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 163109)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 137604

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 137605

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 137606

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 137604, 137606
            'erros tratados nas rotinas chamadas

        Case 137605
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163110)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoDe_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 137607

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 137608

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 137609

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 137607, 137609
            'erro tratado na rotina chamada

        Case 137608
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163111)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataOPFinal_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataOPFinal_DownClick

    DataOPFinal.SetFocus

    If Len(DataOPFinal.ClipText) > 0 Then

        sData = DataOPFinal.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 137610

        DataOPFinal.Text = sData

    End If

    Exit Sub

Erro_UpDownDataOPFinal_DownClick:

    Select Case gErr

        Case 137610

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163112)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataOPFinal_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataOPFinal_UpClick

    DataOPFinal.SetFocus

    If Len(Trim(DataOPFinal.ClipText)) > 0 Then

        sData = DataOPFinal.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 137611

        DataOPFinal.Text = sData

    End If

    Exit Sub

Erro_UpDownDataOPFinal_UpClick:

    Select Case gErr

        Case 137611

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163113)

    End Select

    Exit Sub

End Sub

Private Sub DataOPInicial_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataOPInicial, iAlterado)
End Sub

Private Sub DataOPInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataOPInicial_Validate

    If Len(Trim(DataOPInicial.ClipText)) <> 0 Then

        lErro = Data_Critica(DataOPInicial.Text)
        If lErro <> SUCESSO Then gError 137612

    End If

    Exit Sub

Erro_DataOPInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 137612

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163114)

    End Select

    Exit Sub

End Sub

Private Sub DataOPInicial_Change()
    iAlteradoTab = REGISTRO_ALTERADO
End Sub

Private Sub UpDownDataOPInicial_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataOPInicial_DownClick

    DataOPInicial.SetFocus

    If Len(DataOPInicial.ClipText) > 0 Then

        sData = DataOPInicial.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 137613

        DataOPInicial.Text = sData

    End If

    Exit Sub

Erro_UpDownDataOPInicial_DownClick:

    Select Case gErr

        Case 137613

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163115)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataOPInicial_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataOPInicial_UpClick

    DataOPInicial.SetFocus

    If Len(Trim(DataOPInicial.ClipText)) > 0 Then

        sData = DataOPInicial.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 137614

        DataOPInicial.Text = sData

    End If

    Exit Sub

Erro_UpDownDataOPInicial_UpClick:

    Select Case gErr

        Case 137614

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163116)

    End Select

    Exit Sub

End Sub

Private Sub DataOPFinal_Change()
    iAlteradoTab = REGISTRO_ALTERADO
End Sub

Private Sub DataOPFinal_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataOPFinal, iAlterado)
End Sub

Private Sub DataOPFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataOPFinal_Validate

    If Len(Trim(DataOPFinal.ClipText)) <> 0 Then

        lErro = Data_Critica(DataOPFinal.Text)
        If lErro <> SUCESSO Then gError 137615

    End If

    Exit Sub

Erro_DataOPFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 137615

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163117)

    End Select

    Exit Sub

End Sub

Private Sub OpFinal_Change()
    iAlteradoTab = REGISTRO_ALTERADO
End Sub

Private Sub OpFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_OpFinal_Validate

    giOp_Inicial = 0

    Exit Sub

Erro_OpFinal_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163118)

    End Select

    Exit Sub

End Sub

Private Sub OpInicial_Change()
    iAlteradoTab = REGISTRO_ALTERADO
End Sub

Private Sub OpInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_OpInicial_Validate

    giOp_Inicial = 1

    Exit Sub

Erro_OpInicial_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163119)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoFinal_Change()
    iAlteradoTab = REGISTRO_ALTERADO
    iAlteradoFiltro = REGISTRO_ALTERADO
End Sub

Private Sub ProdutoFinal_GotFocus()
    Call MaskEdBox_TrataGotFocus(ProdutoFinal, iAlterado)
End Sub

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate

    lErro = CF("Produto_Perde_Foco", ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 137616
    
    If lErro <> SUCESSO Then gError 137617
  
    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 137616
            'erro tratado na rotina chamada

        Case 137617
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163120)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Change()
    iAlteradoTab = REGISTRO_ALTERADO
    iAlteradoFiltro = REGISTRO_ALTERADO
End Sub

Private Sub ProdutoInicial_GotFocus()
    Call MaskEdBox_TrataGotFocus(ProdutoInicial, iAlterado)
End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    lErro = CF("Produto_Perde_Foco", ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 137618
    
    If lErro <> SUCESSO Then gError 137619

    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 137618
            'erro tratado na rotina chamada
            
        Case 137619
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163121)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)

Dim lErro As Long
Dim objPMP As New ClassPMP
Dim objMRPSeleciona As New ClassMRPSeleciona
Dim bFalse As Boolean
Dim objPMPItem As ClassPMPItens
Dim objItemOP As ClassItemOP
Dim iIndice As Integer

On Error GoTo Erro_TabStrip1_BeforeClick

    GL_objMDIForm.MousePointer = vbHourglass

    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
    
    sOPs = ""
        
    'Se estava no tab de seleção e está passando para outro tab
    If iFrameAtual = TAB_Selecao Then
    
        'Valida a seleção
        lErro = ValidaSelecao()
        If lErro <> SUCESSO Then gError 136384
        
        If iAlteradoTab = REGISTRO_ALTERADO Then
        
            Set gobjPMP = New ClassPMP
                           
            lErro = Move_TabSelecao_Memoria(objMRPSeleciona)
            If lErro <> SUCESSO Then gError 138428
                           
            'Le os itens que podem entrar para o plano mestre
            lErro = CF("PMP_Gera", objMRPSeleciona, objPMP)
            If lErro <> SUCESSO Then gError 136386
                           
            If lErro = SUCESSO Then
            
                For iIndice = objPMP.colItens.Count To 1 Step -1
                
                    Set objPMPItem = objPMP.colItens(iIndice)
                
                    If objPMPItem.objItemOP.colOrdemProducaoOperacoes.Count = 0 Then
                        If Len(Trim(sOPs)) = 0 Then
                            sOPs = sOPs & objPMPItem.sCodOPOrigem
                        Else
                            sOPs = sOPs & ", " & objPMPItem.sCodOPOrigem
                        End If
                        
                        objPMP.colItens.Remove iIndice
                    End If
                Next
                            
                Call GridOP_Preenche(objPMP)
                iAlteradoTab = 0
            End If
                           
            If TrazerPV.Value = vbChecked And iAlteradoFiltro = REGISTRO_ALTERADO Then
                                    
                Set gobjGeracaoOP.colItens = New Collection
                
                bFalse = False
                
                'Preenche a Coleção de Pedidos
                lErro = CF("GeracaoOP_SelecionarItens", gobjGeracaoOP, bFalse)
                If lErro <> SUCESSO And lErro <> 27437 Then gError 138429
                
                If lErro = SUCESSO Then
                    Call GridPV_Preenche(gobjGeracaoOP.colItens)
                    iAlteradoFiltro = 0
                End If
                
                
            End If
            
        End If
        
    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_TabStrip1_BeforeClick:

    Cancel = True

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 136384, 138428, 138429, 136386

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163122)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long
Dim objPMP As New ClassPMP

On Error GoTo Erro_TabStrip1_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub
        
        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index
        
        If iFrameAtual = TAB_MRP Then
        
            'Se o PMP não foi gerado ou foi incluido itens
            If gobjPMP.colItens.Count = 0 Or iAlteradoOPPV = REGISTRO_ALTERADO Then
            
                Call Grid_Limpa(objGridItens)
            
                'Inclui as OPs selecionadas
                lErro = PMP_Inclui_OP(objPMP)
                If lErro <> SUCESSO Then gError 138510
                
                'Gera e Inclui as OPs por PV
                lErro = PMP_Inclui_OP_Por_PV(objPMP)
                If lErro <> SUCESSO Then gError 138511
                
                Set gobjPMP = objPMP
                
            End If
        
        End If
                              
    End If

    GL_objMDIForm.MousePointer = vbDefault

'    If Len(Trim(sOPs)) <> 0 Then Call Rotina_Aviso(vbOKOnly, "AVISO_OPS_NAO_APTAS_MRP", sOPs)
'    sOPs = ""

    Exit Sub

Erro_TabStrip1_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 138510, 138511
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163123)

    End Select

    Exit Sub

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty(True, UserControl.Enabled, True)
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

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property
'**** fim do trecho a ser copiado *****

Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load
        
    iFrameAtual = 1
    
    Set gobjPMPOP = New ClassPMP
        
    Set objEventoClienteInic = New AdmEvento
    Set objEventoClienteFim = New AdmEvento
    
    Set objEventoOp = New AdmEvento
    
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    
    Set gobjPMP = New ClassPMP

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

    Ordenados.ListIndex = 0

    'inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoInicial)
    If lErro <> SUCESSO Then gError 137620

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoFinal)
    If lErro <> SUCESSO Then gError 137621
    
    'Le as versões cadastradas
    lErro = CF("PMPCombo_Le_Versoes", VersaoMRP)
    If lErro <> SUCESSO Then gError 136467
        
    'Grid Itens
    Set objGridItens = New AdmGrid
    Set objGridPV = New AdmGrid
    Set objGridOP = New AdmGrid

    'tela em questão
    Set objGridItens.objForm = Me
    Set objGridPV.objForm = Me
    Set objGridOP.objForm = Me
    
    lErro = Inicializa_GridItens(objGridItens)
    If lErro <> SUCESSO Then gError 137622
    
    lErro = Inicializa_GridPV(objGridPV)
    If lErro <> SUCESSO Then gError 137622
    
    lErro = Inicializa_GridOP(objGridOP)
    If lErro <> SUCESSO Then gError 137622
    
    iAlterado = 0
    iAlteradoTab = REGISTRO_ALTERADO
    iAlteradoVersao = 0
    iAlteradoFiltro = REGISTRO_ALTERADO
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 137620 To 137622, 136467
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163124)

    End Select

    Exit Sub

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set gobjPMPOP = Nothing

    Set objEventoClienteInic = Nothing
    Set objEventoClienteFim = Nothing

    Set objEventoOp = Nothing
    
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing

    Set gobjPMP = Nothing

    Set objGridItens = Nothing
    Set objGridPV = Nothing
    Set objGridOP = Nothing

    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163125)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163126)

    End Select
    
    Exit Function

End Function

Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

        Call Grid_Click(objGridItens, iExecutaEntradaCelula)

        If iExecutaEntradaCelula = 1 Then
            Call Grid_Entrada_Celula(objGridItens, iAlterado)
        End If

End Sub

Private Sub GridItens_GotFocus()
    Call Grid_Recebe_Foco(objGridItens)
End Sub

Private Sub GridItens_EnterCell()
    Call Grid_Entrada_Celula(objGridItens, iAlterado)
End Sub

Private Sub GridItens_LeaveCell()
    Call Saida_Celula(objGridItens)
End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer


    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGridItens)

    Call Mostra_StatusDetalhado

End Sub

Private Sub GridItens_Scroll()
    Call Grid_Scroll(objGridItens)
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
    
        'Verifica se é o GridItens
        If objGridInt.objGrid.Name = GridItens.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case Else
                    'Não há tratamento específico na saída de célula,
                    'uma vez que os campos não serão editados
        
            End Select
                        
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 137623

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 137623
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163127)

    End Select

    Exit Function

End Function

Private Function Inicializa_GridItens(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Status")
    objGrid.colColuna.Add ("Cliente")
    objGrid.colColuna.Add ("O.P.")
    objGrid.colColuna.Add ("P.V.")
    objGrid.colColuna.Add ("Prior.")
    objGrid.colColuna.Add ("Produto")
    objGrid.colColuna.Add ("Versão")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("UM")
    objGrid.colColuna.Add ("Quant")
    objGrid.colColuna.Add ("Dt Necess")
    objGrid.colColuna.Add ("Dt Fim Prod")

    'Controles que participam do Grid
    objGrid.colCampo.Add (Status.Name)
    objGrid.colCampo.Add (Cliente.Name)
    objGrid.colCampo.Add (CodOP.Name)
    objGrid.colCampo.Add (CodPV.Name)
    objGrid.colCampo.Add (Prioridade.Name)
    objGrid.colCampo.Add (Produto.Name)
    objGrid.colCampo.Add (Versao.Name)
    objGrid.colCampo.Add (Descricao.Name)
    objGrid.colCampo.Add (UM.Name)
    objGrid.colCampo.Add (Quantidade.Name)
    objGrid.colCampo.Add (DataNecess.Name)
    objGrid.colCampo.Add (DataProducao.Name)

    'Colunas do Grid
    iGrid_Status_Col = 1
    iGrid_Cliente_Col = 2
    iGrid_CodOP_Col = 3
    iGrid_CodPV_Col = 4
    iGrid_Prioridade_Col = 5
    iGrid_Produto_Col = 6
    iGrid_Versao_Col = 7
    iGrid_Descricao_Col = 8
    iGrid_UM_Col = 9
    iGrid_Quantidade_Col = 10
    iGrid_DataNecess_Col = 11
    iGrid_DataProducao_Col = 12

    objGrid.objGrid = GridItens

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 6

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 250
    
    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    Call Grid_Inicializa(objGrid)

    Inicializa_GridItens = SUCESSO

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable
                
    Select Case objControl.Name
    
        Case GeraOP.Name, OPMRP.Name
            objControl.Enabled = True
        
        Case Else
            objControl.Enabled = False
    
    End Select
                
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163128)

    End Select

    Exit Sub

End Sub

Private Sub Cliente_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Cliente_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Cliente
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Produto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Produto_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Versao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Versao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Versao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Versao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Versao
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Descricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Descricao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Descricao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Descricao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Descricao
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UM_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UM_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub UM_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub UM_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = UM
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Quantidade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Quantidade_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DataNecess_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataNecess_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub DataNecess_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub DataNecess_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DataNecess
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Status_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Status_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Status_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Status_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Status
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Function ValidaSelecao() As Long

Dim objClienteInicial As ClassCliente
Dim objClienteFinal As ClassCliente
Dim sProd_I As String
Dim sProd_F As String
Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim lErro As Long

On Error GoTo Erro_ValidaSelecao

    'Valida Clientes
    If Len(Trim(ClienteInicial.Text)) <> 0 And Len(Trim(ClienteFinal.Text)) <> 0 Then
    
        Set objClienteInicial = New ClassCliente
    
        objClienteInicial.sNomeReduzido = ClienteInicial.Text
        
        'Lê Cliente Inicial pelo NomeReduzido
        lErro = CF("Cliente_Le_NomeReduzido", objClienteInicial)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 137624
                
        Set objClienteFinal = New ClassCliente
        
        objClienteFinal.sNomeReduzido = ClienteFinal.Text
        
        'Lê Cliente Final pelo NomeReduzido
        lErro = CF("Cliente_Le_NomeReduzido", objClienteFinal)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 137625
                
        'codigo do cliente inicial não pode ser maior que o final
        If objClienteInicial.lCodigo > objClienteFinal.lCodigo Then gError 137626
        
    End If
    
    'Valida OPs
    'ordem de produção inicial não pode ser maior que a final
    If Len(Trim(OpInicial.Text)) <> 0 And Len(Trim(OpFinal.Text)) <> 0 Then

        If OpInicial.Text > OpFinal.Text Then gError 137627

    End If

    'Valida Produtos
    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 137628
    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 137629
    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambas os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then gError 137630
    
    End If
    
    'Valida Datas de OPs
    'data da OP inicial não pode ser maior que a final
    If Len(Trim(DataOPInicial.Text)) <> 0 And Len(Trim(DataOPFinal.Text)) <> 0 Then
        
        If StrParaDate(DataOPInicial.Text) > StrParaDate(DataOPFinal.Text) Then gError 137631
    
    End If
    
    ValidaSelecao = SUCESSO
    
    Exit Function
    
Erro_ValidaSelecao:

    ValidaSelecao = gErr

    Select Case gErr
    
        Case 137624, 137625
            'erros tratados nas rotinas chamadas
                
        Case 137626
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)
    
        Case 137627
            Call Rotina_Erro(vbOKOnly, "ERRO_OP_INICIAL_MAIOR", gErr)
            
        Case 137628
            ProdutoInicial.SetFocus

        Case 137629
            ProdutoFinal.SetFocus

        Case 137630
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
                        
        Case 137631
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAOP_INICIAL_MAIOR", gErr)
                                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163129)
    
    End Select

    Exit Function

End Function

Private Sub CodOP_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodOP_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub CodOP_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub CodOp_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = CodOP
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Prioridade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Prioridade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub Prioridade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub Prioridade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Prioridade
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

'por Jorge Specian - chamada por Cliente_Change para localizar pela parte digitada do Nome
'Reduzido do Cliente através da CF Cliente_Pesquisa_NomeReduzido em RotinasCRFAT.ClassCRFATSelect
Private Sub Cliente_Preenche(objControle As Object)

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
    
On Error GoTo Erro_Cliente_Preenche
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objControle, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 137632

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 137632

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163130)

    End Select
    
    Exit Sub

End Sub

Private Function Atualizar(Optional ByVal bForcaAtualizacao As Boolean = False) As Long
'Atualiza os dados da tela com base nas informações que o usuário alterou

Dim lErro As Long
Dim objPMP As New ClassPMP
Dim objPMPItem As ClassPMPItens
Dim colPMPItens As New Collection
Dim objPMPItemAux As ClassPMPItens
Dim objMRPSeleciona As New ClassMRPSeleciona
Dim iIndice As Integer
Dim objPO As ClassPlanoOperacional
Dim objPOAux As ClassPlanoOperacional
Dim bMudouPMPItem As Boolean
Dim bMudouPMP As Boolean
Dim bMudouPO As Boolean
Dim colSaida As New Collection
Dim colCampos As New Collection

On Error GoTo Erro_Atualizar
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    bMudouPMPItem = False
    bMudouPO = False
    bMudouPMP = False

    'Dá preferência por manter a ordem que já está na tela
    'E verifica se algum item foi alterado
    iIndice = 0
    For Each objPMPItem In gobjPMP.colItens
        iIndice = iIndice + 1
        objPMPItem.iOrdem = iIndice
        If objPMPItem.iAlterado = REGISTRO_ALTERADO Then
            bMudouPMPItem = True
        Else
            'Guarda o que não foi mudado
            colPMPItens.Add objPMPItem
        End If
        For Each objPO In objPMPItem.ColPO
            If objPO.iAlterado = REGISTRO_ALTERADO Then
                bMudouPO = True
            End If
        Next

    Next

    'Ordena pela prioridade
    colCampos.Add "iPrioridade"
    colCampos.Add "iOrdem"

    lErro = Ordena_Colecao(gobjPMP.colItens, colSaida, colCampos)
    If lErro <> SUCESSO Then gError 136494

    Set gobjPMP.colItens = colSaida

    StatusDetalhado.Text = ""

    'Se alterou algum item Re-Simula (Data/Prioridade)
    If bMudouPMPItem Or bForcaAtualizacao Then
        
        'Faz a simulação do plano mestre
        lErro = CF("MRP_Simula", gobjPMP)
        If lErro <> SUCESSO Then gError 136387
           
'        'Retira os itens que foram inseridos por pedido de venda
'        'mas que com o cenário atual não necessitam serem produzidos
'        lErro = PMP_Retira_Sem_Quantidade(gobjPMP)
'        If lErro <> SUCESSO Then gError 138445
'
        bMudouPO = True
   
        'Recupera os itens que não foram mudados (CT)
        iIndice = 0
        For Each objPMPItem In objPMP.colItens
            iIndice = iIndice + 1
            For Each objPMPItemAux In colPMPItens
                If objPMPItemAux.objItemOP.lNumIntDoc = objPMPItem.objItemOP.lNumIntDoc Then
                    Set objPMP.colItens.Item(iIndice).ColPO = objPMPItemAux.ColPO
                    objPMP.colItens.Item(iIndice).iPrioridade = objPMPItemAux.iPrioridade
                End If
            Next
        Next

    End If

    'Se alterou informações básicas tem que refazer a validação
    If bMudouPO Then

        lErro = CF("MRP_Valida", gobjPMP)
        If lErro <> SUCESSO Then gError 136388

    End If

    'Traz o que foi gerado ou validado para tela
    lErro = Traz_MRP_Tela(gobjPMP)
    If lErro <> SUCESSO Then gError 136389

    gobjPMP.iAlterado = 0
    
    For Each objPMPItem In gobjPMP.colItens
        objPMPItem.iAlterado = 0
        For Each objPO In objPMPItem.ColPO
            objPO.iAlterado = 0
        Next
    Next
    
    iAlteradoTab = 0
    iAlteradoFiltro = 0
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Atualizar = SUCESSO

    Exit Function

Erro_Atualizar:

    Atualizar = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 136385 To 136389, 138428, 138445

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163131)

    End Select

    Exit Function

End Function

Private Function Traz_MRP_Tela(ByVal objPMP As ClassPMP) As Long
'Traz para tela os dados gerados pelo algoritmo do MRP

Dim lErro As Long

On Error GoTo Erro_Traz_MRP_Tela
       
    objPMP.sVersao = VersaoMRP.Text
       
    Set gobjPMP = objPMP
    
    lErro = Preenche_Grid_Itens(objPMP)
    If lErro <> SUCESSO Then gError 136390
    
    Traz_MRP_Tela = SUCESSO

    Exit Function

Erro_Traz_MRP_Tela:

    Traz_MRP_Tela = gErr

    Select Case gErr
    
        Case 136390

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163132)

    End Select

    Exit Function

End Function

Private Function Preenche_Grid_Itens(objPMP As ClassPMP) As Long
'Preenche o grid com os caditados a itens do Plano Mestre de Produção

Dim lErro As Long
Dim objPMPItens As New ClassPMPItens
Dim iLinha As Integer
Dim objProduto As ClassProduto
Dim sProdutoMascarado  As String
Dim objCliente As ClassCliente
Dim sStatus As String

On Error GoTo Erro_Preenche_Grid_Itens

    Call Grid_Limpa(objGridItens)

    For Each objPMPItens In objPMP.colItens

        If objPMPItens.dQuantidade > 0 Then

            iLinha = iLinha + 1
                    
            Set objProduto = New ClassProduto
            Set objCliente = New ClassCliente
            
            objProduto.sCodigo = objPMPItens.sProduto
            
            'Lê o Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 136391
        
            'Se não achou o Produto --> erro
            If lErro = 28030 Then gError 136392
            
            sProdutoMascarado = String(STRING_PRODUTO, 0)
    
            'Coloca a máscara no produto
            lErro = Mascara_RetornaProdutoTela(objPMPItens.sProduto, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 136393
    
            GridItens.TextMatrix(iLinha, iGrid_Produto_Col) = sProdutoMascarado
            GridItens.TextMatrix(iLinha, iGrid_Descricao_Col) = objProduto.sDescricao
    
            objCliente.lCodigo = objPMPItens.lCliente
                
            'le o nome reduzido do cliente
            lErro = CF("Cliente_Le", objCliente)
            If lErro <> SUCESSO And lErro <> 12293 Then gError 136394
            
            If lErro <> SUCESSO Then
                GridItens.TextMatrix(iLinha, iGrid_Cliente_Col) = "INTERNO"
            Else
                'preenche com o nome reduzido do cliente
                GridItens.TextMatrix(iLinha, iGrid_Cliente_Col) = objCliente.sNomeReduzido
            End If
            
            GridItens.TextMatrix(iLinha, iGrid_CodOP_Col) = objPMPItens.sCodOPOrigem
            GridItens.TextMatrix(iLinha, iGrid_CodPV_Col) = CStr(objPMPItens.objItemOP.lCodPedido)
            GridItens.TextMatrix(iLinha, iGrid_DataNecess_Col) = Format(objPMPItens.dtDataNecessidade, "dd/mm/yyyy")
            GridItens.TextMatrix(iLinha, iGrid_DataProducao_Col) = Format(objPMPItens.objItemOP.dtDataFimProd, "dd/mm/yyyy")
            GridItens.TextMatrix(iLinha, iGrid_Prioridade_Col) = objPMPItens.iPrioridade
            GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col) = Formata_Estoque(objPMPItens.dQuantidade)
            GridItens.TextMatrix(iLinha, iGrid_UM_Col) = objPMPItens.sUM
            GridItens.TextMatrix(iLinha, iGrid_Versao_Col) = objPMPItens.sVersao
    
            If objPMPItens.iStatus <> 0 Then
                
                Call Obtem_Status(objPMPItens.iStatus, sStatus)
                
                GridItens.TextMatrix(iLinha, iGrid_Status_Col) = sStatus
    
            End If
            
        End If

    Next

    Call Grid_Refresh_Checkbox(objGridItens)

    objGridItens.iLinhasExistentes = iLinha
    
    Preenche_Grid_Itens = SUCESSO
    
    Exit Function

Erro_Preenche_Grid_Itens:

    Preenche_Grid_Itens = gErr
    
    Select Case gErr
    
        Case 136391, 136394
        
        Case 136392
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case 136393
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163133)

    End Select
    
    Exit Function

End Function

Private Sub BotaoVerEtapas_Click()
'Abre a tela de etapas de produção

Dim lErro As Long
Dim objPMPItem As ClassPMPItens
   
On Error GoTo Erro_BotaoVerEtapas_Click
    
    'Se não tiver linha selecionada => Erro
    If GridItens.Row = 0 Then gError 136395
    
    'Verifica se a linha selecionada está preenchida
    If Len(GridItens.TextMatrix(GridItens.Row, iGrid_Cliente_Col)) = 0 Then gError 137428
   
    Set objPMPItem = gobjPMP.colItens.Item(GridItens.Row)

    'Chama a tela de ordem de produção
    Call Chama_Tela_Modal("EtapasDaProducao", gobjPMP, objPMPItem)
    
    lErro = Atualizar
    If lErro <> SUCESSO Then gError 136495

    Exit Sub

Erro_BotaoVerEtapas_Click:

    Select Case gErr

        Case 136395
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 137428
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_PREENCHIDA", gErr)
        
        Case 136495

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163134)

    End Select
    
    Exit Sub
    
End Sub

Private Sub Obtem_Status(ByVal iStatus As Integer, sStatus As String)
'Obtém o Status resumido

Dim lErro As Long
    
On Error GoTo Erro_Obtem_Status

    Select Case iStatus
    
        Case PO_STATUS_AMBOS
            sStatus = PO_STATUS_AMBOS & SEPARADOR & PO_STATUS_NOME_AMBOS
    
        Case PO_STATUS_FALTAMATERIAL
            sStatus = PO_STATUS_FALTAMATERIAL & SEPARADOR & PO_STATUS_NOME_FALTAMATERIAL
    
        Case PO_STATUS_OK
            sStatus = PO_STATUS_OK & SEPARADOR & PO_STATUS_NOME_OK
    
        Case PO_STATUS_SOBRECARGA
            sStatus = PO_STATUS_SOBRECARGA & SEPARADOR & PO_STATUS_NOME_SOBRECARGA
    
    End Select
    
    Exit Sub

Erro_Obtem_Status:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163135)

    End Select
    
    Exit Sub
    
End Sub

Private Sub Mostra_StatusDetalhado()
'Exibe o Sttaus no text Box MultLine

Dim lErro As Long
    
On Error GoTo Erro_Mostra_StatusDetalhado
    
    If GridItens.Row <> 0 And GridItens.Row <= gobjPMP.colItens.Count Then

        StatusDetalhado.Text = gobjPMP.colItens.Item(GridItens.Row).sDescErro
    
    End If
    
    Exit Sub

Erro_Mostra_StatusDetalhado:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163136)

    End Select
    
    Exit Sub
    
End Sub

Function Limpa_Tela_MRP() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_MRP
        
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    StatusDetalhado.Text = ""
    DescClienteFinal.Caption = ""
    DescClienteInicial.Caption = ""
    DescProdFim.Caption = ""
    DescProdInic.Caption = ""
    
    Set gobjPMP = New ClassPMP
    Set gobjGeracaoOP = New ClassGeracaoOP

    Call Grid_Limpa(objGridItens)
        
    Call Grid_Limpa(objGridPV)
    
    ItensSelecionados.Caption = "0"
    
    iAlterado = 0
    iAlteradoVersao = 0
    iAlteradoTab = REGISTRO_ALTERADO
    iAlteradoFiltro = REGISTRO_ALTERADO
    
    TrazerPV.Value = vbUnchecked
    
    VersaoMRP.Text = ""
    
    'Torna Frame atual invisível
    Frame1(TabStrip1.SelectedItem.Index).Visible = False
    iFrameAtual = TAB_Selecao
    'Torna Frame atual visível
    Frame1(iFrameAtual).Visible = True
    TabStrip1.Tabs.Item(iFrameAtual).Selected = True
    
    Call TabStrip1_Click

    Limpa_Tela_MRP = SUCESSO

    Exit Function

Erro_Limpa_Tela_MRP:

    Limpa_Tela_MRP = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163137)

    End Select

    Exit Function

End Function

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 136396

    Call Limpa_Tela_MRP
    
    'Fecha o comando de setas
    Call ComandoSeta_Fechar(Me.Name)
       
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 136396

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163138)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGerar_Click()

Dim lErro As Long
    
On Error GoTo Erro_BotaoGerar_Click
    
    If gobjPMP.colItens.Count = 0 Then gError 136397
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = CF("MRP_Gera", gobjPMP)
    If lErro <> SUCESSO Then gError 136398
    
    'Limpa Tela
    Call Limpa_Tela_MRP

    'Le as versões cadastradas
    lErro = CF("PMPCombo_Le_Versoes", VersaoMRP)
    If lErro <> SUCESSO Then gError 136910
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoGerar_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 136397
            Call Rotina_Erro(vbOKOnly, "ERRO_PMP_SEM_DADOS_VALIDOS", gErr)
        
        Case 136398, 136910

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163139)

    End Select
    
    Exit Sub
    
End Sub

Function Move_TabSelecao_Memoria(ByVal objMRPSeleciona As ClassMRPSeleciona) As Long

Dim lErro As Long
Dim objCliente As ClassCliente
Dim sProduto As String
Dim iProdPreenchido As Integer

On Error GoTo Erro_Move_TabSelecao_Memoria

    Set objCliente = New ClassCliente
    
    gobjGeracaoOP.dtEntregaAte = DATA_NULA
    gobjGeracaoOP.dtEntregaDe = DATA_NULA
    gobjGeracaoOP.iFilialEmpresa = EMPRESA_TODA
    gobjGeracaoOP.dtDataOP = DATA_NULA
    gobjGeracaoOP.iIncluiPVcomOP = DESMARCADO
    
    gobjGeracaoOP.sOrdenacao = asOrdenacao(Ordenados.ListIndex)
    
    objCliente.sNomeReduzido = ClienteInicial.Text
    
    'Lê Cliente Inicial pelo NomeReduzido
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 136399
    
    objMRPSeleciona.lClienteInicial = objCliente.lCodigo
    gobjGeracaoOP.lClientesDe = objCliente.lCodigo
            
    Set objCliente = New ClassCliente

    objCliente.sNomeReduzido = ClienteFinal.Text
    
    'Lê Cliente Inicial pelo NomeReduzido
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 136400
    
    objMRPSeleciona.lClienteFinal = objCliente.lCodigo
    gobjGeracaoOP.lClientesAte = objCliente.lCodigo
         
    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProduto, iProdPreenchido)
    If lErro <> SUCESSO Then gError 136401
         
    objMRPSeleciona.sProdutoInicial = sProduto
    gobjGeracaoOP.sProdDe = sProduto
    
    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProduto, iProdPreenchido)
    If lErro <> SUCESSO Then gError 136402
         
    objMRPSeleciona.sProdutoFinal = sProduto
    gobjGeracaoOP.sProdAte = sProduto
    
    objMRPSeleciona.sOPFinal = OpFinal.Text
    objMRPSeleciona.sOPInicial = OpInicial.Text
    objMRPSeleciona.dtDataOPFinal = StrParaDate(DataOPFinal.Text)
    objMRPSeleciona.dtDataOPInicial = StrParaDate(DataOPInicial.Text)
         
    Move_TabSelecao_Memoria = SUCESSO

    Exit Function

Erro_Move_TabSelecao_Memoria:

    Move_TabSelecao_Memoria = gErr

    Select Case gErr
    
        Case 136399 To 136402

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163140)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se a versao está preenchida
    If Len(Trim(VersaoMRP.Text)) = 0 Then gError 136461
        
    gobjPMP.sVersao = VersaoMRP.Text
    
    lErro = Trata_Alteracao(gobjPMP, gobjPMP.sVersao)
    If lErro <> SUCESSO Then gError 137689
        
    'Grava o PMP Simulado no Banco de Dados
    lErro = CF("PMP_Simulacao_Grava", gobjPMP)
    If lErro <> SUCESSO Then gError 136462
    
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 136461
            Call Rotina_Erro(vbOKOnly, "ERRO_VERSAO_NAO_PREENCHIDO", gErr)

        Case 136462, 137689

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163141)

    End Select

    Exit Function

End Function

Private Sub VersaoMRP_Change()

    iAlterado = REGISTRO_ALTERADO
    iAlteradoVersao = REGISTRO_ALTERADO

End Sub

Private Sub VersaoMRP_Click()

Dim lErro As Long
Dim objPMP As New ClassPMP
    
On Error GoTo Erro_VersaoMRP_Click

    objPMP.sVersao = VersaoMRP.Text
    
    lErro = CF("PlanoMestreProducao_Le2", objPMP)
    If lErro <> SUCESSO And lErro <> 136473 Then gError 136475
    
    If lErro = SUCESSO Then
    
        Set gobjPMP = objPMP
    
        'iAlteradoTab = REGISTRO_ALTERADO
        
        lErro = Atualizar
        If lErro <> SUCESSO Then gError 136496
        
        iAlteradoVersao = 0
       
    End If
    
    Exit Sub

Erro_VersaoMRP_Click:

    Select Case gErr

        Case 136475, 136496

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163142)

    End Select
    
    Exit Sub

End Sub

Private Sub VersaoMRP_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objPMP As New ClassPMP
    
On Error GoTo Erro_VersaoMRP_Validate
    
    If iAlteradoVersao = REGISTRO_ALTERADO Then
    
        If Len(VersaoMRP.Text) > STRING_PMP_VERSAO Then gError 136479
        
        objPMP.sVersao = VersaoMRP.Text
        
        lErro = CF("PlanoMestreProducao_Le2", objPMP)
        If lErro <> SUCESSO And lErro <> 136473 Then gError 136476
        
        If lErro = SUCESSO Then
        
            Set gobjPMP = objPMP
            
            'AlteradoTab = REGISTRO_ALTERADO
         
            lErro = Atualizar
            If lErro <> SUCESSO Then gError 136497
        
            iAlteradoVersao = 0
        
        End If
        
    End If
    
    Exit Sub

Erro_VersaoMRP_Validate:

    Cancel = True

    Select Case gErr

        Case 136476, 136497
        
        Case 136479
            Call Rotina_Erro(vbOKOnly, "ERRO_PMP_VERSAO_TAMANHO", gErr, STRING_PMP_VERSAO)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163143)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoPreviaCarga_Click()
'Abre o relatório de Prévia de carga no CTs

Dim lErro As Long
    
On Error GoTo Erro_BotaoPreviaCarga_Click
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = CF("RelPreviaCargaCT_Prepara", gobjPMP)
    If lErro <> SUCESSO Then gError 136533
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoPreviaCarga_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 136533

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163144)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoErros_Click()
'Abre o relatório de críticas

Dim lErro As Long
    
On Error GoTo Erro_BotaoErros_Click
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = CF("RelErrosMRP_Prepara", gobjPMP)
    If lErro <> SUCESSO Then gError 136548
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoErros_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 136548

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163145)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoPreviaReqCompra_Click()
'Abre o relatório de prévia de requisição de compra

Dim lErro As Long
    
On Error GoTo Erro_BotaoPreviaReqCompra_Click
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = CF("RelPreviaReqCompra_Prepara", gobjPMP)
    If lErro <> SUCESSO Then gError 136549
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoPreviaReqCompra_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 136549

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163146)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoTrazer_Click()
'Traz os dados do MRP para tela

Dim lErro As Long

On Error GoTo Erro_BotaoTrazer_Click

    If gobjEST.dtDataInicioMRP = DATA_NULA Then gError 137698

    lErro = Atualizar(True)
    If lErro <> SUCESSO Then gError 136841
        
    Exit Sub

Erro_BotaoTrazer_Click:

    Select Case gErr
    
        Case 137698
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIO_MRP_NAO_PREENCHIDA", gErr)
    
        Case 136841
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163147)

    End Select
    
End Sub

Private Sub DataProducao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataProducao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub DataProducao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub DataProducao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DataProducao
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub BotaoGrafico_Click()
'Abre a tela de Cronograma Gráfico

Dim lErro As Long
Dim objTelaGrafico As New ClassTelaGrafico

On Error GoTo Erro_BotaoGrafico_Click:

    lErro = Atualiza_Cronograma(objTelaGrafico)
    If lErro <> SUCESSO Then gError 138246
    
    Call Chama_Tela_Modal("TelaGrafico", objTelaGrafico)

    lErro = Atualizar
    If lErro <> SUCESSO Then gError 138245

    Exit Sub

Erro_BotaoGrafico_Click:

    Select Case gErr
    
        Case 138245, 138246

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163148)

    End Select
    
    Exit Sub
    
End Sub

Public Function Atualiza_Cronograma(objTelaGrafico As ClassTelaGrafico, Optional bMarcaComErros As Boolean = False) As Long
'Acerta a tela de cronograma gráfico após o retorno da tela de etapas
'Para isso remonta o objTelaGrafico com os novos dados

Dim objTelaGraficoItem As New ClassTelaGraficoItens
Dim objPMPItem As ClassPMPItens
Dim objPO As ClassPlanoOperacional
Dim objPOAux As ClassPlanoOperacional
Dim iIndice As Integer
Dim iCont As Integer
Dim sNL As String
Dim objCT As ClassCentrodeTrabalho
Dim lErro As Long
Dim bPrimeira As Boolean
Dim objBotao As ClassTelaGraficoBotao

On Error GoTo Erro_Atualiza_Cronograma:

    Set objTelaGrafico.colBotoes = New Collection
    Set objTelaGrafico.colItens = New Collection
    Set objTelaGrafico.colParametros = New Collection

    Set objTelaGrafico.objTela = Me
    
    objTelaGrafico.sNomeTela = "Cronograma das Etapas de Produção"
    objTelaGrafico.iTamanhoDia = 540
    objTelaGrafico.iModal = MARCADO
    objTelaGrafico.iAtualizaRetornoClick = MARCADO
    objTelaGrafico.sNomeFuncAtualiza = "Atualiza_Cronograma"
    
    objTelaGrafico.colParametros.Add bMarcaComErros
    
    Set objBotao = New ClassTelaGraficoBotao

    objBotao.colParametros.Add objTelaGrafico
    objBotao.colParametros.Add True
    objBotao.sNome = "Marcar Etapas com Erro"
    objBotao.sNomeFuncao = "Atualiza_Cronograma"
    objBotao.sTextoExibicao = "Exibir Etapas com Erro"
    objBotao.iAtualizaRetornoClick = MARCADO

    objTelaGrafico.colBotoes.Add objBotao

    Set objBotao = New ClassTelaGraficoBotao

    objBotao.colParametros.Add objTelaGrafico
    objBotao.colParametros.Add False
    objBotao.sNome = "Desmarcar Etapas com Erro"
    objBotao.sNomeFuncao = "Atualiza_Cronograma"
    objBotao.sTextoExibicao = "Exibir Etapas com Erro"
    objBotao.iAtualizaRetornoClick = MARCADO

    objTelaGrafico.colBotoes.Add objBotao

    iIndice = 0
    
    sNL = Chr(13) & Chr(10)

    'Para cada item do Plano Mestre
    For Each objPMPItem In gobjPMP.colItens
    
        iIndice = iIndice + 1
    
        iCont = 0
    
        'Para cada etapa
        For Each objPO In objPMPItem.ColPO
        
            iCont = iCont + 1
        
            Set objTelaGraficoItem = New ClassTelaGraficoItens
            Set objCT = New ClassCentrodeTrabalho
        
            objTelaGraficoItem.colobj.Add gobjPMP
            objTelaGraficoItem.colobj.Add objPMPItem
            objTelaGraficoItem.colobj.Add objPO
            
            bPrimeira = True
            
            For Each objPOAux In objPMPItem.ColPO
                'Se existe outra com data de início menor ou igual desde que
                'o nó esteja depois na estrutura de árvore então o PO corrente
                'não é o primeiro
                If (objPOAux.dtDataInicio < objPO.dtDataInicio) Or (objPOAux.dtDataInicio = objPO.dtDataInicio And ((objPOAux.iNivel > objPO.iNivel) Or (objPOAux.iNivel = objPO.iNivel And objPOAux.iSeq > objPO.iSeq))) Then
                    bPrimeira = False
                End If
            Next
            
            'Se for a primeira é a final em termos de data
            If iCont = 1 Then
                objTelaGraficoItem.iIcone = TELA_GRAFICO_ICONE_FIM
            End If
                       
            If bPrimeira Then
                
                If iCont = 1 Then
                    objTelaGraficoItem.iIcone = TELA_GRAFICO_ICONE_INICIO_E_FIM
                Else
                    objTelaGraficoItem.iIcone = TELA_GRAFICO_ICONE_INICIO
                End If
                
            End If
            
            objCT.lNumIntDoc = objPO.lNumIntDocCT
            
            'Le o centro de trabalho
            lErro = CF("CentrodeTrabalho_Le_NumIntDoc", objCT)
            If lErro <> SUCESSO And lErro <> 134590 Then gError 138244
                    
            objTelaGraficoItem.dtDataFim = objPO.dtDataFim
            objTelaGraficoItem.dtDataInicio = objPO.dtDataInicio
            objTelaGraficoItem.sNomeTela = "EtapasDaProducao"
            objTelaGraficoItem.sTextoExibicao = "OP Origem: " & objPMPItem.sCodOPOrigem & sNL & "OP: " & objPO.sCodOPOrigem & sNL & "Prioridade: " & objPMPItem.iPrioridade & sNL & "Data Início: " & Format(objPO.dtDataInicio, "dd/mm/yyyy") & sNL & "Data Fim: " & Format(objPO.dtDataFim, "dd/mm/yyyy") & sNL & "CT: " & objCT.sNomeReduzido
            objTelaGraficoItem.sNome = objPO.sCodOPOrigem
            objTelaGraficoItem.iIndiceCor = iIndice
            
            If bMarcaComErros Then
                If objPO.iStatus <> PO_STATUS_OK Then
                    objTelaGraficoItem.lCor = vbRed
                Else
                    objTelaGraficoItem.lCor = vbGreen
                End If
            End If
            
            objTelaGrafico.colItens.Add objTelaGraficoItem
        
        Next
        
    Next
    
    lErro = Atualizar
    If lErro <> SUCESSO Then gError 141628
    
    Atualiza_Cronograma = SUCESSO

    Exit Function

Erro_Atualiza_Cronograma:

    Atualiza_Cronograma = gErr

    Select Case gErr
    
        Case 138244, 141628

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163149)

    End Select
    
    Exit Function
    
End Function

Private Sub BotaoOP_Click()

Dim lErro As Long
Dim objOP As New ClassOrdemDeProducao
   
On Error GoTo Erro_BotaoVerEtapas_Click
    
    'Se não tiver linha selecionada => Erro
    If GridItens.Row = 0 Then gError 138242
    
    'Verifica se a linha selecionada está preenchida
    If Len(GridItens.TextMatrix(GridItens.Row, iGrid_Cliente_Col)) = 0 Then gError 138243
   
    objOP.sCodigo = gobjPMP.colItens.Item(GridItens.Row).sCodOPOrigem
    objOP.iFilialEmpresa = gobjPMP.colItens.Item(GridItens.Row).iFilialEmpresa

    'Chama a tela de ordem de produção
    Call Chama_Tela("OrdemProducao", objOP)

    Exit Sub

Erro_BotaoVerEtapas_Click:

    Select Case gErr

        Case 138242
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 138243
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_PREENCHIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163150)

    End Select
    
    Exit Sub
    
End Sub

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
    objGridInt.colColuna.Add ("MRP")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Data Emissão")
    objGridInt.colColuna.Add ("Data Entrega")
    objGridInt.colColuna.Add ("Pedido de Venda")
    objGridInt.colColuna.Add ("Filial do Pedido")
    objGridInt.colColuna.Add ("Prioridade")

    'Controles que participam do GridPV
    objGridInt.colCampo.Add (GeraOP.Name)
    objGridInt.colCampo.Add (ProdutoPV.Name)
    objGridInt.colCampo.Add (DescricaoItem.Name)
    objGridInt.colCampo.Add (UMPV.Name)
    objGridInt.colCampo.Add (QuantidadePV.Name)
    objGridInt.colCampo.Add (ClientePV.Name)
    objGridInt.colCampo.Add (Filial.Name)
    objGridInt.colCampo.Add (DataEmissao.Name)
    objGridInt.colCampo.Add (DataEntrega.Name)
    objGridInt.colCampo.Add (Pedido.Name)
    objGridInt.colCampo.Add (FilialPV.Name)
    objGridInt.colCampo.Add (PrioridadePV.Name)

    'Grid do GridInterno
    objGridInt.objGrid = GridPV

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 10

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

    Inicializa_GridPV = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163151)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_PV(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_PV

    Select Case GridPV.Col

        Case GRIDPV_GERAOP_COL

            lErro = Saida_Celula_GeraOP(objGridInt)
            If lErro <> SUCESSO Then gError 138425

        Case GRIDPV_QUANTIDADE_COL

            lErro = Saida_Celula_QuantidadePV(objGridInt)
            If lErro <> SUCESSO Then gError 138426

    End Select

    Saida_Celula_PV = SUCESSO

    Exit Function

Erro_Saida_Celula_PV:

    Saida_Celula_PV = gErr

    Select Case gErr

        Case 138425 To 138426
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163152)

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

Private Sub ProdutoPV_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPV)
End Sub

Private Sub ProdutoPV_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPV)
End Sub

Private Sub ProdutoPV_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridPV.objControle = ProdutoPV
    lErro = Grid_Campo_Libera_Foco(objGridPV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UMPV_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPV)
End Sub

Private Sub UMPV_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPV)
End Sub

Private Sub UMPV_Validate(Cancel As Boolean)
    
Dim lErro As Long

    Set objGridPV.objControle = UMPV
    lErro = Grid_Campo_Libera_Foco(objGridPV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantidadePV_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPV)
End Sub

Private Sub QuantidadePV_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPV)
End Sub

Private Sub QuantidadePV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridPV.objControle = QuantidadePV
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

Private Sub ClientePV_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPV)
End Sub

Private Sub ClientePV_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPV)
End Sub

Private Sub ClientePV_Validate(Cancel As Boolean)
    
Dim lErro As Long

    Set objGridPV.objControle = ClientePV
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
    If lErro <> SUCESSO Then gError 138427

    Saida_Celula_GeraOP = SUCESSO

    Exit Function

Erro_Saida_Celula_GeraOP:

    Saida_Celula_GeraOP = gErr

    Select Case gErr

        Case 138427
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163153)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantidadePV(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantTotal As Double

On Error GoTo Erro_Saida_Celula_QuantidadePV

    Set objGridInt.objControle = QuantidadePV

    'se a quantidade foi preenchida
    If Len(Trim(QuantidadePV.ClipText)) > 0 Then

        lErro = Valor_Positivo_Critica(QuantidadePV.Text)
        If lErro <> SUCESSO Then gError 138423

        Quantidade.Text = Formata_Estoque(QuantidadePV.Text)

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 138424

    Saida_Celula_QuantidadePV = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantidadePV:

    Saida_Celula_QuantidadePV = gErr

    Select Case gErr

        Case 138423, 138424
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 163154)

    End Select

    Exit Function

End Function

Private Sub GeraOP_Click()

Dim iClick As Integer
Dim lErro As Long

On Error GoTo Erro_GeraOP_Click

    iAlteradoTab = REGISTRO_ALTERADO
    iAlteradoOPPV = REGISTRO_ALTERADO
    
    'Verifica se é alguma linha válida
    If GridPV.Row > objGridPV.iLinhasExistentes Then Exit Sub
    
'    If StrParaDate(GridPV.TextMatrix(GridPV.Row, GRIDPV_DATAENTREGA_COL)) = DATA_NULA Then
'        'GeraOP.Value = vbUnchecked
'        GridPV.TextMatrix(GridPV.Row, GRIDPV_GERAOP_COL) = "0"
'
'        lErro = Grid_Refresh_Checkbox(objGridPV)
'        If lErro <> SUCESSO Then gError 138422
'
'        gError 177038
'    End If

    'Verifica se está selecionando ou desselecionando
    If Len(Trim(GridPV.TextMatrix(GridPV.Row, GRIDPV_GERAOP_COL))) > 0 Then
        iClick = CInt(GridPV.TextMatrix(GridPV.Row, GRIDPV_GERAOP_COL)) = 1
    End If

    If iClick = True Then
        iSelecionadosPV = iSelecionadosPV + 1
    Else
        iSelecionadosPV = iSelecionadosPV - 1
    End If

    ItensSelecionados.Caption = CStr(iSelecionadosPV)

    lErro = Grid_Refresh_Checkbox(objGridPV)
    If lErro <> SUCESSO Then gError 138422

    Exit Sub

Erro_GeraOP_Click:

    Select Case gErr

        Case 138422
        
        Case 177038
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMPV_SEM_DATAENTREGA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163155)

    End Select

    Exit Sub

End Sub

Private Sub OPMRP_Click()

Dim iClick As Integer
Dim lErro As Long

On Error GoTo Erro_OPMRP_Click

    iAlteradoTab = REGISTRO_ALTERADO
    iAlteradoOPPV = REGISTRO_ALTERADO
    
    'Verifica se é alguma linha válida
    If GridOP.Row > objGridOP.iLinhasExistentes Then Exit Sub

    'Verifica se está selecionando ou desselecionando
    If Len(Trim(GridOP.TextMatrix(GridOP.Row, GRIDOP_OPMRP_COL))) > 0 Then
        iClick = StrParaInt(GridOP.TextMatrix(GridOP.Row, GRIDOP_OPMRP_COL)) = 1
    End If

    If iClick = True Then
        iSelecionadosOP = iSelecionadosOP + 1
    Else
        iSelecionadosOP = iSelecionadosOP - 1
    End If

    ItensSelecionadosOP.Caption = CStr(iSelecionadosOP)

    lErro = Grid_Refresh_Checkbox(objGridOP)
    If lErro <> SUCESSO Then gError 138422

    Exit Sub

Erro_OPMRP_Click:

    Select Case gErr

        Case 138422

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163156)

    End Select

    Exit Sub

End Sub

Private Sub BotaoMarcarTodos_Click()
'Marca todos os pedidos do Grid

Dim iLinha As Integer

    iAlteradoTab = REGISTRO_ALTERADO
    iAlteradoOPPV = REGISTRO_ALTERADO

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

Private Sub BotaoDesmarcarOP_Click()
'Desmarca todos os pedidos do Grid

Dim iLinha As Integer

    iAlteradoTab = REGISTRO_ALTERADO
    iAlteradoOPPV = REGISTRO_ALTERADO

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridOP.iLinhasExistentes

        'Desmarca na tela o pedido em questão
        GridOP.TextMatrix(iLinha, GRIDOP_OPMRP_COL) = S_DESMARCADO

    Next

    ItensSelecionadosOP.Caption = "0"

    iSelecionadosOP = 0

    'Atualiza na tela a checkbox desmarcada
    Call Grid_Refresh_Checkbox(objGridOP)

End Sub

Private Sub BotaoMarcarOP_Click()
'Marca todos os pedidos do Grid

Dim iLinha As Integer

    iAlteradoTab = REGISTRO_ALTERADO
    iAlteradoOPPV = REGISTRO_ALTERADO

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridOP.iLinhasExistentes

        'Marca na tela o pedido em questão
        GridOP.TextMatrix(iLinha, GRIDOP_OPMRP_COL) = S_MARCADO

    Next

    iSelecionadosOP = iLinha - 1

    ItensSelecionadosOP.Caption = CStr(iSelecionadosOP)

     'Atualiza na tela a checkbox marcada
    Call Grid_Refresh_Checkbox(objGridOP)

End Sub

Private Sub BotaoDesmarcarTodos_Click()
'Desmarca todos os pedidos do Grid

Dim iLinha As Integer

    iAlteradoTab = REGISTRO_ALTERADO
    iAlteradoOPPV = REGISTRO_ALTERADO

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

Private Sub Ordenados_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objMRPSeleciona As New ClassMRPSeleciona
Dim bFalse As Boolean

On Error GoTo Erro_Ordenados_Click

    If Ordenados.ListIndex = -1 Then Exit Sub

    'Verifica se a coleção de OP está vazia
    If gobjGeracaoOP.colItens.Count = 0 Then Exit Sub

    'Passa a Ordenaçao escolhida para o Obj
    gobjGeracaoOP.sOrdenacao = asOrdenacao(Ordenados.ListIndex)

    'Limpa a coleção
    Set gobjGeracaoOP = New ClassGeracaoOP

    lErro = Move_TabSelecao_Memoria(objMRPSeleciona)
    If lErro <> SUCESSO Then gError 138420

    bFalse = False

    'Preenche a Coleção
    lErro = CF("GeracaoOP_SelecionarItens", gobjGeracaoOP, bFalse)
    If lErro <> SUCESSO Then gError 138421

    'Limpa o GridPedido
    Call Grid_Limpa(objGridPV)

    'Preenche o GridPedido
    Call GridPV_Preenche(gobjGeracaoOP.colItens)

    Exit Sub

Erro_Ordenados_Click:

    Select Case gErr

        Case 138421, 138420

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163157)

    End Select

    Exit Sub

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

    Call Grid_Limpa(objGridPV)
    Call Grid_Limpa(objGridItens)

    iLinha = 0

    'Percorre todas as OP da Coleção
    For Each objGeracaoOPItem In colItens

        iLinha = iLinha + 1

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        lErro = Mascara_RetornaProdutoTela(objGeracaoOPItem.objItemOP.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 138414

        'Passa para a tela os dados da OP em questão
        GridPV.TextMatrix(iLinha, GRIDPV_GERAOP_COL) = CStr(objGeracaoOPItem.iSelecionado)
        GridPV.TextMatrix(iLinha, GRIDPV_PRODUTO_COL) = sProdutoMascarado
        GridPV.TextMatrix(iLinha, GRIDPV_DESCRICAOITEM_COL) = objGeracaoOPItem.sProdutoDescricao
        GridPV.TextMatrix(iLinha, GRIDPV_UM_COL) = objGeracaoOPItem.objItemOP.sSiglaUMEstoque
        
        lErro = CF("UM_Conversao", objGeracaoOPItem.objItemOP.iClasseUM, objGeracaoOPItem.objItemOP.sSiglaUM, objGeracaoOPItem.objItemOP.sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 138415
         
        GridPV.TextMatrix(iLinha, GRIDPV_QUANTIDADE_COL) = Formata_Estoque(objGeracaoOPItem.dQtdeFalta * dFator)
        GridPV.TextMatrix(iLinha, GRIDPV_CLIENTE_COL) = objGeracaoOPItem.sClienteReduzido
        
        objFilialCliente.iCodFilial = objGeracaoOPItem.iFilialCliente
        sNomeRedCliente = objGeracaoOPItem.sClienteReduzido
        
        'le nome da filial cliente
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sNomeRedCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 138416
        
        If lErro = 17660 Then gError 138417
        
        'PREENCHE CODIGO-NOME DA FILIAL
        GridPV.TextMatrix(iLinha, GRIDPV_FILIAL_COL) = CStr(objFilialCliente.iCodFilial) & SEPARADOR & objFilialCliente.sNome
        
        If objGeracaoOPItem.dtEmissaoPV <> DATA_NULA And objGeracaoOPItem.dtEmissaoPV <> 0 Then GridPV.TextMatrix(iLinha, GRIDPV_DATAEMISSAO_COL) = Format(objGeracaoOPItem.dtEmissaoPV, "dd/mm/yyyy")
        If objGeracaoOPItem.dtEntregaItemPV <> DATA_NULA And objGeracaoOPItem.dtEntregaItemPV <> 0 Then GridPV.TextMatrix(iLinha, GRIDPV_DATAENTREGA_COL) = Format(objGeracaoOPItem.dtEntregaItemPV, "dd/mm/yyyy")
        
        'preenche o objFilialEmpresa
        objFilialEmpresa.iCodFilial = objGeracaoOPItem.objItemOP.iFilialPedido
        
        'Le o nome da filial
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 138418
        
        'se não encontrou a filial ----- ERRO
        If lErro = 27378 Then gError 138419
        
        'CODIGO-NOME
        GridPV.TextMatrix(iLinha, GRIDPV_FILIALPV_COL) = CStr(objFilialEmpresa.iCodFilial) & SEPARADOR & objFilialEmpresa.sNome
                        
        GridPV.TextMatrix(iLinha, GRIDPV_PEDIDO_COL) = CStr(objGeracaoOPItem.objItemOP.lCodPedido)
        GridPV.TextMatrix(iLinha, GRIDPV_PRIORIDADE_COL) = objGeracaoOPItem.iPrioridade

    Next

    Call Grid_Refresh_Checkbox(objGridPV)

    'Passa para o Obj o número de itens passados pela Coleção
    objGridPV.iLinhasExistentes = colItens.Count

    Call Totaliza_Selecionados(colItens.Count)

    Exit Sub

Erro_GridPV_Preenche:

    Select Case gErr

        Case 138414, 138418, 138416, 138415
        
        Case 138419
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)
            
        Case 138417
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA2", gErr, sNomeRedCliente, objFilialCliente.iCodFilial)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163158)

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

Private Function PMP_Inclui_OP_Por_PV(ByVal objPMP As ClassPMP) As Long

Dim lErro As Long
Dim objOP As ClassOrdemDeProducao
Dim objItemOP As ClassItemOP
Dim iLinha As Integer
Dim objGeracaoOPItem As ClassGeracaoOPItem
Dim objPMPItem As ClassPMPItens
Dim objCliente As ClassCliente
Dim sCodigoOP As String
Dim iSeq As Integer
Dim iSeq2 As Integer
Dim objOPInsumo As ClassOrdemProducaoInsumos
Dim objOPOperacao As ClassOrdemProducaoOperacoes
Dim objAlmoxarifado As ClassAlmoxarifado
Dim objProdutoFilial As ClassProdutoFilial
Dim objPV As ClassPedidoDeVenda
Dim objItemPV As ClassItemPedido
Dim objItemPVData As ClassDataEntrega
Dim bContinua As Boolean
Dim iContador As Integer
Dim bAchouDatas As Boolean
Dim dQuantProd As Double
Dim colSaida As Collection
Dim colSaidaAux As Collection
Dim colCampos As Collection
Dim iIndice As Integer
Dim iStatus As Integer

On Error GoTo Erro_PMP_Inclui_OP_Por_PV

    For Each objGeracaoOPItem In gobjGeracaoOP.colItens
    
        iLinha = iLinha + 1

        If StrParaInt(GridPV.TextMatrix(iLinha, GRIDPV_GERAOP_COL)) = MARCADO Then

            Set objPV = New ClassPedidoDeVenda

            objPV.iFilialEmpresa = giFilialEmpresa
            objPV.lCodigo = objGeracaoOPItem.objItemOP.lCodPedido

            'Lê os dados do Pedido de Venda
            lErro = CF("PedidoDeVenda_Le_Todos_Completo", objPV, iStatus)
            If lErro <> SUCESSO Then gError 182844
            
            bAchouDatas = False
            For Each objItemPV In objPV.colItensPedido
                If objItemPV.sProduto = objGeracaoOPItem.objItemOP.sProduto Then
                    If objItemPV.colItemPVEntrega.Count <> 0 Then bAchouDatas = True
                    Exit For
                End If
            Next
            
            If bAchouDatas Then
            
                Set colSaida = New Collection
                Set colSaidaAux = New Collection
                Set colCampos = New Collection
                
                colCampos.Add "dtDataEntrega"
                
                'Ordem descrescente
                Call Ordena_Colecao(objItemPV.colItemPVEntrega, colSaidaAux, colCampos)

                For iIndice = colSaidaAux.Count To 1 Step -1
                    colSaida.Add colSaidaAux.Item(iIndice)
                Next

            End If
    
            bContinua = True
            iContador = 0
            dQuantProd = 0
            Do While bContinua
                
                If bAchouDatas Then
                    'Se é a última data de entrega a tratar-> Sai do loop
                    If objItemPV.colItemPVEntrega.Count = iContador + 1 Then bContinua = False
                    'Se já manou produzir tudo que precisava-> Sai do loop
                    If Abs(dQuantProd - objGeracaoOPItem.dQtdeFalta) < QTDE_ESTOQUE_DELTA Then bContinua = False
                Else
                    'Se não tem data de entrega-> Sai do loop
                    bContinua = False
                End If
                
                iContador = iContador + 1

                Set objPMPItem = New ClassPMPItens
                Set objOP = New ClassOrdemDeProducao
                Set objItemOP = New ClassItemOP
                Set objCliente = New ClassCliente
                Set objAlmoxarifado = New ClassAlmoxarifado
                Set objProdutoFilial = New ClassProdutoFilial
                
                lErro = CF("OrdemProducao_Automatico", sCodigoOP, giFilialEmpresa)
                If lErro <> SUCESSO Then gError 138411
                
                iSeq = iSeq - 1
                           
                'Preenche o objProduto Filial para poder ler o Almoxarifado Padrao
                objProdutoFilial.sProduto = objGeracaoOPItem.objItemOP.sProduto
                objProdutoFilial.iFilialEmpresa = giFilialEmpresa
                
                'Le o AlmoxarifadoPadrao
                lErro = CF("ProdutoFilial_Le", objProdutoFilial)
                If lErro <> SUCESSO And lErro <> 28261 Then gError 138430
                
                If lErro = 28261 Then gError 138431
                
                'Prenche para que seja Lido o Nome Reduzido do Almoxarifado Padrão
                objAlmoxarifado.iCodigo = objProdutoFilial.iAlmoxarifado
                
                lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25056 Then gError 138432
                            
                If lErro = 25056 Then gError 138433
                
                If bAchouDatas Then
                    objPMPItem.dtDataEntregaPV = colSaida.Item(iContador).dtDataEntrega
                    
                    If colSaida.Item(iContador).dQuantidade < objGeracaoOPItem.dQtdeFalta Then
                        objItemOP.dQuantidade = colSaida.Item(iContador).dQuantidade
                    Else
                        objItemOP.dQuantidade = objGeracaoOPItem.dQtdeFalta
                    End If
                    dQuantProd = dQuantProd + objItemOP.dQuantidade
                Else
                    objPMPItem.dtDataEntregaPV = objGeracaoOPItem.dtEntregaItemPV
                    objItemOP.dQuantidade = objGeracaoOPItem.dQtdeFalta
                End If
                
                If objPMPItem.dtDataEntregaPV <> DATA_NULA And objPMPItem.dtDataEntregaPV >= gdtDataAtual Then
                    objItemOP.dtDataFimProd = objPMPItem.dtDataEntregaPV
                    objItemOP.dtDataInicioProd = objPMPItem.dtDataEntregaPV
                Else
                    objItemOP.dtDataFimProd = gdtDataAtual
                    objItemOP.dtDataInicioProd = gdtDataAtual
                    
                    objPMPItem.iProduzLogo = MARCADO
                End If
                
                objItemOP.lNumIntDoc = iSeq
                objItemOP.iFilialEmpresa = giFilialEmpresa
                objItemOP.sCodigo = sCodigoOP
                objItemOP.iAlmoxarifado = objAlmoxarifado.iCodigo
                objItemOP.iItem = 1
                objItemOP.iSituacao = ITEMOP_SITUACAO_NORMAL
                objItemOP.iDestinacao = ITEMOP_DESTINACAO_PV
                objItemOP.iFilialPedido = objGeracaoOPItem.objItemOP.iFilialPedido
                objItemOP.lCodPedido = objGeracaoOPItem.objItemOP.lCodPedido
                objItemOP.sSiglaUM = objGeracaoOPItem.objItemOP.sSiglaUM
                objItemOP.sProduto = objGeracaoOPItem.objItemOP.sProduto
                objItemOP.sVersao = objGeracaoOPItem.objItemOP.sVersao
                objItemOP.iPrioridade = objGeracaoOPItem.iPrioridade
                
                objOP.iNumItens = 1
                objOP.iFilialEmpresa = giFilialEmpresa
                objOP.dtDataEmissao = gdtDataAtual
                objOP.iGeraReqCompra = DESMARCADO
                objOP.iTipo = OP_TIPO_OP
                objOP.sCodigo = sCodigoOP
                objOP.iStatusOP = ITEMOP_SITUACAO_PLANEJADA
                
                objOP.colItens.Add objItemOP
    
                lErro = CF("OP_Gera_Dados", objOP)
                If lErro <> SUCESSO Then gError 138412
                
                For Each objOPOperacao In objItemOP.colOrdemProducaoOperacoes
                
                    iSeq2 = iSeq2 - 1
                
                    objOPOperacao.lNumIntDoc = iSeq2
                    objOPOperacao.lNumIntDocItemOP = iSeq
                    
                    For Each objOPInsumo In objOPOperacao.colOPInsumos
                    
                        objOPInsumo.lNumIntDocOper = iSeq2
                    
                    Next
                    
                    objOPOperacao.objOperacoesTempo.lNumIntDocOperacao = iSeq2
                
                Next
                
                objPMPItem.dQuantidade = objItemOP.dQuantidade
                objPMPItem.dQuantidadeNecess = objItemOP.dQuantidade
                objPMPItem.dtDataNecessidade = objItemOP.dtDataFimProd
                objPMPItem.iFilialCli = objGeracaoOPItem.iFilialCliente
                objPMPItem.iFilialEmpresa = objItemOP.iFilialEmpresa
                objPMPItem.sProduto = objItemOP.sProduto
                objPMPItem.sUM = objItemOP.sSiglaUM
                objPMPItem.sVersao = objItemOP.sVersao
                objPMPItem.lCodGeracao = objPMP.lCodGeracao
                objPMPItem.sCodOPOrigem = sCodigoOP
                objPMPItem.iAlterado = REGISTRO_ALTERADO
                objPMPItem.iPrioridade = objItemOP.iPrioridade
                
                objCliente.sNomeReduzido = objGeracaoOPItem.sClienteReduzido
                
                lErro = CF("Cliente_Le_NomeReduzido", objCliente)
                If lErro <> SUCESSO And lErro <> 12348 Then gError 138413
                
                objPMPItem.lCliente = objCliente.lCodigo
                
                Set objPMPItem.objItemOP = objItemOP
                Set objPMPItem.objOP = objOP
                
                objPMP.colItens.Add objPMPItem
                
            Loop
            
        End If

    Next

    PMP_Inclui_OP_Por_PV = SUCESSO

    Exit Function

Erro_PMP_Inclui_OP_Por_PV:

    PMP_Inclui_OP_Por_PV = gErr

    Select Case gErr
    
        Case 138411 To 138413, 138430, 138432, 182844

        Case 138431
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTOFILIAL_INEXISTENTE", gErr, objProdutoFilial.sProduto, objProdutoFilial.iFilialEmpresa)
        
        Case 138432
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", gErr, objProdutoFilial.iAlmoxarifado)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163159)

    End Select

    Exit Function

End Function

Private Function PMP_Inclui_OP(ByVal objPMP As ClassPMP) As Long

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_PMP_Inclui_OP

    For iLinha = 1 To objGridOP.iLinhasExistentes
    
        If StrParaInt(GridOP.TextMatrix(iLinha, GRIDOP_OPMRP_COL)) = MARCADO Then

            gobjPMPOP.colItens.Item(iLinha).iAlterado = REGISTRO_ALTERADO

            objPMP.colItens.Add gobjPMPOP.colItens.Item(iLinha)
            
        End If

    Next

    PMP_Inclui_OP = SUCESSO

    Exit Function

Erro_PMP_Inclui_OP:

    PMP_Inclui_OP = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163160)

    End Select

    Exit Function

End Function

Function PMP_Retira_Sem_Quantidade(ByVal objPMP As ClassPMP) As Long

Dim lErro As Long
Dim objPMPItens As ClassPMPItens
Dim iIndice As Integer

On Error GoTo Erro_PMP_Retira_Sem_Quantidade

    For iIndice = objPMP.colItens.Count To 1 Step -1
    
        Set objPMPItens = objPMP.colItens.Item(iIndice)
        
        If objPMPItens.dQuantidade < QTDE_ESTOQUE_DELTA Then
        
            objPMP.colItens.Remove (iIndice)
        
        End If
    
    Next

    PMP_Retira_Sem_Quantidade = SUCESSO

    Exit Function

Erro_PMP_Retira_Sem_Quantidade:

    PMP_Retira_Sem_Quantidade = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163161)

    End Select

    Exit Function

End Function

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
    objGridInt.colColuna.Add ("MRP")
    objGridInt.colColuna.Add ("OP")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quant. OP")
    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("Previsão Início")
    objGridInt.colColuna.Add ("Previsão Fim")
    objGridInt.colColuna.Add ("Pedido de Venda")
    objGridInt.colColuna.Add ("Filial do Pedido")
    objGridInt.colColuna.Add ("Prioridade")

    'Controles que participam do GridPV
    objGridInt.colCampo.Add (OPMRP.Name)
    objGridInt.colCampo.Add (OPCodigo.Name)
    objGridInt.colCampo.Add (ProdutoOP.Name)
    objGridInt.colCampo.Add (DescricaoItemOP.Name)
    objGridInt.colCampo.Add (UMOP.Name)
    objGridInt.colCampo.Add (QuantOP.Name)
    objGridInt.colCampo.Add (Almoxarifado.Name)
    objGridInt.colCampo.Add (DataPrevisaoInicio.Name)
    objGridInt.colCampo.Add (DataPrevisaoFim.Name)
    objGridInt.colCampo.Add (PedidoDeVendaOP.Name)
    objGridInt.colCampo.Add (FilialPVOP.Name)
    objGridInt.colCampo.Add (PrioridadeOP.Name)

    objGridInt.objGrid = GridOP

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE + 1

    'linhas visiveis do grid sem contar com as linhas fixas
    objGridInt.iLinhasVisiveis = 10

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

    Inicializa_GridOP = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163162)

    End Select

    Exit Function

End Function

Private Sub GridOP_Preenche(ByVal objPMP As ClassPMP)

Dim lErro As Long
Dim iLinha As Integer
Dim sProdutoMascarado As String
Dim objFilialEmpresa As New AdmFiliais
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objPMPItem As ClassPMPItens
Dim objItemOP As ClassItemOP

On Error GoTo Erro_GridOP_Preenche

    Call Grid_Limpa(objGridOP)
    Call Grid_Limpa(objGridItens)
    
    Set gobjPMPOP = objPMP

    iLinha = 0

    'Percorre todas as OP da Coleção
    For Each objPMPItem In objPMP.colItens
    
        Set objItemOP = objPMPItem.objItemOP
        
        iLinha = iLinha + 1

        'Passa para a tela os dados da OP em questão
        sProdutoMascarado = String(STRING_PRODUTO, 0)

        'Mascara produto
        lErro = Mascara_RetornaProdutoTela(objItemOP.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 138501

        GridOP.TextMatrix(iLinha, GRIDOP_CODOP_COL) = objItemOP.sCodigo
        GridOP.TextMatrix(iLinha, GRIDOP_PRODUTO_COL) = sProdutoMascarado

        GridOP.TextMatrix(iLinha, GRIDOP_DESCRICAOITEMOP_COL) = objItemOP.sDescricao
        GridOP.TextMatrix(iLinha, GRIDOP_UM_COL) = objItemOP.sSiglaUM
            
        GridOP.TextMatrix(iLinha, GRIDOP_QUANT_COL) = Formata_Estoque(objItemOP.dQuantidade)
             
        Set objAlmoxarifado = New ClassAlmoxarifado

        'Prenche para que seja Lido o Nome Reduzido do Almoxarifado Padrão
        objAlmoxarifado.iCodigo = objItemOP.iAlmoxarifado
        
        lErro = CF("Almoxarifado_Le", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25056 Then gError 138502
                    
        If lErro = 25056 Then gError 138503
            
        'Prenche o Almoxarifado com o Almoxarifado Padrão
        GridOP.TextMatrix(iLinha, GRIDOP_ALMOXARIFADO_COL) = objAlmoxarifado.sNomeReduzido

        GridOP.TextMatrix(iLinha, GRIDOP_DATAPREVISAOINICIO_COL) = Format(objItemOP.dtDataInicioProd, "dd/mm/yyyy")
        GridOP.TextMatrix(iLinha, GRIDOP_DATAPREVISAOFIM_COL) = Format(objItemOP.dtDataFimProd, "dd/mm/yyyy")
            
        'preenche o objFilialEmpresa
        objFilialEmpresa.iCodFilial = objItemOP.iFilialPedido
    
        'Le o nome da filial
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 138504
    
        'se não encontrou a filial ----- ERRO
        If lErro = 27378 Then gError 138505
    
        'CODIGO-NOME
        GridOP.TextMatrix(iLinha, GRIDOP_FILIALPV_COL) = CStr(objFilialEmpresa.iCodFilial) & SEPARADOR & objFilialEmpresa.sNome
        
        GridOP.TextMatrix(iLinha, GRIDOP_PEDIDODEVENDA_COL) = CStr(objItemOP.lCodPedido)
        GridOP.TextMatrix(iLinha, GRIDOP_PRIORIDADE_COL) = objItemOP.iPrioridade

    Next

    Call Grid_Refresh_Checkbox(objGridOP)

    'Passa para o Obj o número de OP's passados pela Coleção
    objGridOP.iLinhasExistentes = iLinha

    Call Totaliza_SelecionadosOP(iLinha)

    Exit Sub

Erro_GridOP_Preenche:

    Select Case gErr

        Case 138501, 138502, 138504
        
        Case 138503
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", gErr, objAlmoxarifado.iCodigo)

        Case 138505
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)
               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163163)

    End Select

    Exit Sub

End Sub

Private Sub Totaliza_SelecionadosOP(iNumLinhas As Integer)

Dim iLinha As Integer
Dim iTotal As Integer

    iTotal = 0

    For iLinha = 1 To iNumLinhas

        If GridOP.TextMatrix(iLinha, GRIDOP_OPMRP_COL) = "1" Then iTotal = iTotal + 1

    Next

    iSelecionadosOP = iTotal

    ItensSelecionadosOP.Caption = CStr(iTotal)

End Sub

Private Sub TrazerPV_Click()
    iAlteradoTab = REGISTRO_ALTERADO
    iAlteradoFiltro = REGISTRO_ALTERADO
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

    Call Chama_Tela_Modal("DataEntrega", sProdutoTela, dQuantidade, objItemPV.colItemPVEntrega, DESMARCADO)

    Exit Sub

Erro_BotaoEntrega_Click:

    Select Case gErr

        Case 182841
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 182842, 182844, 182845

        Case 182843
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182846)

    End Select

    Exit Sub
    
End Sub

