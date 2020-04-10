VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl OrdemCorte 
   ClientHeight    =   5565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9345
   KeyPreview      =   -1  'True
   ScaleHeight     =   5565
   ScaleWidth      =   9345
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "'"
      Height          =   4560
      Index           =   1
      Left            =   105
      TabIndex        =   11
      Top             =   870
      Width           =   9060
      Begin VB.Frame Frame5 
         Caption         =   "Relatórios"
         Height          =   1395
         Left            =   60
         TabIndex        =   35
         Top             =   3090
         Width           =   8985
         Begin VB.OptionButton OpcaoRelatorio 
            Caption         =   "Imprimir Ambos."
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
            Left            =   630
            TabIndex        =   46
            Top             =   1050
            Value           =   -1  'True
            Width           =   1665
         End
         Begin VB.OptionButton OpcaoRelatorio 
            Caption         =   "Imprimir Rótulos para Corte."
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
            Left            =   630
            TabIndex        =   45
            Top             =   690
            Width           =   3045
         End
         Begin VB.OptionButton OpcaoRelatorio 
            Caption         =   "Imprimir Ordem de Corte."
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
            Left            =   630
            TabIndex        =   44
            Top             =   330
            Width           =   2775
         End
         Begin VB.CheckBox ImprimeAoGravar 
            Caption         =   "Imprimir ao Gravar uma O.C."
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
            Left            =   5295
            TabIndex        =   36
            Top             =   690
            Width           =   2715
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ordem de Corte"
         Height          =   1455
         Left            =   75
         TabIndex        =   27
         Top             =   -15
         Width           =   8970
         Begin VB.TextBox Codigo 
            Height          =   285
            Left            =   1365
            MaxLength       =   6
            TabIndex        =   28
            Top             =   285
            Width           =   1350
         End
         Begin MSComCtl2.UpDown UpDownData 
            Height          =   300
            Left            =   5055
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   270
            Width           =   225
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox Data 
            Height          =   300
            Left            =   3990
            TabIndex        =   30
            Top             =   270
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
            Caption         =   "Status:"
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
            Left            =   6315
            TabIndex        =   34
            Top             =   330
            Width           =   615
         End
         Begin VB.Label StatusOP 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6975
            TabIndex        =   33
            Top             =   270
            Width           =   1305
         End
         Begin VB.Label CodigoOPLabel 
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
            Left            =   630
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   32
            Top             =   330
            Width           =   660
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
            Left            =   3465
            TabIndex        =   31
            Top             =   330
            Width           =   480
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Padrões "
         Height          =   1590
         Left            =   75
         TabIndex        =   12
         Top             =   1470
         Width           =   8985
         Begin VB.ComboBox DestinacaoPadrao 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "OrdemCorte.ctx":0000
            Left            =   2955
            List            =   "OrdemCorte.ctx":0002
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1125
            Width           =   1710
         End
         Begin MSMask.MaskEdBox CclPadrao 
            Height          =   315
            Left            =   2955
            TabIndex        =   14
            Top             =   225
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox AlmoxPadrao 
            Height          =   315
            Left            =   6990
            TabIndex        =   15
            Top             =   225
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownInicio 
            Height          =   300
            Left            =   4035
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   675
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataInicioPadrao 
            Height          =   300
            Left            =   2955
            TabIndex        =   17
            Top             =   675
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownFim 
            Height          =   300
            Left            =   8070
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   675
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataFimPadrao 
            Height          =   300
            Left            =   6990
            TabIndex        =   19
            Top             =   675
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PrioridadePadrao 
            Height          =   315
            Left            =   6990
            TabIndex        =   20
            Top             =   1125
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   3
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin VB.Label CclPadraoLabel 
            AutoSize        =   -1  'True
            Caption         =   "Centro de Custo/Lucro Padrão:"
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
            Left            =   255
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   26
            Top             =   285
            Width           =   2670
         End
         Begin VB.Label AlmoxPadraoLabel 
            AutoSize        =   -1  'True
            Caption         =   "Almoxarifado Padrão:"
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
            Left            =   5130
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   25
            Top             =   285
            Width           =   1815
         End
         Begin VB.Label DataPrevIniLbl 
            AutoSize        =   -1  'True
            Caption         =   "Data de Previsão de Início:"
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
            Left            =   555
            TabIndex        =   24
            Top             =   735
            Width           =   2370
         End
         Begin VB.Label DataPrevFimLbl 
            AutoSize        =   -1  'True
            Caption         =   "Data de Previsão de Fim:"
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
            Left            =   4785
            TabIndex        =   23
            Top             =   735
            Width           =   2160
         End
         Begin VB.Label DestPadraoLbl 
            AutoSize        =   -1  'True
            Caption         =   "Destinação Padrão:"
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
            Left            =   1230
            TabIndex        =   22
            Top             =   1185
            Width           =   1695
         End
         Begin VB.Label PrioridadePadraoLbl 
            AutoSize        =   -1  'True
            Caption         =   " Prioridade Padrão:"
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
            Left            =   5295
            TabIndex        =   21
            Top             =   1185
            Width           =   1650
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4530
      Index           =   2
      Left            =   150
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   9030
      Begin VB.ComboBox Versao 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "OrdemCorte.ctx":0004
         Left            =   6615
         List            =   "OrdemCorte.ctx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   2370
         Width           =   1875
      End
      Begin VB.ComboBox Destinacao 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "OrdemCorte.ctx":0008
         Left            =   6840
         List            =   "OrdemCorte.ctx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   1155
         Width           =   1830
      End
      Begin VB.CheckBox Benef 
         Height          =   210
         Left            =   4605
         TabIndex        =   53
         Top             =   2760
         Width           =   870
      End
      Begin VB.TextBox FilialCliente 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   240
         Left            =   4185
         TabIndex        =   52
         Text            =   "Filial do Cliente"
         Top             =   2430
         Width           =   1110
      End
      Begin VB.TextBox Cliente 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   3240
         TabIndex        =   51
         Text            =   "Cliente"
         Top             =   2130
         Width           =   1260
      End
      Begin VB.TextBox UnidadeMed 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   1065
         MaxLength       =   50
         TabIndex        =   50
         Top             =   465
         Width           =   600
      End
      Begin VB.TextBox DescricaoItem 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   5865
         MaxLength       =   50
         TabIndex        =   49
         Top             =   1980
         Width           =   2600
      End
      Begin VB.ComboBox Situacao 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "OrdemCorte.ctx":000C
         Left            =   4935
         List            =   "OrdemCorte.ctx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   1170
         Width           =   1830
      End
      Begin VB.ComboBox ComboFilialPedido 
         Height          =   315
         ItemData        =   "OrdemCorte.ctx":0010
         Left            =   2190
         List            =   "OrdemCorte.ctx":0012
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   1620
         Width           =   1875
      End
      Begin MSMask.MaskEdBox Maquina 
         Height          =   255
         Left            =   6885
         TabIndex        =   56
         Top             =   2730
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox PedidoDeVendaId 
         Height          =   255
         Left            =   4935
         TabIndex        =   57
         Top             =   1620
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Prioridade 
         Height          =   255
         Left            =   6660
         TabIndex        =   58
         Top             =   1620
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Quantidade 
         Height          =   285
         Left            =   2040
         TabIndex        =   59
         Top             =   510
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   503
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
      Begin MSMask.MaskEdBox Produto 
         Height          =   255
         Left            =   345
         TabIndex        =   60
         Top             =   1680
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Almoxarifado 
         Height          =   270
         Left            =   3585
         TabIndex        =   61
         Top             =   645
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Ccl 
         Height          =   270
         Left            =   5325
         TabIndex        =   62
         Top             =   750
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   476
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
         Left            =   3525
         TabIndex        =   63
         Top             =   1200
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
         Left            =   7455
         TabIndex        =   64
         Top             =   735
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
      Begin MSMask.MaskEdBox MetragemCons 
         Height          =   285
         Left            =   1305
         TabIndex        =   65
         Top             =   975
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   503
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
      Begin MSMask.MaskEdBox Enfesto 
         Height          =   285
         Left            =   2505
         TabIndex        =   66
         Top             =   930
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   503
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
      Begin MSMask.MaskEdBox Risco 
         Height          =   285
         Left            =   2115
         TabIndex        =   67
         Top             =   1245
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   503
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
      Begin VB.CommandButton BotaoMaquinas 
         Caption         =   "Maquinas"
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
         Left            =   7320
         TabIndex        =   6
         Top             =   4170
         Width           =   1680
      End
      Begin VB.CommandButton BotaoPedidoDeVenda 
         Caption         =   "Pedido de Venda"
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
         Left            =   5505
         TabIndex        =   5
         Top             =   4170
         Width           =   1680
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
         Height          =   330
         Left            =   3690
         TabIndex        =   4
         Top             =   4170
         Width           =   1680
      End
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
         Left            =   105
         TabIndex        =   3
         Top             =   4170
         Width           =   1680
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
         Height          =   330
         Left            =   1890
         TabIndex        =   2
         Top             =   4170
         Width           =   1680
      End
      Begin VB.CommandButton BotaoGrade 
         Caption         =   "Grade"
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
         Left            =   105
         TabIndex        =   1
         Top             =   3705
         Width           =   1680
      End
      Begin MSFlexGridLib.MSFlexGrid GridMovimentos 
         Height          =   3285
         Left            =   105
         TabIndex        =   7
         Top             =   315
         Width           =   8880
         _ExtentX        =   15663
         _ExtentY        =   5794
         _Version        =   393216
         Rows            =   21
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin VB.Label QuantDisponivel 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7620
         TabIndex        =   10
         Top             =   3720
         Width           =   1365
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade Disponível:"
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
         Left            =   5565
         TabIndex        =   9
         Top             =   3765
         Width           =   2025
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Material a ser Cortado"
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
         TabIndex        =   8
         Top             =   90
         Width           =   1890
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6555
      ScaleHeight     =   495
      ScaleWidth      =   2610
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   90
      Width           =   2670
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   120
         Picture         =   "OrdemCorte.ctx":0014
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Imprimir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   615
         Picture         =   "OrdemCorte.ctx":0116
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   1125
         Picture         =   "OrdemCorte.ctx":0270
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1620
         Picture         =   "OrdemCorte.ctx":03FA
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   2100
         Picture         =   "OrdemCorte.ctx":092C
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4980
      Left            =   75
      TabIndex        =   43
      Top             =   510
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   8784
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Principais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Itens"
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
Attribute VB_Name = "OrdemCorte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iLinhaAntiga As Integer
Dim iCodigoAlterado As Integer

Dim iPrestServAlterado As Integer

Dim gcolItemOP As Collection
'criado por causa da forma como foi construido a tela romaneiograde
Dim gobjOP As ClassOrdemDeProducao

Dim objGrid As AdmGrid
Dim iGrid_Produto_Col As Integer
Dim iGrid_UnidadeMed_Col  As Integer
Dim iGrid_Quantidade_Col  As Integer
Dim iGrid_Almoxarifado_Col  As Integer
Dim iGrid_Benef_Col  As Integer
Dim iGrid_Ccl_Col  As Integer
Dim iGrid_DescricaoItem_Col  As Integer
Dim iGrid_DataPrevInicio_Col  As Integer
Dim iGrid_DataPrevFim_Col  As Integer
Dim iGrid_Situacao_Col  As Integer
Dim iGrid_Destinacao_Col  As Integer
Dim iGrid_PedidoDeVenda_Col  As Integer
Dim iGrid_FilialPedido_Col  As Integer
Dim iGrid_Cliente_Col As Integer
Dim iGrid_FilialCliente_Col As Integer
Dim iGrid_Prioridade_Col  As Integer
Dim iGrid_Versao_Col  As Integer
Dim iGrid_Maquina_Col  As Integer
Dim iGrid_MetragemCons_Col As Integer
Dim iGrid_Enfesto_Col  As Integer
Dim iGrid_Risco_Col  As Integer

Dim iFrameAtual As Integer

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1
Private WithEvents objEventoCclPadrao As AdmEvento
Attribute objEventoCclPadrao.VB_VarHelpID = -1
Private WithEvents objEventoAlmoxPadrao As AdmEvento
Attribute objEventoAlmoxPadrao.VB_VarHelpID = -1
Private WithEvents objEventoEstoque As AdmEvento
Attribute objEventoEstoque.VB_VarHelpID = -1
Private WithEvents objEventoPedidoDeVenda As AdmEvento
Attribute objEventoPedidoDeVenda.VB_VarHelpID = -1
Private WithEvents objEventoMaquina As AdmEvento
Attribute objEventoMaquina.VB_VarHelpID = -1
Private WithEvents objEventoPrestServ As AdmEvento
Attribute objEventoPrestServ.VB_VarHelpID = -1

Private Function Inicializa_GridMovimentos() As Long

Dim iIndice As Integer

    Set objGrid = New AdmGrid

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Produto")
    objGrid.colColuna.Add ("Versão")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("U.M.")
    objGrid.colColuna.Add ("Quantidade")
    objGrid.colColuna.Add ("Metragem")
    objGrid.colColuna.Add ("Enfesto")
    objGrid.colColuna.Add ("Risco")
    objGrid.colColuna.Add ("Almoxarifado")
    objGrid.colColuna.Add ("Benef.")
    objGrid.colColuna.Add ("Ccl")
    objGrid.colColuna.Add ("Previsão Início")
    objGrid.colColuna.Add ("Previsão Fim")
    objGrid.colColuna.Add ("Situação")
    objGrid.colColuna.Add ("Destinação")
    objGrid.colColuna.Add ("Pedido de Venda")
    objGrid.colColuna.Add ("Filial do Pedido")
    objGrid.colColuna.Add ("Cliente")
    objGrid.colColuna.Add ("Filial do Cliente")
    objGrid.colColuna.Add ("Prioridade")
    objGrid.colColuna.Add ("Maquina")

    'Controles que participam do Grid
    objGrid.colCampo.Add (Produto.Name)
    objGrid.colCampo.Add (Versao.Name)
    objGrid.colCampo.Add (DescricaoItem.Name)
    objGrid.colCampo.Add (UnidadeMed.Name)
    objGrid.colCampo.Add (Quantidade.Name)
    objGrid.colCampo.Add (MetragemCons.Name)
    objGrid.colCampo.Add (Enfesto.Name)
    objGrid.colCampo.Add (Risco.Name)
    objGrid.colCampo.Add (Almoxarifado.Name)
    objGrid.colCampo.Add (Benef.Name)
    objGrid.colCampo.Add (Ccl.Name)
    objGrid.colCampo.Add (DataPrevisaoInicio.Name)
    objGrid.colCampo.Add (DataPrevisaoFim.Name)
    objGrid.colCampo.Add (Situacao.Name)
    objGrid.colCampo.Add (Destinacao.Name)
    objGrid.colCampo.Add (PedidoDeVendaId.Name)
    objGrid.colCampo.Add (ComboFilialPedido.Name)
    objGrid.colCampo.Add (Cliente.Name)
    objGrid.colCampo.Add (FilialCliente.Name)
    objGrid.colCampo.Add (Prioridade.Name)
    objGrid.colCampo.Add (Maquina.Name)

    'Colunas do Grid
    iGrid_Produto_Col = 1
    iGrid_Versao_Col = 2
    iGrid_DescricaoItem_Col = 3
    iGrid_UnidadeMed_Col = 4
    iGrid_Quantidade_Col = 5
    iGrid_MetragemCons_Col = 6
    iGrid_Enfesto_Col = 7
    iGrid_Risco_Col = 8
    iGrid_Almoxarifado_Col = 9
    iGrid_Benef_Col = 10
    iGrid_Ccl_Col = 11
    iGrid_DataPrevInicio_Col = 12
    iGrid_DataPrevFim_Col = 13
    iGrid_Situacao_Col = 14
    iGrid_Destinacao_Col = 15
    iGrid_PedidoDeVenda_Col = 16
    iGrid_FilialPedido_Col = 17
    iGrid_Cliente_Col = 18
    iGrid_FilialCliente_Col = 19
    iGrid_Prioridade_Col = 20
    iGrid_Maquina_Col = 21

    objGrid.objGrid = GridMovimentos

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 8

    'Largura da primeira coluna
    GridMovimentos.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    objGrid.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridMovimentos = SUCESSO

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim iIndice As Integer
Dim sUnidadeMed As String
Dim sCodProduto As String
Dim objProduto As New ClassProduto
Dim objClasseUM As New ClassClasseUM
Dim objUnidadeDeMedida As ClassUnidadeDeMedida
Dim colSiglas As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iNum As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    If iLocalChamada <> ROTINA_GRID_ABANDONA_CELULA Then

        'Verifica se produto está preenchido
        sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)

        lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 21922

        If gcolItemOP.Count >= GridMovimentos.Row Then
            iNum = gcolItemOP.Item(GridMovimentos.Row)
        Else
            iNum = 0
        End If

        'Pesquisa o controle da coluna em questão
        Select Case objControl.Name
    
            'Produto
            Case Produto.Name
    
                If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                    
                    Produto.Enabled = False
                
                Else
                    Produto.Enabled = True
                End If
    
            Case UnidadeMed.Name, DescricaoItem.Name, Cliente.Name, FilialCliente.Name
                'ficam sempre desabilitadas
    
            Case Versao.Name
                
                If Len(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Versao_Col)) > 0 Then
                    'Impede a atualização dos demais dados da OP
                    If iProdutoPreenchido = PRODUTO_PREENCHIDO And iNum = 0 And Left(GridMovimentos.TextMatrix(GridMovimentos.Row, 0), 1) <> "#" Then
                        objControl.Enabled = True
                        Call Carrega_ComboVersoes(sProdutoFormatado)
                    Else
                        objControl.Enabled = False
                    End If
                Else
                        objControl.Enabled = False
                End If
            
            Case Situacao.Name, Maquina.Name, MetragemCons.Name, Enfesto.Name, Risco.Name
    
                'habilita atualização da situacao da OP
                If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                    objControl.Enabled = True
                Else
                    objControl.Enabled = False
                End If
    
            Case ComboFilialPedido.Name, PedidoDeVendaId.Name
                
                If Destinacao.ListIndex <> -1 Then
                    If Destinacao.ItemData(Destinacao.ListIndex) = ITEMOP_DESTINACAO_PV Then
                        'Impede a atualização dos demais dados da OP
                        If iProdutoPreenchido = PRODUTO_PREENCHIDO And iNum = 0 Then
                            objControl.Enabled = True
                        Else
                            objControl.Enabled = False
                        End If
                    Else
                        objControl.Enabled = False
                    End If
                Else
                    objControl.Enabled = False
                   
                End If
            
            Case Destinacao.Name
    
                    If GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Benef_Col) = CStr(MARCADO) Then
                        objControl.Enabled = False
                    Else
                        objControl.Enabled = True
                   End If
    
            Case Quantidade.Name, Almoxarifado.Name
                If iProdutoPreenchido = PRODUTO_PREENCHIDO And iNum = 0 And Left(GridMovimentos.TextMatrix(GridMovimentos.Row, 0), 1) <> "#" Then
                    objControl.Enabled = True
                Else
                    objControl.Enabled = False
                End If
    
            Case Else
                'Impede a atualização dos demais dados da OP
                If iProdutoPreenchido = PRODUTO_PREENCHIDO And iNum = 0 Then
                    objControl.Enabled = True
                Else
                    objControl.Enabled = False
                End If
    
        End Select

    End If

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 21922

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163712)

    End Select

    Exit Sub

End Sub

Private Sub Almoxarifado_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Almoxarifado_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Almoxarifado_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Almoxarifado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Almoxarifado
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub AlmoxPadraoLabel_Click()

Dim colSelecao As New Collection
Dim objAlmoxarifado As New ClassAlmoxarifado

    Call Chama_Tela("AlmoxarifadoLista_Consulta", colSelecao, objAlmoxarifado, objEventoAlmoxPadrao)

End Sub

Private Sub BotaoGrade_Click()

Dim lErro  As Long
Dim objRomaneioGrade As ClassRomaneioGrade
Dim objItemOP As ClassItemOP

On Error GoTo Erro_BotaoGrade_Click

    If GridMovimentos.Row > 0 And GridMovimentos.Row <= objGrid.iLinhasExistentes Then
    
        Set objItemOP = gobjOP.colItens(GridMovimentos.Row)
        
        If objItemOP.iPossuiGrade = DESMARCADO Then gError 126517
            
        objItemOP.sAlmoxarifadoNomeRed = AlmoxPadrao.Text
            
        Set objRomaneioGrade = New ClassRomaneioGrade
        
        objRomaneioGrade.sNomeTela = Me.Name
        Set objRomaneioGrade.objObjetoTela = objItemOP
        Set objRomaneioGrade.objTela = Me
                    
        Call Chama_Tela_Modal("RomaneioGrade", objRomaneioGrade)
    
        Call Atualiza_Grid_Movimentos(objItemOP)
            
    End If
    
    Exit Sub

Erro_BotaoGrade_Click:

    Select Case gErr
      
        Case 126517
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_NAO_GRADE", gErr, GridMovimentos.Row)
      
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163713)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoMaquinas_Click()
'Chama o Browser de Maquinas...

Dim lErro As Long
Dim objMaquina As New ClassMaquinas
Dim sProduto As String
Dim iPreenchido As Integer
Dim colSelecao As Collection

On Error GoTo Erro_BotaoMaquinas_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridMovimentos.Row = 0 Then gError 106320

    'Se o equipamento foi preenchido => armazena no obj
    If Len(Trim(Maquina.Text)) > 0 Then
    
        If IsNumeric(Maquina.Text) Then
    
            objMaquina.iCodigo = StrParaInt(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Maquina_Col))
        
        Else
            
            objMaquina.sNomeReduzido = CStr(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Maquina_Col))
        
        End If
    
    End If

    'Lista de Equipamentos
    Call Chama_Tela("MaquinasLista", colSelecao, objMaquina, objEventoMaquina)

    Exit Sub

Erro_BotaoMaquinas_Click:

    Select Case gErr

        Case 106320
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 106433
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_NOME_INEXISTENTE", gErr, Maquina.Text)
            
        Case 106434
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_CODIGO_INEXISTENTE", gErr, Maquina.Text, giFilialEmpresa)

        Case 55325, 106432

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163714)

    End Select

End Sub

Private Sub BotaoImprimir_Click()

Dim lErro As Long
Dim objOrdemDeProducao As New ClassOrdemDeProducao

On Error GoTo Erro_BotaoImprimir_Click

    'Verifica se os campos obrigatórios foram preenchidos
    If Len(Trim(Codigo.Text)) = 0 Then gError 57722
    If Len(Trim(Data.ClipText)) = 0 Then gError 57723
    
    objOrdemDeProducao.sCodigo = Codigo.Text
    objOrdemDeProducao.dtDataEmissao = StrParaDate(Data.Text)
    objOrdemDeProducao.iFilialEmpresa = giFilialEmpresa
    
    'lErro = CF("OrdemDeProducao_TestaExistencia", objOrdemDeProducao)
    'If lErro <> SUCESSO And lErro <> 57721 Then gError 57724
    'If lErro = 57721 Then gError 57725
    
    'pesquisa a op nao baixada, e preenche seus itens
    lErro = CF("OrdemDeProducao_Le_ComItens", objOrdemDeProducao)
    If lErro <> SUCESSO And lErro <> 21960 Then gError 57724
    
    If lErro = 21960 Then
    
        'se nao achou antes, tenta ver se ja estava baixada
        lErro = CF("OrdemDeProducaoBaixada_Le_ComItens", objOrdemDeProducao)
        If lErro <> SUCESSO And lErro <> 82797 Then gError 111823
    
        'se nao achou => erro de inexistencia de op
        If lErro <> SUCESSO Then gError 57725
    
    End If
    
    'Executa o(s) Relatorio(s) de acordo com a selecao
    lErro = Executa_Relatorio(objOrdemDeProducao)
    If lErro <> SUCESSO Then gError 106409
    
    Exit Sub
    
Erro_BotaoImprimir_Click:

    Select Case gErr
    
        Case 57722
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGOOP_NAO_PREENCHIDO", gErr)
        
        Case 57723
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)
    
        Case 57724, 106409, 111823
        
        Case 57725
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_INEXISTENTE", gErr, objOrdemDeProducao.sCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163715)

    End Select
    
End Sub

Private Sub BotaoPedidoDeVenda_Click()

Dim lErro As Long
Dim objItemPedido As New ClassItemPedido
Dim colSelecao As New Collection
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoPedidoDeVenda_Click

    If GridMovimentos.Row = 0 Then gError 52016

    sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 52017

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
        'verifica se a destinação é Pedido de Venda - - Se não for - - > Erro
        If GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Destinacao_Col) <> STRING_PEDIDOVENDA Then gError 52018
        
        'Selecao
        colSelecao.Add sProdutoFormatado
        
        'chama a tela de lista de estoque do produto corrente
        Call Chama_Tela("ItemPVProdutoLista", colSelecao, objItemPedido, objEventoPedidoDeVenda)
    Else
        Error 52019
    End If

    Exit Sub

Erro_BotaoPedidoDeVenda_Click:

    Select Case gErr
        
        Case 52016
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 52017
                
        Case 52018
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESTINACAO_DEPENDENTE", gErr)
        
        Case 52019
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163716)

    End Select

    Exit Sub
End Sub

Private Sub Benef_Click()
    
Dim lErro As Long

On Error GoTo Erro_Benef_Click

    iAlterado = REGISTRO_ALTERADO

    If GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Benef_Col) = CStr(MARCADO) Then
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Destinacao_Col) = "Estoque"
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_PedidoDeVenda_Col) = ""
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Cliente_Col) = ""
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialPedido_Col) = ""
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialCliente_Col) = ""
    End If

    lErro = QuantDisponivel_Calcula(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col))
    If lErro <> SUCESSO Then gError 91274

    Exit Sub
    
Erro_Benef_Click:

    Select Case gErr
    
        Case 91274
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163717)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Benef_GotFocus()
'trata o evento gotfocus associado ao campo Benef

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Benef_KeyPress(KeyAscii As Integer)
'trata o evento keypress associado ao campo Benef

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Benef_Validate(Cancel As Boolean)
'trata o evento validate associado ao campo Benef

Dim lErro As Long

    Set objGrid.objControle = Benef
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Ccl_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Ccl_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Ccl_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Ccl_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Ccl
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CclPadraoLabel_Click()

Dim objCcl As New ClassCcl
Dim colSelecao As New Collection

    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCclPadrao)

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO
    iCodigoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComboFilialPedido_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub ComboFilialPedido_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Data_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)

End Sub

Private Sub DataFimPadrao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFimPadrao, iAlterado)

End Sub

Private Sub DataInicioPadrao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicioPadrao, iAlterado)

End Sub

Private Sub DataPrevisaoFim_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataPrevisaoFim_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub DataPrevisaoFim_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub DataPrevisaoFim_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = DataPrevisaoFim
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DataPrevisaoInicio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataPrevisaoInicio_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub DataPrevisaoInicio_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub DataPrevisaoInicio_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGrid.objControle = DataPrevisaoInicio
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Destinacao_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Destinacao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Destinacao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Destinacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Destinacao
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub GeraReqCompra_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub GridMovimentos_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer
Dim iLinhaAnterior As Integer
Dim iNum As Integer
Dim lErro As Long
Dim iLinhasExistentes As Integer 'm

On Error GoTo Erro_GridMovimentos_KeyDown

    'Define se OP é nova ou existente
    If gcolItemOP.Count >= GridMovimentos.Row Then
        iNum = gcolItemOP.Item(GridMovimentos.Row)
    Else
        iNum = 0
    End If

    'Se for uma nova OP
    If iNum = 0 Then

        'Guarda iLinhasExistentes
        iLinhasExistentesAnterior = objGrid.iLinhasExistentes

        'Verifica se a Tecla apertada foi Del
        If KeyCode = vbKeyDelete Then

            'Guarda o índice da Linha a ser Excluída
            iLinhaAnterior = GridMovimentos.Row

        End If

        Call Grid_Trata_Tecla1(KeyCode, objGrid)

        'Verifica se a Linha foi realmente excluída
        If objGrid.iLinhasExistentes < iLinhasExistentesAnterior Then

            gcolItemOP.Remove (iLinhaAnterior)
            gobjOP.colItens.Remove iLinhaAnterior  'm

            For iLinhasExistentes = 1 To objGrid.iLinhasExistentes 'm
                If gobjOP.colItens(iLinhasExistentes).iPossuiGrade = MARCADO Then
                    GridMovimentos.TextMatrix(iLinhasExistentes, 0) = "# " & iLinhasExistentes
                Else
                    GridMovimentos.TextMatrix(iLinhasExistentes, 0) = iLinhasExistentes
                End If
                
            Next

            GridMovimentos.TextMatrix(iLinhasExistentes, 0) = iLinhasExistentes

            lErro = QuantDisponivel_Calcula(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col))
            If lErro <> SUCESSO Then gError 55328

        End If

    End If

    Exit Sub
    
Erro_GridMovimentos_KeyDown:

    Select Case gErr
    
        Case 55328
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163718)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Maquina_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoCclPadrao_evSelecao(obj1 As Object)
'Preenche CclPadrao

Dim objCcl As New ClassCcl
Dim sCclFormatada As String
Dim sCclMascarado As String
Dim lErro As Long

On Error GoTo Erro_objEventoCclPadrao_evSelecao

    Set objCcl = obj1

    sCclMascarado = String(STRING_CCL, 0)

    lErro = Mascara_RetornaCclEnxuta(objCcl.sCCL, sCclMascarado)
    If lErro <> SUCESSO Then gError 22930

    CclPadrao.PromptInclude = False
    CclPadrao.Text = sCclMascarado
    CclPadrao.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoCclPadrao_evSelecao:

    Select Case gErr

        Case 22930
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", gErr, objCcl.sCCL)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163719)

    End Select

    Exit Sub

End Sub

Private Sub objEventoAlmoxPadrao_evSelecao(obj1 As Object)

Dim objAlmoxarifado As ClassAlmoxarifado

    Set objAlmoxarifado = obj1

    'Preenche AlmoxPadrao
    AlmoxPadrao.Text = objAlmoxarifado.sNomeReduzido

    Me.Show

End Sub

Private Sub BotaoCcls_Click()
'chama tela de Lista de Ccl

Dim lErro As Long
Dim objCcls As New ClassCcl
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoCcls_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridMovimentos.Row = 0 Then gError 43742

    'Verifica se o Produto está preenchido
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))) = 0 Then gError 43743

    Call Chama_Tela("CclLista", colSelecao, objCcls, objEventoCcl)
    
    Exit Sub
    
Erro_BotaoCcls_Click:

    Select Case gErr
    
        Case 43742
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 43743
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163720)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoEstoque_Click()
'Informa se produto é estocado em algum almoxarifado

Dim lErro As Long
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoEstoque_Click

    If GridMovimentos.Row = 0 Then gError 43719

    sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 21930

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        colSelecao.Add sProdutoFormatado
        'chama a tela de lista de estoque do produto corrente
        Call Chama_Tela("EstoqueProdutoFilialLista", colSelecao, objEstoqueProduto, objEventoEstoque)
    Else
        Error 43739
    End If

    Exit Sub

Erro_BotaoEstoque_Click:

    Select Case gErr

        Case 21930
        
        Case 43719
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 43739
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163721)

    End Select

    Exit Sub

End Sub

Private Sub objEventoEstoque_evselecao(obj1 As Object)

Dim objEstoqueProduto As New ClassEstoqueProduto
Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sCodProduto As String

On Error GoTo Erro_objEventoEstoque_evselecao

    If GridMovimentos.Row <> 0 Then

        Set objEstoqueProduto = obj1

        sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)

        lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 22941

        'Verifica se o produto está preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            'Preenche o Nome do Almoxarifado
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col) = objEstoqueProduto.sAlmoxarifadoNomeReduzido

            Almoxarifado.Text = objEstoqueProduto.sAlmoxarifadoNomeReduzido

            'Calcula a Quantidade Disponível nesse Almoxarifado
            lErro = QuantDisponivel_Calcula(sCodProduto, objEstoqueProduto.sAlmoxarifadoNomeReduzido)
            If lErro <> SUCESSO Then gError 41304

        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoEstoque_evselecao:

    Select Case gErr

        Case 22941, 41304

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163722)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()
'ativa a exclusão de uma OP

Dim lErro As Long
Dim objOrdemDeProducao As New ClassOrdemDeProducao
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'verifica se código está preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 21936

    objOrdemDeProducao.sCodigo = Codigo.Text
    objOrdemDeProducao.iFilialEmpresa = giFilialEmpresa

    'Pede ao usuário que confire a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_OP", objOrdemDeProducao.sCodigo, objOrdemDeProducao.iFilialEmpresa)
    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If
    
    'exclui a OP
    lErro = CF("OrdemDeProducao_Exclui", objOrdemDeProducao)
    If lErro <> SUCESSO And lErro <> 21936 Then gError 21940

    'se OP não existir -> erro
    If lErro = 21936 Then gError 21950

    'Limpa a tela
    lErro = Limpa_Tela_OrdemDeProducao
    If lErro <> SUCESSO Then gError 21951

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 21936
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 21940, 21951

        Case 21950
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_INEXISTENTE", gErr, objOrdemDeProducao.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163723)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()
'implementa gravação de uma nova ou atualizacao de uma OP

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Rotina de gravação da OP
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 21750

    'limpa a tela
    lErro = Limpa_Tela_OrdemDeProducao
    If lErro <> SUCESSO Then gError 21951
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 21750, 21951

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163724)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'testa se houva alguma alteração
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 21952

    'limpa a tela
    lErro = Limpa_Tela_OrdemDeProducao
    If lErro <> SUCESSO Then gError 21953

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 21952, 21953

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163725)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProdutos_Click()

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim colSelecao As Collection

On Error GoTo Erro_BotaoProdutos_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridMovimentos.Row = 0 Then gError 43718

    'Verifica se o Produto está preenchido
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))) > 0 Then
    
        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 55325
        
        If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""
        
    End If

    objProduto.sCodigo = sProduto

    'Lista de produtos produzíveis
    Call Chama_Tela("ProdutoProduzivelLista", colSelecao, objProduto, objEventoProduto)
        
    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr
    
        Case 43718
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 55325
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163726)
    
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoMaquina_evSelecao(obj1 As Object)

Dim objMaquinas As ClassMaquinas
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_objEventoMaquina_evSelecao

    Set objMaquinas = obj1
    
    sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 106426

    'verifica se o produto esta preenchido...
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        Maquina.Text = objMaquinas.sNomeReduzido
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Maquina_Col) = objMaquinas.sNomeReduzido
    
    End If
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoMaquina_evSelecao:

    Select Case gErr
        
        Case 106426
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163727)
    
    End Select

End Sub

Private Sub objEventoPedidoDeVenda_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objItemPedido As ClassItemPedido
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objCliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente
Dim objPedidoDeVenda As New ClassPedidoDeVenda
Dim iIndice As Integer

On Error GoTo Erro_objEventoPedidoDeVenda_evSelecao

    If GridMovimentos.Row <> 0 Then

        Set objItemPedido = obj1
        
        sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)

        lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 52020

        'verifica se o produto esta preenchido e se a destinacao é pedido de Venda
        If iProdutoPreenchido = PRODUTO_PREENCHIDO And GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Destinacao_Col) = STRING_PEDIDOVENDA Then
                               
            objPedidoDeVenda.lCodigo = objItemPedido.lCodPedido
            objPedidoDeVenda.iFilialEmpresa = objItemPedido.iFilialEmpresa
        
            lErro = CF("PedidoDeVenda_Le", objPedidoDeVenda)
            If lErro <> SUCESSO And lErro <> 26509 Then gError 52051
        
            If lErro = 26509 Then gError 52052
            
            For iIndice = 0 To ComboFilialPedido.ListCount - 1
                If objItemPedido.iFilialEmpresa = ComboFilialPedido.ItemData(iIndice) Then
                    ComboFilialPedido.ListIndex = iIndice
                    Exit For
                End If
            Next
            
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_PedidoDeVenda_Col) = CStr(objItemPedido.lCodPedido)
            
            PedidoDeVendaId.Text = CStr(objItemPedido.lCodPedido)
            
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialPedido_Col) = ComboFilialPedido.Text
            
            objCliente.lCodigo = objPedidoDeVenda.lCliente
            
            'le o nome reduzido do cliente
            lErro = CF("Cliente_Le", objCliente)
            If lErro <> SUCESSO And lErro <> 12293 Then gError 52021
        
            If lErro <> SUCESSO Then gError 52022
                    
            'preenche com o nome reduzido do cliente
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Cliente_Col) = objCliente.sNomeReduzido
                    
            objFilialCliente.lCodCliente = objCliente.lCodigo
            objFilialCliente.iCodFilial = objPedidoDeVenda.iFilial
        
            'le o nome da filial do cliente
            lErro = CF("FilialCliente_Le", objFilialCliente)
            If lErro <> SUCESSO And lErro <> 12567 Then gError 52023
        
            If lErro = 12567 Then gError 52024
        
            'preenche CODIGO - NOME
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialCliente_Col) = objFilialCliente.iCodFilial & SEPARADOR & objFilialCliente.sNome
        
        End If
        
    End If
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoPedidoDeVenda_evSelecao:
    
    Select Case gErr
    
        Case 52020, 52021, 52023, 52051
        
        Case 52022
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)
    
        Case 52024
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_SEM_FILIAL", gErr, objFilialCliente.lCodCliente)
        
        Case 52052
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_NAO_CADASTRADA", gErr, objPedidoDeVenda.lCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163728)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String
Dim objItemOP As New ClassItemOP

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    If GridMovimentos.Row <> 0 Then

        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 22935

        'Se o produto não estiver preenchido
        If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then

            'Lê o produto no BD para obter UM de estoque
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 126585

            If lErro = 28030 Then gError 126586

            sProdutoMascarado = String(STRING_PRODUTO, 0)


            'mascara produto escolhido
            lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 22936

            'verifica sua existência na OP
            lErro = VerificaUso_Produto(objProduto)
            If lErro <> SUCESSO And lErro <> 41316 Then gError 55327

            If lErro = 41316 Then gError 22958

'            Call Carrega_ComboVersoes(objProduto.sCodigo)
            
            'Lê os demais atributos do Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 22937

            If lErro = 28030 Then gError 22939

            If objProduto.iPCP = PRODUTO_PCP_NAOPODE Or objProduto.iCompras <> PRODUTO_PRODUZIVEL Then gError 55276

            Produto.PromptInclude = False
            Produto.Text = sProdutoMascarado
            Produto.PromptInclude = True
            
            

            If Not (Me.ActiveControl Is Produto) Then

                'preenche produto
                GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col) = sProdutoMascarado
    
                'Preenche a Linha do Grid
                lErro = ProdutoLinha_Preenche(objProduto, objItemOP)
                If lErro <> SUCESSO Then gError 22938
    
                'calcula a qtd disponível
                lErro = QuantDisponivel_Calcula(Produto.Text, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col), objProduto)
                If lErro <> SUCESSO Then gError 41305

            End If

        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 22935, 22937, 22938, 41305, 55327, 126585

        Case 22936
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objProduto.sCodigo)
            
        Case 22939
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 22958
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_DUPLICADO", gErr, sProdutoMascarado, Codigo.Text)
            
        Case 55276
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL1", gErr, sProdutoMascarado)

        Case 126586
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163729)

    End Select

    Exit Sub

End Sub

Private Sub CodigoOPLabel_Click()

Dim objOrdemDeProducao As New ClassOrdemDeProducao
Dim colSelecao As New Collection
Dim sSelecao As String

    'preenche o objOrdemDeProducao com o código da tela , se estiver preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then objOrdemDeProducao.sCodigo = Codigo.Text
    
    sSelecao = "Tipo = 1"
    
    'lista as OP's
    Call Chama_Tela("OrdemProducaoLista", colSelecao, objOrdemDeProducao, objEventoCodigo, sSelecao)

End Sub

Private Sub ComboFilialPedido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComboFilialPedido_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = ComboFilialPedido
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function VerificaQuantidade_ItemPedido(objItemPedido As ClassItemPedido) As Long

Dim lErro As Long
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_VerificaQuantidade_ItemPedido

        lErro = CF("ItemPedido_Le", objItemPedido)
        If lErro <> SUCESSO And lErro <> 23971 Then gError 41308

        If lErro = 23971 Then gError 41309

        'avisa que qtd ordenada é diferente da qtd do item de pedido
        If Len(Quantidade.Text) > 0 Then
        
            If Trim(Quantidade.Text) <> CStr(objItemPedido.dQuantidade) Then
            
                vbMsg = Rotina_Aviso(vbOKOnly, "AVISO_QUANTIDADE_ITEMPEDIDO")
                
            End If
            
        End If

    VerificaQuantidade_ItemPedido = SUCESSO

    Exit Function

Erro_VerificaQuantidade_ItemPedido:

    VerificaQuantidade_ItemPedido = gErr

    Select Case gErr

        Case 41308

        Case 41309
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDIDO_INEXISTENTE1", gErr, objItemPedido.lCodPedido, objItemPedido.iFilialEmpresa, objItemPedido.sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163730)

    End Select

    Exit Function

End Function

Public Sub Form_Activate()

    'Carrega índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub CargaCombo_Situacao(objSituacao As Object)
'Carga dos itens da combo Situação

    objSituacao.AddItem STRING_NORMAL
    objSituacao.ItemData(objSituacao.NewIndex) = ITEMOP_SITUACAO_NORMAL
    objSituacao.AddItem STRING_DESABILITADA
    objSituacao.ItemData(objSituacao.NewIndex) = ITEMOP_SITUACAO_DESAB
    objSituacao.AddItem STRING_SACRAMENTADA
    objSituacao.ItemData(objSituacao.NewIndex) = ITEMOP_SITUACAO_SACR
    objSituacao.AddItem STRING_BAIXADA
    objSituacao.ItemData(objSituacao.NewIndex) = ITEMOP_SITUACAO_BAIXADA

End Sub

Public Sub CargaCombo_Destinacao(objDestinacao As Object)
'Carga dos itens da combo Destinação

    objDestinacao.AddItem STRING_ESTOQUE
    objDestinacao.ItemData(objDestinacao.NewIndex) = ITEMOP_DESTINACAO_ESTOQUE
    objDestinacao.AddItem STRING_PEDIDOVENDA
    objDestinacao.ItemData(objDestinacao.NewIndex) = ITEMOP_DESTINACAO_PV
    objDestinacao.AddItem STRING_CONSUMO
    objDestinacao.ItemData(objDestinacao.NewIndex) = ITEMOP_DESTINACAO_CONSUMO

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim sItem As String
Dim sMascaraCclPadrao As String
Dim objFiliais As AdmFiliais
Dim colModulo As New AdmColModulo
Dim tModulo As typeModulo
Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    iFrameAtual = 1

    Set gcolItemOP = New Collection

    Set objEventoCodigo = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoCcl = New AdmEvento
    Set objEventoCclPadrao = New AdmEvento
    Set objEventoAlmoxPadrao = New AdmEvento
    Set objEventoEstoque = New AdmEvento
    Set objEventoPedidoDeVenda = New AdmEvento
    Set objEventoMaquina = New AdmEvento
    Set objEventoPrestServ = New AdmEvento
    
    For Each objFiliais In gcolFiliais
        If objFiliais.iCodFilial <> 0 Then
            sItem = CStr(objFiliais.iCodFilial) & SEPARADOR & CStr(objFiliais.sNome)
            ComboFilialPedido.AddItem sItem
            ComboFilialPedido.ItemData(ComboFilialPedido.NewIndex) = objFiliais.iCodFilial
        End If
    Next

    'Carrega Ítens das Combos
    Call CargaCombo_Situacao(Situacao)
    Call CargaCombo_Destinacao(DestinacaoPadrao)
    Call CargaCombo_Destinacao(Destinacao)
    

    'Inicializa máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 22963

    'Inicializa Máscara de CclPadrao e Ccl
    sMascaraCclPadrao = String(STRING_CCL, 0)

    lErro = MascaraCcl(sMascaraCclPadrao)
    If lErro <> SUCESSO Then gError 22964

    CclPadrao.Mask = sMascaraCclPadrao
    Ccl.Mask = sMascaraCclPadrao

    Quantidade.Format = FORMATO_ESTOQUE

    'Preenche a combo Destinação com o padrão (Estoque)
    For iIndice = 0 To DestinacaoPadrao.ListCount - 1
        If DestinacaoPadrao.ItemData(iIndice) = ITEMOP_DESTINACAO_ESTOQUE Then
            DestinacaoPadrao.ListIndex = iIndice
            Exit For
        End If
    Next
    
    'Coloca a Data Atual na Tela
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True

    'Coloca a Data atual nas Datas de Previsão
    DataInicioPadrao.PromptInclude = False
    DataInicioPadrao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataInicioPadrao.PromptInclude = True

    DataFimPadrao.PromptInclude = False
    DataFimPadrao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataFimPadrao.PromptInclude = True

    'inicializa Grid
    lErro = Inicializa_GridMovimentos
    If lErro <> SUCESSO Then gError 21974
        
    Set gobjOP = New ClassOrdemDeProducao
        
    iAlterado = 0
    iCodigoAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 21974, 22963, 22964

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163731)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objEventoCodigo = Nothing
    Set objEventoProduto = Nothing
    Set objEventoCcl = Nothing
    Set objEventoCclPadrao = Nothing
    Set objEventoAlmoxPadrao = Nothing
    Set objEventoEstoque = Nothing
    Set objEventoPedidoDeVenda = Nothing
    Set objEventoMaquina = Nothing
    Set objEventoPrestServ = Nothing

    Set gcolItemOP = Nothing
    Set gobjOP = Nothing
    
    Set objGrid = Nothing

   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
   
    If lErro <> SUCESSO Then gError 21976

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case 21976

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163732)

    End Select

    Exit Sub

End Sub

Private Sub GridMovimentos_Click()

Dim iExecutaEntradaCelula As Integer

        Call Grid_Click(objGrid, iExecutaEntradaCelula)

        If iExecutaEntradaCelula = 1 Then
            Call Grid_Entrada_Celula(objGrid, iAlterado)
        End If

End Sub

Private Sub GridMovimentos_GotFocus()
    Call Grid_Recebe_Foco(objGrid)
End Sub

Private Sub GridMovimentos_EnterCell()

    Call Grid_Entrada_Celula(objGrid, iAlterado)

End Sub

Private Sub GridMovimentos_LeaveCell()
    Call Saida_Celula(objGrid)
End Sub

Private Sub GridMovimentos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Private Sub GridMovimentos_RowColChange()

Dim lErro As Long

On Error GoTo Erro_GridMovimentos_RowColChange

    Call Grid_RowColChange(objGrid)

    If (GridMovimentos.Row <> iLinhaAntiga) Then

        lErro = QuantDisponivel_Calcula(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col))
        If lErro <> SUCESSO Then gError 41310

        'Guarda a Linha corrente
        iLinhaAntiga = GridMovimentos.Row

    End If

    Exit Sub

Erro_GridMovimentos_RowColChange:

    Select Case gErr

        Case 41310

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163733)

    End Select

    Exit Sub

End Sub

Private Sub GridMovimentos_Scroll()

    Call Grid_Scroll(objGrid)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        Select Case GridMovimentos.Col

            Case iGrid_Produto_Col

                lErro = Saida_Celula_Produto(objGridInt)
                If lErro <> SUCESSO Then gError 21977
                
            Case iGrid_Versao_Col

                lErro = Saida_Celula_Versao(objGridInt)
                If lErro <> SUCESSO Then gError 106337
            
            Case iGrid_Maquina_Col

                lErro = Saida_Celula_Maquina(objGridInt)
                If lErro <> SUCESSO Then gError 106361
            
            Case iGrid_Quantidade_Col

                lErro = Saida_Celula_Quantidade(objGridInt)
                If lErro <> SUCESSO Then gError 21979

            Case iGrid_MetragemCons_Col

                lErro = Saida_Celula_MetragemCons(objGridInt)
                If lErro <> SUCESSO Then gError 117650

            Case iGrid_Enfesto_Col

                lErro = Saida_Celula_Enfesto(objGridInt)
                If lErro <> SUCESSO Then gError 117651

            Case iGrid_Risco_Col

                lErro = Saida_Celula_Risco(objGridInt)
                If lErro <> SUCESSO Then gError 117652

            Case iGrid_Almoxarifado_Col

                lErro = Saida_Celula_Almoxarifado(objGridInt)
                If lErro <> SUCESSO Then gError 21980
            

            Case iGrid_Benef_Col
                lErro = Saida_Celula_Benef(objGridInt)
                If lErro <> SUCESSO Then gError 91272
            
            Case iGrid_Ccl_Col

                lErro = Saida_Celula_Ccl(objGridInt)
                If lErro <> SUCESSO Then gError 21981

            Case iGrid_DataPrevInicio_Col

                lErro = Saida_Celula_DataPrevInicio(objGridInt)
                If lErro <> SUCESSO Then gError 21982

            Case iGrid_DataPrevFim_Col

                lErro = Saida_Celula_DataPrevFim(objGridInt)
                If lErro <> SUCESSO Then gError 21983

            Case iGrid_Destinacao_Col

                lErro = Saida_Celula_Destinacao(objGridInt)
                If lErro <> SUCESSO Then gError 21984

            Case iGrid_FilialPedido_Col

                lErro = Saida_Celula_FilialPedido(objGridInt)
                If lErro <> SUCESSO Then gError 41311

            Case iGrid_Situacao_Col

                lErro = Saida_Celula_Situacao(objGridInt)
                If lErro <> SUCESSO Then gError 41312

            Case iGrid_PedidoDeVenda_Col

                lErro = Saida_Celula_PedidoDeVenda(objGridInt)
                If lErro <> SUCESSO Then gError 41313

            Case iGrid_Prioridade_Col

                lErro = Saida_Celula_Prioridade(objGridInt)
                If lErro <> SUCESSO Then gError 41315

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 21986

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 21977, 21978, 21979, 21980, 21981, 21982, 21983, 21984, 41311 To 41315, 91272, 106337, 106361, 117650 To 117652

        Case 21986
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163734)

    End Select

    Exit Function

End Function

Private Function VerificaUso_Produto(ByVal objProduto As ClassProduto) As Long
'Verifica se existem produtos repetidos na OP

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sProdutoPai As String
Dim sProdutoFormatadoPai As String

On Error GoTo Erro_VerificaUso_Produto

    If objGrid.iLinhasExistentes > 0 Then

        For iIndice = 1 To objGrid.iLinhasExistentes

            If GridMovimentos.Row <> iIndice Then

                lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
                If lErro <> SUCESSO Then gError 22947

                If sProdutoFormatado = objProduto.sCodigo Then gError 41316
                
            End If

        Next

        'Se existir um produto pai de grade no grid ==> verificar se o produto em questao é filho deste pai de grade, se for ==> erro
        If Grid_Possui_Grade Then
            
            'Busca, caso exista, o produto pai de grade do prod em questão
            lErro = CF("Produto_Le_PaiGrade", objProduto, sProdutoPai)
            If lErro <> SUCESSO Then gError 126493
            
            'Se o produto tem um pai de grade
            If Len(Trim(sProdutoPai)) > 0 Then
                
                'Verifica se seu pai aparece no grid
                For iIndice = 1 To objGrid.iLinhasExistentes
                    
                    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatadoPai, iProdutoPreenchido)
                    If lErro <> SUCESSO Then gError 126492
                    
                    'Se aparecer ==> erro
                    If sProdutoFormatadoPai = sProdutoPai Then gError 126494
                
                Next
            
            End If
            
        End If

    End If

    VerificaUso_Produto = SUCESSO

    Exit Function

Erro_VerificaUso_Produto:

    VerificaUso_Produto = gErr

    Select Case gErr

        Case 22947, 41316, 126492, 126493

        Case 126494
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_PAI_GRADE_GRID", gErr, Trim(sProdutoPai), objProduto.sCodigo)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163735)

    End Select

End Function

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long
'faz a critica da celula de proddduto do grid que está deixando de ser a corrente

Dim lErro As Long
Dim sProduto As String
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult
Dim objMaquina As New ClassMaquinas
Dim objKit As New ClassKit
Dim objProdutoFilial As New ClassProdutoFilial
Dim iPossuiGrade As Integer
Dim iIndice As Integer
Dim colItensRomaneioGrade As New Collection
Dim objItensRomaneio As ClassItemRomaneioGrade
Dim objItemOP As New ClassItemOP
Dim objRomaneioGrade As New ClassRomaneioGrade

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 106433
    
    'se o produto foi preenchido
    If Len(Trim(Produto.ClipText)) <> 0 Then
        
        lErro = CF("Produto_Critica2", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then gError 126489
        
        'se produto estiver preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                
            'se é um produto gerencial e não é pai de grade ==> erro
            If lErro = 25043 And Len(Trim(objProduto.sGrade)) = 0 Then gError 126490
            
            'se o produto nao for gerencial e ainda assim deu erro ==> nao está cadastrado
            If lErro <> SUCESSO And lErro <> 25043 Then gError 21988
    
            'verifica se este produto já foi usado na OP
            lErro = VerificaUso_Produto(objProduto)
            If lErro <> SUCESSO And lErro <> 41316 Then gError 55326

            If lErro = 41316 Then gError 41317

            iPossuiGrade = DESMARCADO
    
            If Len(Trim(objProduto.sGrade)) > 0 Then iPossuiGrade = MARCADO
    
            If iPossuiGrade = DESMARCADO Then
    
                'se o produto nao controla estoque ==> erro
                If objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then gError 126491
        
                If objProduto.iPCP = PRODUTO_PCP_NAOPODE Or objProduto.iCompras <> PRODUTO_PRODUZIVEL Then gError 22946
    
                 'Preenche a linha do grid
                lErro = ProdutoLinha_Preenche(objProduto, objItemOP)
                If lErro <> SUCESSO Then gError 22945
    
                'Calcula a Quantidade Disponível
                lErro = QuantDisponivel_Calcula(Produto.Text, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col), objProduto)
                If lErro <> SUCESSO Then gError 41318
    
                'Verifica se é um kit
                objKit.sProdutoRaiz = sProdutoFormatado
                lErro = CF("Kit_Le_Padrao", objKit)
                If lErro <> SUCESSO And lErro <> 106304 Then gError 106430
                
                'Se encontrou => É UM KIT => Carrega a Combo com as Versoes
                If lErro <> 106304 Then Call Carrega_ComboVersoes(objProduto.sCodigo)
                
                If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))) = 0 Then
                    
                    objProdutoFilial.sProduto = sProdutoFormatado
                    objProdutoFilial.iFilialEmpresa = giFilialEmpresa
                    
                    'Busca o Lote Economico do Produto/FilialEmpresa
                    lErro = CF("ProdutoFilial_Le", objProdutoFilial)
                    If lErro <> SUCESSO And lErro <> 28261 Then gError 106402
                    
                    'preenche com o lote econômico (caso exista)
                    If objProdutoFilial.dLoteEconomico > 0 Then GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col) = objProdutoFilial.dLoteEconomico
                    
                End If
        
            'se é um produto pai de grade
            Else
            
                'Verifica se há filhos válidos da grade pai
                lErro = CF("Produto_Le_Filhos_Grade", objProduto, colItensRomaneioGrade)
                If lErro <> SUCESSO Then gError 126495
                
                'Se nao existir, erro
                If colItensRomaneioGrade.Count = 0 Then gError 126496
                
                'Para cada filho de grade do produto
                For Each objItensRomaneio In colItensRomaneioGrade
                    
                    'Verifica se ele já aparece no grid
                    For iIndice = 1 To objGridInt.iLinhasExistentes
                        
                        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
                        If lErro <> SUCESSO Then gError 126498
                        
                        'Se aparecer ==> Erro
                        If sProdutoFormatado = objItensRomaneio.sProduto Then gError 126497
                        
                    Next
                    
                Next
        
                objItemOP.sProduto = objProduto.sCodigo
                objItemOP.sSiglaUMEstoque = objProduto.sSiglaUMEstoque
                objItemOP.sCodigo = Codigo.Text
                objItemOP.iItem = GridMovimentos.Row
                objItemOP.sDescricao = objProduto.sDescricao
                objItemOP.sAlmoxarifadoNomeRed = AlmoxPadrao.Text
                        
                Set objRomaneioGrade = New ClassRomaneioGrade
                Set objRomaneioGrade.objTela = Me
                
                objRomaneioGrade.sNomeTela = Me.Name
                
                Set objRomaneioGrade.objObjetoTela = objItemOP
                            
                Call Chama_Tela_Modal("RomaneioGrade", objRomaneioGrade)
                If giRetornoTela <> vbOK Then gError 126499
        
                 'Preenche a linha do grid
                lErro = ProdutoLinha_Preenche(objProduto, objItemOP)
                If lErro <> SUCESSO Then gError 126552
        
                Call Atualiza_Grid_Movimentos(objItemOP)
            
            End If
        
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 21985

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 21987, 22945, 41318, 21985, 55326, 106430, 106433, 106402, 126495, 126498, 126499, 126552
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 21988
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)

            If vbMsg = vbYes Then
            
                objProduto.sCodigo = Produto.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 22946
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL1", gErr, Produto.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 41317
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_DUPLICADO", gErr, Produto.Text, Codigo.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 126490
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)
        
        Case 126491
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ESTOQUE", gErr, objProduto.sCodigo)
        
        Case 126496
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_PAI_GRADE_SEM_FILHOS", gErr, Produto.Text)
        
        Case 126497
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_FILHO_GRADE_GRID", gErr, Trim(objProduto.sCodigo), GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col))

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163736)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantTotal As Double
Dim objProdutoFilial As New ClassProdutoFilial
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade
    
    'se a quantidade foi preenchida
    If Len(Quantidade.ClipText) > 0 Then

        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 41319
    
        '******* Dia 30/10/2002 Sergio: Alteração necessária para verificar se a quantidade solicitada para a produção(Produto), é maior que a Quantidade mínima que pode ser produzida *********************
        
        'Coloca o Produto no Formato do Banco de Dados
        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 111429
        
        objProdutoFilial.sProduto = sProdutoFormatado
        objProdutoFilial.iFilialEmpresa = giFilialEmpresa
        
        'Busca o Lote Mínino do Produto/FilialEmpresa
        lErro = CF("ProdutoFilial_Le", objProdutoFilial)
        If lErro <> SUCESSO And lErro <> 28261 Then gError 111430
        
        'preenche Verifica se Lote mínimo esta preenchdo (caso exista)
        If objProdutoFilial.dLoteMinimo > 0 Then
        
            If StrParaDbl(Quantidade.Text) < objProdutoFilial.dLoteMinimo Then gError 111432
        
        End If
                
        '**************** Sergio  *********************
    
        Quantidade.Text = Formata_Estoque(Quantidade.Text)
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 21994

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 21994
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 41319
            Quantidade.SetFocus
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 111429, 111430
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 111432
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QDTPRODUTO_MENOR_LOTEMININO", gErr, objProdutoFilial.dLoteMinimo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163737)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MetragemCons(objGridInt As AdmGrid) As Long
'faz a critica da celula de metragemCons do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantTotal As Double
Dim objProdutoFilial As New ClassProdutoFilial
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Saida_Celula_MetragemCons

    Set objGridInt.objControle = MetragemCons
    
    'se a quantidade foi preenchida
    If Len(MetragemCons.ClipText) > 0 Then

        lErro = Valor_Positivo_Critica(MetragemCons.Text)
        If lErro <> SUCESSO Then gError 117653
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 117654

    Saida_Celula_MetragemCons = SUCESSO

    Exit Function

Erro_Saida_Celula_MetragemCons:

    Saida_Celula_MetragemCons = gErr

    Select Case gErr

        Case 117653, 117654
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163738)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Enfesto(objGridInt As AdmGrid) As Long
'faz a critica da celula de Enfesto do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantTotal As Double
Dim objProdutoFilial As New ClassProdutoFilial
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Saida_Celula_Enfesto

    Set objGridInt.objControle = Enfesto
    
    'se a quantidade foi preenchida
    If Len(Enfesto.ClipText) > 0 Then

        lErro = Valor_Positivo_Critica(Enfesto.Text)
        If lErro <> SUCESSO Then gError 117655
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 117656

    Saida_Celula_Enfesto = SUCESSO

    Exit Function

Erro_Saida_Celula_Enfesto:

    Saida_Celula_Enfesto = gErr

    Select Case gErr

        Case 117655, 117656
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163739)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Risco(objGridInt As AdmGrid) As Long
'faz a critica da celula de Risco do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantTotal As Double
Dim objProdutoFilial As New ClassProdutoFilial
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Saida_Celula_Risco

    Set objGridInt.objControle = Risco
    
    'se a quantidade foi preenchida
    If Len(Risco.ClipText) > 0 Then

        lErro = Valor_Positivo_Critica(Risco.Text)
        If lErro <> SUCESSO Then gError 117657
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 117658

    Saida_Celula_Risco = SUCESSO

    Exit Function

Erro_Saida_Celula_Risco:

    Saida_Celula_Risco = gErr

    Select Case gErr

        Case 117657, 117658
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163740)

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

        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 22947

        'se produto estiver preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            'verifica almoxarifado
            lErro = TP_Almoxarifado_Filial_Produto_Grid(sProdutoFormatado, Almoxarifado, objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25157 And lErro <> 25162 Then gError 22948

            If lErro = 25157 Then gError 22949

            If lErro = 25162 Then gError 22950

            'calcula qtd disponivel
            lErro = QuantDisponivel_Calcula(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), Almoxarifado.Text)
            If lErro <> SUCESSO Then gError 41322

        End If

    Else

        'Limpa a Quantidade Disponível da Tela
        QuantDisponivel.Caption = ""

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 31501

    Saida_Celula_Almoxarifado = SUCESSO

    Exit Function

Erro_Saida_Celula_Almoxarifado:

    Saida_Celula_Almoxarifado = gErr

    Select Case gErr

        Case 22947, 22948, 31501, 41322
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 22949
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ALMOXARIFADO2", Almoxarifado.Text)

            If vbMsg = vbYes Then

                objAlmoxarifado.sNomeReduzido = Almoxarifado.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 22950
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_INEXISTENTE1", CInt(Almoxarifado.Text))

            If vbMsg = vbYes Then

                objAlmoxarifado.iCodigo = CInt(Almoxarifado.Text)

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163741)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Ccl(objGridInt As AdmGrid) As Long
'faz a critica da celula de produto do grid que está deixando de ser a corrente

Dim lErro As Long, sCclFormatada As String
Dim objCcl As New ClassCcl
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Ccl

    Set objGridInt.objControle = Ccl

    If Len(Ccl.ClipText) > 0 Then

        lErro = CF("Ccl_Critica", Ccl.Text, sCclFormatada, objCcl)
        If lErro <> SUCESSO And lErro <> 5703 Then gError 22951

        If lErro = 5703 Then gError 22952

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 31503

    Saida_Celula_Ccl = SUCESSO

    Exit Function

Erro_Saida_Celula_Ccl:

    Saida_Celula_Ccl = gErr

    Select Case gErr

        Case 22951, 31503
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 22952
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", Ccl.Text)
            
            If vbMsgRes = vbYes Then
            
                objCcl.sCCL = sCclFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("CclTela", objCcl)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163742)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataPrevInicio(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DataPrevInicio

    Set objGridInt.objControle = DataPrevisaoInicio

    'verifica se a data está preenchida
    If Len(Trim(DataPrevisaoInicio.ClipText)) > 0 Then

        'verifica se a data é válida
        lErro = Data_Critica(DataPrevisaoInicio.Text)
        If lErro <> SUCESSO Then gError 31504

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 31505

    Saida_Celula_DataPrevInicio = SUCESSO

    Exit Function

Erro_Saida_Celula_DataPrevInicio:

    Saida_Celula_DataPrevInicio = gErr

    Select Case gErr

        Case 31504, 31505
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163743)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataPrevFim(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DataPrevFim

    Set objGridInt.objControle = DataPrevisaoFim

    'verifica se a data está preenchida
    If Len(Trim(DataPrevisaoFim.ClipText)) > 0 Then

        'verifica se a data é válida
        lErro = Data_Critica(DataPrevisaoFim.Text)
        If lErro <> SUCESSO Then gError 31506

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 31507

    Saida_Celula_DataPrevFim = SUCESSO

    Exit Function

Erro_Saida_Celula_DataPrevFim:

    Saida_Celula_DataPrevFim = gErr

    Select Case gErr

        Case 31506, 31507
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163744)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Prioridade(objGridInt As AdmGrid) As Long
'faz a critica da celula de Prioridade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Prioridade

    Set objGridInt.objControle = Prioridade

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 41323

    Saida_Celula_Prioridade = SUCESSO

    Exit Function

Erro_Saida_Celula_Prioridade:

    Saida_Celula_Prioridade = gErr

    Select Case gErr

        Case 41323
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163745)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PedidoDeVenda(objGridInt As AdmGrid) As Long
'faz a critica da celula de PedidoDeVenda do grid que está deixando de ser a corrente

Dim lErro As Long, iIndice As Integer
Dim objItemPedido As New ClassItemPedido
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sProduto As String
Dim objCliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente
Dim objPedidoDeVenda As New ClassPedidoDeVenda

On Error GoTo Erro_Saida_Celula_PedidoDeVenda

    Set objGridInt.objControle = PedidoDeVendaId

    If Len(Trim(PedidoDeVendaId.Text)) > 0 Then
    
        If GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Destinacao_Col) <> STRING_PEDIDOVENDA Then gError 41645

        If Trim(PedidoDeVendaId.Text) <> Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_PedidoDeVenda_Col)) And Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialPedido_Col))) <> 0 Then

            For iIndice = 0 To ComboFilialPedido.ListCount - 1
                If ComboFilialPedido.List(iIndice) = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialPedido_Col) Then Exit For
            Next
            
            objItemPedido.iFilialEmpresa = ComboFilialPedido.ItemData(iIndice)
            objItemPedido.lCodPedido = CLng(PedidoDeVendaId.Text)

            sProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)

            lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 41430

            objItemPedido.sProduto = sProdutoFormatado

            'Verifica se existe um ítem com os dados passados objItemPedido
            lErro = VerificaQuantidade_ItemPedido(objItemPedido)
            If lErro <> SUCESSO And lErro <> 41309 Then gError 41433

            If lErro = 41309 Then gError 41434
                        
            If Trim(PedidoDeVendaId.Text) <> "" And ComboFilialPedido.ListIndex <> -1 Then
                        
                objPedidoDeVenda.lCodigo = CLng(PedidoDeVendaId.Text)
                objPedidoDeVenda.iFilialEmpresa = ComboFilialPedido.ItemData(ComboFilialPedido.ListIndex)
            
                lErro = CF("PedidoDeVenda_Le", objPedidoDeVenda)
                If lErro <> SUCESSO And lErro <> 26509 Then gError 52029
            
                If lErro = 26509 Then gError 52030
                        
                objCliente.lCodigo = objPedidoDeVenda.lCliente
                
                'le o nome reduzido do cliente
                lErro = CF("Cliente_Le", objCliente)
                If lErro <> SUCESSO And lErro <> 12293 Then gError 52025
            
                If lErro <> SUCESSO Then gError 52026
                        
                'preenche com o nome reduzido do cliente
                GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Cliente_Col) = objCliente.sNomeReduzido
                        
                objFilialCliente.lCodCliente = objCliente.lCodigo
                objFilialCliente.iCodFilial = objPedidoDeVenda.iFilial
            
                'le o nome da filial do cliente
                lErro = CF("FilialCliente_Le", objFilialCliente)
                If lErro <> SUCESSO And lErro <> 12567 Then gError 52027
            
                If lErro = 12567 Then gError 52028
            
                'preenche CODIGO - NOME
                GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialCliente_Col) = objFilialCliente.iCodFilial & SEPARADOR & objFilialCliente.sNome
            End If
            
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 41325

    Saida_Celula_PedidoDeVenda = SUCESSO

    Exit Function

Erro_Saida_Celula_PedidoDeVenda:

    Saida_Celula_PedidoDeVenda = gErr

    Select Case gErr
        
        Case 41325, 41430, 41433, 52025, 52027, 52029
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 41434
            PedidoDeVendaId.SetFocus
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 41645
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESTINACAO_DEPENDENTE", gErr)
            PedidoDeVendaId.SetFocus
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 52026
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 52028
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_SEM_FILIAL", gErr, objFilialCliente.lCodCliente)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 52030
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_NAO_CADASTRADA", gErr, objPedidoDeVenda.lCodigo)
            PedidoDeVendaId.SetFocus
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163746)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Situacao(objGridInt As AdmGrid) As Long
'faz a critica da celula de Situacao do grid que está deixando de ser a corrente
Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Situacao

    Set objGridInt.objControle = Situacao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 41326

    Saida_Celula_Situacao = SUCESSO

    Exit Function

Erro_Saida_Celula_Situacao:

    Saida_Celula_Situacao = gErr

    Select Case gErr

        Case 41326
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163747)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Benef(objGridInt As AdmGrid) As Long
'faz a critica da celula de Benef do grid que está deixando de ser a corrente


Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Benef

    Set objGridInt.objControle = Benef

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 91273
    
    Saida_Celula_Benef = SUCESSO

    Exit Function

Erro_Saida_Celula_Benef:

   Saida_Celula_Benef = gErr

    Select Case gErr

        Case 91273
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163748)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Destinacao(objGridInt As AdmGrid) As Long
'faz a critica da celula de Destinacao do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Destinacao

    Set objGridInt.objControle = Destinacao

    If Destinacao.ListIndex >= 0 Then
        'limpa celulas se destinacao for diferente de Pedido de Venda e desabilita
        If Destinacao.List(Destinacao.ListIndex) <> STRING_PEDIDOVENDA Then
        
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialPedido_Col) = ""
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_PedidoDeVenda_Col) = ""
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Cliente_Col) = ""
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialCliente_Col) = ""
        
        End If
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 31508

    Saida_Celula_Destinacao = SUCESSO

    Exit Function

Erro_Saida_Celula_Destinacao:

    Saida_Celula_Destinacao = gErr

    Select Case gErr

        Case 31508
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163749)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_FilialPedido(objGridInt As AdmGrid) As Long
'faz a critica da celula FilialPedido do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objItemPedido As New ClassItemPedido
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sProduto As String
Dim objCliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente
Dim objPedidoDeVenda As New ClassPedidoDeVenda

On Error GoTo Erro_Saida_Celula_FilialPedido

    Set objGridInt.objControle = ComboFilialPedido

    If ComboFilialPedido.ListIndex >= 0 Then
    
        If GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Destinacao_Col) <> STRING_PEDIDOVENDA Then gError 41644

        If ComboFilialPedido.List(ComboFilialPedido.ListIndex) <> GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialPedido_Col) And Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_PedidoDeVenda_Col))) <> 0 Then

            objItemPedido.iFilialEmpresa = ComboFilialPedido.ItemData(ComboFilialPedido.ListIndex)
            objItemPedido.lCodPedido = CLng(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_PedidoDeVenda_Col))

            sProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)
            lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 41431

            objItemPedido.sProduto = sProdutoFormatado

            'Verifica se existe um ítem com os dados passados objItemPedido
            lErro = VerificaQuantidade_ItemPedido(objItemPedido)
            If lErro <> SUCESSO And lErro <> 41309 Then gError 41432

            If lErro = 41309 Then gError 41435
                
            If Trim(PedidoDeVendaId.Text) <> "" And ComboFilialPedido.ListIndex <> -1 Then
                        
                objPedidoDeVenda.lCodigo = CLng(PedidoDeVendaId.Text)
                objPedidoDeVenda.iFilialEmpresa = ComboFilialPedido.ItemData(ComboFilialPedido.ListIndex)
            
                lErro = CF("PedidoDeVenda_Le", objPedidoDeVenda)
                If lErro <> SUCESSO And lErro <> 26509 Then gError 52035
            
                If lErro = 26509 Then gError 52036
                        
                objCliente.lCodigo = objPedidoDeVenda.lCliente
                
                'le o nome reduzido do cliente
                lErro = CF("Cliente_Le", objCliente)
                If lErro <> SUCESSO And lErro <> 12293 Then gError 52031
            
                If lErro <> SUCESSO Then gError 52032
                        
                'preenche com o nome reduzido do cliente
                GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Cliente_Col) = objCliente.sNomeReduzido
                        
                objFilialCliente.lCodCliente = objCliente.lCodigo
                objFilialCliente.iCodFilial = objPedidoDeVenda.iFilial
            
                'le o nome da filial do cliente
                lErro = CF("FilialCliente_Le", objFilialCliente)
                If lErro <> SUCESSO And lErro <> 12567 Then gError 52033
            
                If lErro = 12567 Then gError 52034
            
                'preenche CODIGO - NOME
                GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialCliente_Col) = objFilialCliente.iCodFilial & SEPARADOR & objFilialCliente.sNome
            End If
            
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 31510

    Saida_Celula_FilialPedido = SUCESSO

    Exit Function

Erro_Saida_Celula_FilialPedido:

    Saida_Celula_FilialPedido = gErr

    Select Case gErr

        Case 31510, 41431, 41432, 52031, 52033, 52035
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 41435
            ComboFilialPedido.ListIndex = -1
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 41644
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESTINACAO_DEPENDENTE", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 52032
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 52034
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_SEM_FILIAL", gErr, objFilialCliente.lCodCliente)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 52036
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_NAO_CADASTRADA", gErr, objPedidoDeVenda.lCodigo)
            PedidoDeVendaId.SetFocus
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163750)

    End Select

    Exit Function

End Function

Sub Limpa_GridMovimentos()

    Call Grid_Limpa(objGrid)

End Sub

Function Preenche_GridMovimentos(colItensOP As Collection) As Long
'preenche o grid com os dados contidos na coleção colItensOP

Dim lErro As Long, sCclMascarado As String
Dim iIndice As Integer, iIndice1 As Integer, sProdutoMascarado As String
Dim objItemOP As New ClassItemOP, objProduto As New ClassProduto
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objPedidoDeVenda As New ClassPedidoDeVenda
Dim objCliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente
Dim objMaquina As New ClassMaquinas

On Error GoTo Erro_Preenche_GridMovimentos

    'Remove os ítens de gcolItemOP
    Set gcolItemOP = New Collection

    iIndice = 1

    'preenche o grid com os dados retornados na coleção colItensOP
    For Each objItemOP In colItensOP

        '****** IF INCLUÍDO PARA TRATAMENTO DE GRADE ***************
        If objItemOP.iPossuiGrade = MARCADO Then GridMovimentos.TextMatrix(iIndice, 0) = "# " & GridMovimentos.TextMatrix(iIndice, 0)

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        'Mascara produto
        lErro = Mascara_RetornaProdutoTela(objItemOP.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 22927

        Produto.PromptInclude = False
        Produto.Text = sProdutoMascarado
        Produto.PromptInclude = True

        GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col) = sProdutoMascarado

        'le o produto para obter sua descricao
        objProduto.sCodigo = objItemOP.sProduto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO Then gError 22928

        GridMovimentos.TextMatrix(iIndice, iGrid_DescricaoItem_Col) = objProduto.sDescricao
        GridMovimentos.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItemOP.sSiglaUM
        GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemOP.dQuantidade)
        GridMovimentos.TextMatrix(iIndice, iGrid_MetragemCons_Col) = objItemOP.dMetragemCons
        GridMovimentos.TextMatrix(iIndice, iGrid_Enfesto_Col) = objItemOP.dEnfesto
        GridMovimentos.TextMatrix(iIndice, iGrid_Risco_Col) = objItemOP.dRisco

        If objItemOP.iPossuiGrade = DESMARCADO Then
        
            'Tenta ler almoxarifado
            objAlmoxarifado.iCodigo = objItemOP.iAlmoxarifado
    
            lErro = CF("Almoxarifado_Le", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 22235 Then gError 21970
    
            If lErro = 22235 Then gError 21971

            GridMovimentos.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido

        End If

        'Preenche Benef
        GridMovimentos.TextMatrix(iIndice, iGrid_Benef_Col) = objItemOP.iBeneficiamento
        
        'mascara Ccl , se estiver informada
        If objItemOP.sCCL <> "" Then

            sCclMascarado = String(STRING_CCL, 0)

            lErro = Mascara_MascararCcl(objItemOP.sCCL, sCclMascarado)
            If lErro <> SUCESSO Then gError 22929

        Else
            sCclMascarado = ""
        End If

        GridMovimentos.TextMatrix(iIndice, iGrid_Ccl_Col) = sCclMascarado

        'preenche datas
        If objItemOP.dtDataInicioProd <> DATA_NULA Then
            GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevInicio_Col) = Format(objItemOP.dtDataInicioProd, "dd/mm/yyyy")
        Else
            GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevInicio_Col) = ""
        End If

        If objItemOP.dtDataFimProd <> DATA_NULA Then
            GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevFim_Col) = Format(objItemOP.dtDataFimProd, "dd/mm/yyyy")
        Else
            GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevFim_Col) = ""
        End If

        'preenche Situação
        
        For iIndice1 = 0 To Situacao.ListCount - 1
            If Situacao.ItemData(iIndice1) = objItemOP.iSituacao Then
                Situacao.ListIndex = iIndice1
                Exit For
            End If
        Next
            
        GridMovimentos.TextMatrix(iIndice, iGrid_Situacao_Col) = Situacao.Text

        'preenche Destinação
        For iIndice1 = 0 To Destinacao.ListCount - 1
            If Destinacao.ItemData(iIndice1) = objItemOP.iDestinacao Then
                Destinacao.ListIndex = iIndice1
                Exit For
            End If
        Next

        GridMovimentos.TextMatrix(iIndice, iGrid_Destinacao_Col) = Destinacao.Text
        
        If GridMovimentos.TextMatrix(iIndice, iGrid_Destinacao_Col) = STRING_PEDIDOVENDA Then
        
            'le o pedido para pegar o Cliente e a filial do cliente
            objPedidoDeVenda.lCodigo = objItemOP.lCodPedido
            objPedidoDeVenda.iFilialEmpresa = objItemOP.iFilialEmpresa
            
            lErro = CF("PedidoDeVenda_Le", objPedidoDeVenda)
            If lErro <> SUCESSO And lErro <> 26509 Then gError 52081
            
            'se nao encontrou  - - - > Erro
            If lErro = 26509 Then
                'Verifica se o Pedido de Venda está baixado
                lErro = CF("PedidoVendaBaixado_Le", objPedidoDeVenda)
                If lErro <> SUCESSO And lErro <> 46135 Then gError 76085
                
                'se nao encontrou ==> Erro
                If lErro = 46135 Then gError 52082
            
            End If
            
            'preenche Filial do pedido
            For iIndice1 = 0 To ComboFilialPedido.ListCount - 1
                If ComboFilialPedido.ItemData(iIndice1) = objItemOP.iFilialPedido Then
                    ComboFilialPedido.ListIndex = iIndice1
                    Exit For
                End If
            Next
            
            GridMovimentos.TextMatrix(iIndice, iGrid_FilialPedido_Col) = ComboFilialPedido.Text
        
            'preenche Pedido de Venda e ItemPV
            GridMovimentos.TextMatrix(iIndice, iGrid_PedidoDeVenda_Col) = CStr(objItemOP.lCodPedido)
                            
            objCliente.lCodigo = objPedidoDeVenda.lCliente
                
            'le o nome reduzido do cliente
            lErro = CF("Cliente_Le", objCliente)
            If lErro <> SUCESSO And lErro <> 12293 Then gError 52083
            
            If lErro <> SUCESSO Then gError 52084
                        
            'preenche com o nome reduzido do cliente
            GridMovimentos.TextMatrix(iIndice, iGrid_Cliente_Col) = objCliente.sNomeReduzido
                        
            objFilialCliente.lCodCliente = objCliente.lCodigo
            objFilialCliente.iCodFilial = objPedidoDeVenda.iFilial
            
            'le o nome da filial do cliente
            lErro = CF("FilialCliente_Le", objFilialCliente)
            If lErro <> SUCESSO And lErro <> 12567 Then gError 52085
            
            If lErro = 12567 Then gError 52086
            
            'preenche CODIGO - NOME
            GridMovimentos.TextMatrix(iIndice, iGrid_FilialCliente_Col) = objFilialCliente.iCodFilial & SEPARADOR & objFilialCliente.sNome
            
        End If
        
        'preenche prioridade
        GridMovimentos.TextMatrix(iIndice, iGrid_Prioridade_Col) = CStr(objItemOP.iPrioridade)
        
        'Preenche com a Versão do Kit
        GridMovimentos.TextMatrix(iIndice, iGrid_Versao_Col) = objItemOP.sVersao
        
        If objItemOP.lNumIntEquipamento <> 0 Then
        
            objMaquina.lNumIntDoc = objItemOP.lNumIntEquipamento
            
            'Le a Máquina atraves do NumIntDoc
            lErro = CF("Maquinas_Le_NumIntDoc", objMaquina)
            If lErro <> SUCESSO And lErro <> 106353 Then gError 106355
            
            'Se nao encontrou => Erro
            If lErro = 106353 Then gError 106356
            
            'Preenche a máquina
            GridMovimentos.TextMatrix(iIndice, iGrid_Maquina_Col) = objMaquina.sNomeReduzido
        
        End If
        
        'adiciona item à coleção
        gcolItemOP.Add objItemOP.iItem
        gobjOP.colItens.Add objItemOP

        iIndice = iIndice + 1

    Next

    objGrid.iLinhasExistentes = colItensOP.Count
    
    Call Grid_Refresh_Checkbox(objGrid)
    
    Preenche_GridMovimentos = SUCESSO

    Exit Function

Erro_Preenche_GridMovimentos:

    Preenche_GridMovimentos = gErr

    Select Case gErr

        Case 21968, 21969, 21970, 22929, 52081, 52083, 52085, 76085

        Case 21971
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO", gErr, objAlmoxarifado.iCodigo)
        
        Case 22927
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objProduto.sCodigo)
        
        Case 22928
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 52082
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_NAO_CADASTRADA", gErr, objPedidoDeVenda.lCodigo)
        
        Case 52084
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)
    
        Case 52086
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_SEM_FILIAL", gErr, objFilialCliente.lCodCliente)
            
        Case 106355
        
        Case 106356
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_NAO_CADASTRADA", gErr, objMaquina.sNomeReduzido)
        
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163751)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objOrdemDeProducao As New ClassOrdemDeProducao
Dim sPedidoDeVenda As String
Dim vbMsg As VbMsgBoxResult
Dim sSituacao As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim bAchou As Boolean

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    bAchou = False
    
    If Len(Trim(Codigo.Text)) = 0 Then gError 31513

    'Verifica se a Data foi preenchida
    If Len(Trim(Data.ClipText)) = 0 Then gError 22954

    If objGrid.iLinhasExistentes = 0 Then gError 22953
    
    'Loop de Validação dos dados do Grid
    For iIndice = 1 To objGrid.iLinhasExistentes

        'Verifica se a quantidade foi digitada
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 31515

        'Verifica se o almoxarifado foi informado
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Almoxarifado_Col))) = 0 And Left(GridMovimentos.TextMatrix(iIndice, 0), 1) <> "#" Then gError 31516

        'Verifica se a data de previsão de inicio da OP foi informada
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevInicio_Col))) = 0 Then gError 41358

        'Verifica se a data de previsão de fim da OP foi informada e se é menor que a data de início
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevFim_Col))) > 0 Then
            If CDate(GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevFim_Col)) < CDate(GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevInicio_Col)) Then gError 41336
        Else
            gError 41337
        End If

        'Verifica se a Situação é Baixada para ítens novos
        sSituacao = GridMovimentos.TextMatrix(iIndice, iGrid_Situacao_Col)
        If Len(Trim(sSituacao)) > 0 Then
            If sSituacao = STRING_BAIXADA And gcolItemOP.Item(iIndice) = 0 Then gError 41327
        End If
        
        'Verifica se é uma ordem de producao baixada
        If StatusOP.Caption = STRING_STATUS_BAIXADO Then
            
            'Verifica se a Situação é Normal para o item
            If Len(Trim(sSituacao)) > 0 Then
                
                If sSituacao = STRING_NORMAL Then bAchou = True
                
            End If
            
        End If
        
        ' se destinacao for PV ==> filialPV e COdPV tem que estar preenchidos
        sPedidoDeVenda = GridMovimentos.TextMatrix(iIndice, iGrid_Destinacao_Col)
        If Len(Trim(sPedidoDeVenda)) > 0 Then

            If sPedidoDeVenda = STRING_PEDIDOVENDA Then

                If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_PedidoDeVenda_Col))) = 0 Then gError 41328
                If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_FilialPedido_Col))) = 0 Then gError 41329

            End If

        End If

    Next
    
    'se a Op está baixada e existe item com situacao='normal'
    If bAchou = True Then
    
        vbMsg = Rotina_Aviso(vbYesNo, "AVISO_REATIVACAO_OP", Codigo.Text)
        'se não for reativar a OP sai da gravação
        If vbMsg = vbNo Then gError 82804
    
    ElseIf bAchou = False And StatusOP.Caption = STRING_STATUS_BAIXADO Then
        gError 82803
    End If
    
    If StatusOP.Caption = STRING_STATUS_BAIXADO Then
        objOrdemDeProducao.iStatusOP = ITEMOP_SITUACAO_BAIXADA
    
    ElseIf StatusOP.Caption = STRING_NORMAL Then
        objOrdemDeProducao.iStatusOP = ITEMOP_SITUACAO_NORMAL
    End If
    
    lErro = Move_Tela_Memoria(objOrdemDeProducao)
    If lErro <> SUCESSO Then gError 31518

    lErro = CF("OrdemDeProducao_Grava", objOrdemDeProducao)
    If lErro <> SUCESSO Then gError 31520

    'Se a opcao de imprimir o Relatorio estiver marcada
    If ImprimeAoGravar.Value = MARCADO Then
        
        'Gera o(s) Relatorio(s)
        lErro = Executa_Relatorio(objOrdemDeProducao)
        If lErro <> SUCESSO Then gError 106405
        
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 22953
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NENHUM_ITEMOP_INFORMADO", gErr)

        Case 22954, 41358, 41337
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)

        Case 31513
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGOOP_NAO_PREENCHIDO", gErr)

        Case 31515
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA", gErr, iIndice)

        Case 31516
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_PREENCHIDO", gErr, iIndice)

        Case 31518, 31520, 82804, 106405

        Case 41327
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BAIXAR_ITEMNOVO", gErr)
        
        Case 41328
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDAID_NAO_PREENCHIDO", gErr, iIndice)

        Case 41329
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALPEDIDO_NAO_PREENCHIDA", gErr, iIndice)

        Case 41336
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_FIM_MENOR", gErr)

        Case 62562
''''            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODOP_GERAR_NAO_INFORMADO", gErr)
        
        Case 82803
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OPBAIXADA_NAO_REATIVADA", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163752)

    End Select

    Exit Function

End Function

Function Move_Grid_Memoria(objOrdemDeProducao As ClassOrdemDeProducao) As Long
'move itens do Grid para objOrdemDeProducao

Dim lErro As Long
Dim iIndice As Integer, iCount As Integer
Dim iProdutoPreenchido As Integer
Dim sProduto As String, sCCL As String, sCclFormatada As String, iCclPreenchida As Integer
Dim sProdutoFormatado As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objItemOP As ClassItemOP
Dim sSituacao As String, sDestinacao As String, sFilialPedido As String
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome
Dim objMaquina As New ClassMaquinas

On Error GoTo Erro_Move_Grid_Memoria

    objOrdemDeProducao.iNumItens = 0
    objOrdemDeProducao.iNumItensBaixados = 0

    For iIndice = 1 To objGrid.iLinhasExistentes

        Set objItemOP = New ClassItemOP

        If gobjOP.colItens.Count >= iIndice Then

            objItemOP.lNumIntDoc = gobjOP.colItens(iIndice).lNumIntDoc
        
        End If

        objItemOP.sCodigo = objOrdemDeProducao.sCodigo
        objItemOP.iFilialEmpresa = objOrdemDeProducao.iFilialEmpresa
        objItemOP.iTipo = OP_TIPO_OC

        sProduto = GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col)

        'Critica o formato do Produto
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 31563

        objItemOP.sProduto = sProdutoFormatado

        objItemOP.sSiglaUM = GridMovimentos.TextMatrix(iIndice, iGrid_UnidadeMed_Col)

        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))) > 0 Then
            objItemOP.dQuantidade = CDbl(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))
        Else
            objItemOP.dQuantidade = 0
        End If

        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_MetragemCons_Col))) > 0 Then
            objItemOP.dMetragemCons = CDbl(GridMovimentos.TextMatrix(iIndice, iGrid_MetragemCons_Col))
        Else
            objItemOP.dMetragemCons = 0
        End If

        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Enfesto_Col))) > 0 Then
            objItemOP.dEnfesto = CDbl(GridMovimentos.TextMatrix(iIndice, iGrid_Enfesto_Col))
        Else
            objItemOP.dEnfesto = 0
        End If

        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Risco_Col))) > 0 Then
            objItemOP.dRisco = CDbl(GridMovimentos.TextMatrix(iIndice, iGrid_Risco_Col))
        Else
            objItemOP.dRisco = 0
        End If
        
        'se nao for pai de grade
        If Left(GridMovimentos.TextMatrix(iIndice, 0), 1) <> "#" Then

            objAlmoxarifado.sNomeReduzido = GridMovimentos.TextMatrix(iIndice, iGrid_Almoxarifado_Col)
            objItemOP.iPossuiGrade = DESMARCADO
    
            If colCodigoNome.Count > 0 Then
                
                For Each objCodigoNome In colCodigoNome
                    If objCodigoNome.sNome = objAlmoxarifado.sNomeReduzido Then
                        objItemOP.iAlmoxarifado = objCodigoNome.iCodigo
                        Exit For
                    End If
                Next
            
            End If
                    
            If objItemOP.iAlmoxarifado = 0 Then
    
                'lê nome reduzido do almoxarifado
                lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25060 Then gError 19307
        
                If lErro = 25060 Then gError 19308
        
                objItemOP.iAlmoxarifado = objAlmoxarifado.iCodigo
        
                colCodigoNome.Add objAlmoxarifado.iCodigo, objAlmoxarifado.sNomeReduzido
                
            End If
        
        Else
        
            objItemOP.iPossuiGrade = MARCADO
        
        End If
        
        objItemOP.iBeneficiamento = StrParaInt(GridMovimentos.TextMatrix(iIndice, iGrid_Benef_Col))

        sCCL = GridMovimentos.TextMatrix(iIndice, iGrid_Ccl_Col)

        If Len(Trim(sCCL)) <> 0 Then

            'Formata Ccl para BD
            lErro = CF("Ccl_Formata", sCCL, sCclFormatada, iCclPreenchida)
            If lErro <> SUCESSO Then gError 19309

        Else
            sCclFormatada = ""
        End If

        objItemOP.sCCL = sCclFormatada

        If Len(GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevInicio_Col)) > 0 Then
            objItemOP.dtDataInicioProd = CDate(GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevInicio_Col))
        Else
            objItemOP.dtDataInicioProd = DATA_NULA
        End If

        If Len(GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevFim_Col)) > 0 Then
            objItemOP.dtDataFimProd = CDate(GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevFim_Col))
        Else
            objItemOP.dtDataFimProd = DATA_NULA
        End If

        'Seleciona a situação
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Situacao_Col))) > 0 Then
            sSituacao = GridMovimentos.TextMatrix(iIndice, iGrid_Situacao_Col)
            For iCount = 0 To Situacao.ListCount - 1
                If Situacao.List(iCount) = sSituacao Then
                    objItemOP.iSituacao = Situacao.ItemData(iCount)
                    Exit For
                End If
            Next
        End If

        'seleciona a destinação , junto com FilialPedido,PedidoDeVenda e ItemPV
        sDestinacao = GridMovimentos.TextMatrix(iIndice, iGrid_Destinacao_Col)
        If Len(Trim(sDestinacao)) > 0 Then

            If sDestinacao = STRING_PEDIDOVENDA Then
                sFilialPedido = GridMovimentos.TextMatrix(iIndice, iGrid_FilialPedido_Col)
                If Len(Trim(sFilialPedido)) > 0 Then
                    For iCount = 0 To ComboFilialPedido.ListCount - 1
                        If ComboFilialPedido.List(iCount) = sFilialPedido Then
                            objItemOP.iFilialPedido = ComboFilialPedido.ItemData(iCount)
                            Exit For
                        End If
                    Next
                End If
                If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_PedidoDeVenda_Col))) > 0 Then objItemOP.lCodPedido = CLng(GridMovimentos.TextMatrix(iIndice, iGrid_PedidoDeVenda_Col))
            End If

            'seleciona a Destinação
            For iCount = 0 To Destinacao.ListCount - 1
                If Destinacao.List(iCount) = sDestinacao Then
                    objItemOP.iDestinacao = Destinacao.ItemData(iCount)
                    Exit For
                End If
            Next

        End If

        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Prioridade_Col))) > 0 Then objItemOP.iPrioridade = CInt(GridMovimentos.TextMatrix(iIndice, iGrid_Prioridade_Col))

        objItemOP.iItem = iIndice
        
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Maquina_Col))) > 0 Then
            
            'Le a Maquina a partir do nome reduzido
            objMaquina.sNomeReduzido = GridMovimentos.TextMatrix(iIndice, iGrid_Maquina_Col)
            
            lErro = CF("Maquinas_Le_NomeReduzido", objMaquina)
            If lErro <> SUCESSO And lErro <> 103100 Then gError 106345
            
            'Se nao encontrou => Erro
            If lErro = 103100 Then gError 106346
            
            objItemOP.lNumIntEquipamento = objMaquina.lNumIntDoc
            
        End If
        
        objItemOP.sVersao = GridMovimentos.TextMatrix(iIndice, iGrid_Versao_Col)

        objOrdemDeProducao.colItens.Add objItemOP

        If objItemOP.iPossuiGrade = DESMARCADO Then

            objOrdemDeProducao.iNumItens = objOrdemDeProducao.iNumItens + 1
            If objItemOP.iSituacao = ITEMOP_SITUACAO_BAIXADA Then objOrdemDeProducao.iNumItensBaixados = objOrdemDeProducao.iNumItensBaixados + 1

        End If

        '********************* TRATAMENTO DE GRADE *****************
        Call Move_ItensGrade_Tela(objOrdemDeProducao, objItemOP, objItemOP.colItensRomaneioGrade, gobjOP.colItens(iIndice).colItensRomaneioGrade)
    
    Next

    Move_Grid_Memoria = SUCESSO

    Exit Function

Erro_Move_Grid_Memoria:

    Move_Grid_Memoria = gErr

    Select Case gErr

        Case 19307, 19309, 31563

        Case 19308
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case 106345
        
        Case 106346
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_NAO_CADASTRADA", gErr, objMaquina.sNomeReduzido)
            'ERRO_MAQUINA_NAO_CADASTRADA = "A máquina %s não está cadastrada."
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163753)

    End Select

    Exit Function

End Function

Function Move_ItensGrade_Tela(objOrdemProducao As ClassOrdemDeProducao, objItemOP As ClassItemOP, colItensRomaneio As Collection, colItensRomaneioTela As Collection) As Long

Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim objItemRomaneioGradeTela As ClassItemRomaneioGrade
Dim objReservaItemTela As ClassReservaItem
Dim objReservaItem As ClassReservaItem

    'Para cada Item de Romaneio vindo da tela ( Aqueles que já tem quantidade)
    For Each objItemRomaneioGradeTela In colItensRomaneioTela
                    
        Set objItemRomaneioGrade = New ClassItemRomaneioGrade
            
        objItemRomaneioGrade.sProduto = objItemRomaneioGradeTela.sProduto
        objItemRomaneioGrade.sDescricao = objItemRomaneioGradeTela.sDescricao
        objItemRomaneioGrade.dQuantidade = objItemRomaneioGradeTela.dQuantidade
        objItemRomaneioGrade.sUMEstoque = objItemRomaneioGradeTela.sUMEstoque
        objItemRomaneioGrade.iAlmoxarifado = objItemRomaneioGradeTela.iAlmoxarifado
        objItemRomaneioGrade.sAlmoxarifado = objItemRomaneioGradeTela.sAlmoxarifado
        objItemRomaneioGrade.sVersao = objItemRomaneioGradeTela.sVersao
        objItemRomaneioGrade.lNumIntDoc = objItemRomaneioGradeTela.lNumIntDoc
                    
        colItensRomaneio.Add objItemRomaneioGrade
        
        objOrdemProducao.iNumItens = objOrdemProducao.iNumItens + 1
        If objItemOP.iSituacao = ITEMOP_SITUACAO_BAIXADA Then objOrdemProducao.iNumItensBaixados = objOrdemProducao.iNumItensBaixados + 1
        
    Next

End Function

Function Limpa_Tela_OrdemDeProducao(Optional iFechaSetas As Integer = FECHAR_SETAS) As Long
'Limpa a Tela

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Tela_OrdemDeProducao

    If iFechaSetas = FECHAR_SETAS Then
    'Fecha o comando das setas se estiver aberto
     lErro = ComandoSeta_Fechar(Me.Name)
     If lErro <> SUCESSO Then gError 21801
    End If
    Call Limpa_Tela(Me)

    QuantDisponivel.Caption = ""
    StatusOP.Caption = ""
    
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    For iIndice = 1 To objGrid.iLinhasExistentes 'm
        GridMovimentos.TextMatrix(iIndice, 0) = iIndice
    Next
    
    Call Limpa_GridMovimentos

    Set gcolItemOP = New Collection

    Set gobjOP = New ClassOrdemDeProducao

    'Coloca a Data Atual na Tela
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True

    'Coloca a Data atual nas Datas de Previsão
    DataInicioPadrao.PromptInclude = False
    DataInicioPadrao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataInicioPadrao.PromptInclude = True

    DataFimPadrao.PromptInclude = False
    DataFimPadrao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataFimPadrao.PromptInclude = True

    ImprimeAoGravar.Value = DESMARCADO
    OpcaoRelatorio(2).Value = True
    
    iAlterado = 0
    iCodigoAlterado = 0

    Limpa_Tela_OrdemDeProducao = SUCESSO

    Exit Function

Erro_Limpa_Tela_OrdemDeProducao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163754)

    End Select

    Exit Function

End Function

Function Trata_Parametros(Optional objOrdemDeProducao As ClassOrdemDeProducao) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Trata_Parametros

    If Not (objOrdemDeProducao Is Nothing) Then

        'traz OP para a tela
        lErro = Traz_Tela_OrdemDeProducao(objOrdemDeProducao)
        If lErro <> SUCESSO And lErro <> 21966 Then gError 31556

        If lErro = 21966 Then

            'Se não existe exibe apenas o código
            Codigo.Text = objOrdemDeProducao.sCodigo

        End If

        Call ComandoSeta_Fechar(Me.Name)
                
    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 31556

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163755)

    End Select

    iAlterado = 0

    Exit Function

End Function


'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA
'""""""""""""""""""""""""""""""""""""""""""""""

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objOrdemDeProducao As New ClassOrdemDeProducao

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "OrdemProducaoOC"

    objOrdemDeProducao.sCodigo = Codigo.Text
    objOrdemDeProducao.dtDataEmissao = CDate(Data.Text)
    objOrdemDeProducao.iFilialEmpresa = giFilialEmpresa
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "FilialEmpresa", objOrdemDeProducao.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "Codigo", objOrdemDeProducao.sCodigo, STRING_ORDEM_DE_PRODUCAO, "Codigo"
    colCampoValor.Add "DataEmissao", objOrdemDeProducao.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "OPGeradora", objOrdemDeProducao.sOPGeradora, STRING_ORDEM_DE_PRODUCAO, "OPGeradora"

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        'Erro já tratado
        Case 31557

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163756)

    End Select

    Exit Sub

End Sub

Function Move_Tela_Memoria(objOrdemDeProducao As ClassOrdemDeProducao) As Long

Dim lErro As Long, objPrestServ As New ClassPrestServ

On Error GoTo Erro_Move_Tela_Memoria

    objOrdemDeProducao.sCodigo = Codigo.Text
    objOrdemDeProducao.dtDataEmissao = CDate(Data.Text)
    objOrdemDeProducao.iFilialEmpresa = giFilialEmpresa
    objOrdemDeProducao.iTipo = OP_TIPO_OC
    
    lErro = Move_Grid_Memoria(objOrdemDeProducao)
    If lErro <> SUCESSO Then gError 31519

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 31519, 124249
        
        Case 124250
            Call Rotina_Erro(vbOKOnly, "ERRO_PRESTSERV_NAO_CADASTRADO1", Err, objPrestServ.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163757)

    End Select

    Exit Function

End Function

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objOrdemDeProducao As New ClassOrdemDeProducao

On Error GoTo Erro_Tela_Preenche

    objOrdemDeProducao.sCodigo = colCampoValor.Item("Codigo").vValor
    objOrdemDeProducao.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
    objOrdemDeProducao.dtDataEmissao = colCampoValor.Item("DataEmissao").vValor
    objOrdemDeProducao.sOPGeradora = colCampoValor.Item("OPGeradora").vValor

    'Traz dados da Ordem de Produção para a Tela
    lErro = Traz_Tela_OrdemDeProducao(objOrdemDeProducao)
    If lErro <> SUCESSO And lErro <> 21966 Then gError 31558

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 31558

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163758)

    End Select

    Exit Sub

End Sub

Function Traz_Tela_OrdemDeProducao(objOrdemDeProducao As ClassOrdemDeProducao) As Long
'preenche a tela com os dados da OP

Dim lErro As Long

On Error GoTo Erro_Traz_Tela_OrdemDeProducao

    lErro = Limpa_Tela_OrdemDeProducao(NAO_FECHAR_SETAS)
    If lErro <> SUCESSO Then gError 41330

    Codigo.Text = CStr(objOrdemDeProducao.sCodigo)

    lErro = CF("OrdemDeProducao_Le_ComItens", objOrdemDeProducao)
    If lErro <> SUCESSO And lErro <> 21960 Then gError 21963

    If lErro = 21960 Then

        lErro = CF("OrdemDeProducaoBaixada_Le_ComItens", objOrdemDeProducao)
        If lErro <> SUCESSO And lErro <> 82797 Then gError 82801
        
        If lErro = 82797 Then gError 21966
        
        StatusOP.Caption = STRING_STATUS_BAIXADO
        
    End If
    
    If objOrdemDeProducao.iTipo = OP_TIPO_OP Then gError 117661
    
    
    Call DateParaMasked(Data, objOrdemDeProducao.dtDataEmissao)
    
    If Len(Trim(StatusOP.Caption)) = 0 Then
        StatusOP.Caption = "ATIVO"
    End If
    
    'preenche o grid
    lErro = Preenche_GridMovimentos(objOrdemDeProducao.colItens)
    If lErro <> SUCESSO Then gError 21972
    
'    lErro = CF("OPGeradora_Le_OPGerada", objOrdemDeProducao)
'    If lErro <> SUCESSO And lErro <> 62637 Then gError 62633
'
    iAlterado = 0
    iCodigoAlterado = 0

    Traz_Tela_OrdemDeProducao = SUCESSO

    Exit Function

Erro_Traz_Tela_OrdemDeProducao:

    Traz_Tela_OrdemDeProducao = gErr

    Select Case gErr

        Case 21963, 21972, 21966, 41330, 62633, 82801

        Case 117661
            Call Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO", gErr, objOrdemDeProducao.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163759)

    End Select

    Exit Function

End Function

Private Sub objEventoCcl_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclMascarado As String
Dim sCclFormatada As String

On Error GoTo Erro_objEventoCcl_evSelecao

    Set objCcl = obj1

    'Se o produto da linha corrente estiver preenchido e Linha corrente diferente da Linha fixa
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))) <> 0 And GridMovimentos.Row <> 0 Then

        sCclMascarado = String(STRING_CCL, 0)

        lErro = Mascara_MascararCcl(objCcl.sCCL, sCclMascarado)
        If lErro <> SUCESSO Then gError 22934

        'Coloca o valor do Ccl na coluna correspondente
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Ccl_Col) = sCclMascarado

        Ccl.PromptInclude = False
        Ccl.Text = sCclMascarado
        Ccl.PromptInclude = True

    End If

    Me.Show

    Exit Sub

Erro_objEventoCcl_evSelecao:

    Select Case gErr

        Case 22934

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163760)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOrdemDeProducao As ClassOrdemDeProducao

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objOrdemDeProducao = obj1

    'traz OP para a tela
    lErro = Traz_Tela_OrdemDeProducao(objOrdemDeProducao)
    If lErro <> SUCESSO Then gError 34675

    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 34675

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163761)
    End Select

    Exit Sub

End Sub

Private Sub PedidoDeVendaId_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PedidoDeVendaId_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)
    
End Sub

Private Sub PedidoDeVendaId_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub PedidoDeVendaId_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = PedidoDeVendaId
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Prioridade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Prioridade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Prioridade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Prioridade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Prioridade
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub PrioridadePadrao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(PrioridadePadrao, iAlterado)

End Sub

Private Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Produto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Quantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGrid.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MetragemCons_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MetragemCons_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub MetragemCons_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub MetragemCons_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGrid.objControle = MetragemCons
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Enfesto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Enfesto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Enfesto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Enfesto_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGrid.objControle = Enfesto
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Risco_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Risco_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Risco_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Risco_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGrid.objControle = Risco
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Situacao_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Situacao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Situacao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Situacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Situacao
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UpDownData_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_DownClick

    If Len(Trim(Data.ClipText)) = 0 Then Exit Sub

    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 22931

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 22931

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163762)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_UpClick

    If Len(Trim(Data.ClipText)) = 0 Then Exit Sub

    lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 22932

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 22932

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163763)

    End Select

    Exit Sub

End Sub

Private Sub UpDownInicio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownInicio_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownInicio_DownClick

    If Len(Trim(DataInicioPadrao.ClipText)) = 0 Then Exit Sub

    lErro = Data_Up_Down_Click(DataInicioPadrao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 41298

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_UpDownInicio_DownClick:

    Select Case gErr

        Case 41298

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163764)

    End Select

    Exit Sub

End Sub

Private Sub UpDownInicio_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownInicio_UpClick

    If Len(Trim(DataInicioPadrao.ClipText)) = 0 Then Exit Sub

    lErro = Data_Up_Down_Click(DataInicioPadrao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 41299

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_UpDownInicio_UpClick:

    Select Case gErr

        Case 41299

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163765)

    End Select

    Exit Sub

End Sub

Private Sub UpDownFim_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownFim_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownFim_DownClick

    If Len(Trim(DataFimPadrao.ClipText)) = 0 Then Exit Sub

    lErro = Data_Up_Down_Click(DataFimPadrao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 41300

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_UpDownFim_DownClick:

    Select Case gErr

        Case 41300

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163766)

    End Select

    Exit Sub

End Sub

Private Sub UpDownFim_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownFim_UpClick

    If Len(Trim(DataFimPadrao.ClipText)) = 0 Then Exit Sub

    lErro = Data_Up_Down_Click(DataFimPadrao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 41301

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_UpDownFim_UpClick:

    Select Case gErr

        Case 41301

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163767)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_gcolItemOP(gcolItemOP As Collection)
'limpa gcolItemOP e a coluna de qtde das linhas que estavam na tela apos a troca do codigo da OP

Dim lErro As Long
Dim iCount As Integer
Dim iIndice As Integer

On Error GoTo Erro_Limpa_gcolItemOp

    iCount = gcolItemOP.Count
    Set gcolItemOP = New Collection

    For iIndice = 1 To iCount

        gcolItemOP.Add 0

    Next

    Exit Sub

Erro_Limpa_gcolItemOp:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163768)

    End Select

    Exit Sub

End Sub

Private Function ProdutoLinha_Preenche(objProduto As ClassProduto, objItemOP As ClassItemOP) As Long

Dim iIndice As Integer
Dim lErro As Long
Dim iCclPreenchida As Integer
Dim sCclFormata As String
Dim sAlmoxarifadoPadrao As String

On Error GoTo Erro_ProdutoLinha_Preenche

    'Unidade de Medida
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMEstoque

    'Descricao
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_DescricaoItem_Col) = objProduto.sDescricao

    If Len(Trim(objProduto.sGrade)) = 0 Then

        'Almoxarifado
        '(Utiliza Almoxarifado Padrão caso esteja preenchido)
        If Len(Trim(AlmoxPadrao.ClipText)) > 0 And Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col))) = 0 Then
        
            lErro = CF("EstoqueProduto_TestaAssociacao", Produto.Text, AlmoxPadrao)
            
            If lErro = SUCESSO Then
                GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col) = AlmoxPadrao.Text
            Else
                GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col) = ""
            End If
        Else
    
            'le o Nome reduzido do almoxarifado Padrão do Produto em Questão
            lErro = CF("AlmoxarifadoPadrao_Le_NomeReduzido", objProduto.sCodigo, sAlmoxarifadoPadrao)
            If lErro <> SUCESSO Then gError 52282
    
            'preenche o grid
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col) = sAlmoxarifadoPadrao
    
        End If

    End If

    'Ccl
    lErro = CF("Ccl_Formata", CclPadrao.Text, sCclFormata, iCclPreenchida)
    If lErro <> SUCESSO Then gError 22940

    If iCclPreenchida = CCL_PREENCHIDA Then GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Ccl_Col) = CclPadrao.Text

    If Len(Trim(DataInicioPadrao.ClipText)) > 0 Then
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_DataPrevInicio_Col) = Format(DataInicioPadrao.Text, "dd/mm/yyyy")
    End If

    If Len(Trim(DataFimPadrao.ClipText)) > 0 Then
        If CDate(DataFimPadrao.Text) >= CDate(DataInicioPadrao.Text) Then
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_DataPrevFim_Col) = Format(DataFimPadrao.Text, "dd/mm/yyyy")
        Else
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_DataPrevFim_Col) = Format(DataInicioPadrao.Text, "dd/mm/yyyy")
        End If
    End If

    Situacao.ListIndex = ITEMOP_SITUACAO_NORMAL
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Situacao_Col) = Situacao.Text

    If DestinacaoPadrao.ListIndex <> -1 Then
        For iIndice = 0 To Destinacao.ListCount - 1
            If Destinacao.List(iIndice) = DestinacaoPadrao.Text Then
                Destinacao.ListIndex = iIndice
                GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Destinacao_Col) = Destinacao.Text
                Exit For
            End If
        Next
    Else
        
        For iIndice = 0 To Destinacao.ListCount - 1
            If Destinacao.ItemData(iIndice) = ITEMOP_DESTINACAO_ESTOQUE Then
                Destinacao.ListIndex = iIndice
                Exit For
            End If
        Next

        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Destinacao_Col) = Destinacao.Text
    End If

    If Len(Trim(PrioridadePadrao.ClipText)) > 0 Then
        Prioridade.PromptInclude = False
        Prioridade.Text = PrioridadePadrao.Text
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Prioridade_Col) = PrioridadePadrao.Text
        Prioridade.PromptInclude = True
    End If

    'ALTERAÇÃO DE LINHAS EXISTENTES
    If (GridMovimentos.Row - GridMovimentos.FixedRows) = objGrid.iLinhasExistentes Then
        objGrid.iLinhasExistentes = objGrid.iLinhasExistentes + 1
        gcolItemOP.Add 0
        gobjOP.colItens.Add objItemOP
        
        If Len(Trim(objProduto.sGrade)) = 0 Then

            gobjOP.colItens(GridMovimentos.Row).iPossuiGrade = DESMARCADO
            
        Else
        
            gobjOP.colItens(GridMovimentos.Row).iPossuiGrade = MARCADO
'            Set gobjOP.colItens(GridMovimentos.Row).colItensRomaneioGrade = objItemOP.colItensRomaneioGrade
            GridMovimentos.TextMatrix(GridMovimentos.Row, 0) = "# " & GridMovimentos.TextMatrix(GridMovimentos.Row, 0)
        
        End If
        
        gobjOP.colItens(GridMovimentos.Row).sSiglaUMEstoque = objProduto.sSiglaUMEstoque
        gobjOP.colItens(GridMovimentos.Row).iItem = GridMovimentos.Row
        gobjOP.colItens(GridMovimentos.Row).sProduto = objProduto.sCodigo
    
    End If

    ProdutoLinha_Preenche = SUCESSO

    Exit Function

Erro_ProdutoLinha_Preenche:

    ProdutoLinha_Preenche = gErr

    Select Case gErr

        Case 22940, 52282

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163769)

    End Select

    Exit Function

End Function

Private Function QuantDisponivel_Calcula(sProduto As String, sAlmoxarifado As String, Optional objProduto As ClassProduto) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sUnidadeMed As String
Dim dFator As Double
Dim dQuantTotal As Double
Dim dQuantidade As Double
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEstoqueProduto As New ClassEstoqueProduto

On Error GoTo Erro_QuantDisponivel_Calcula

    'Verifica se o produto está preenchido
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 41350

    If GridMovimentos.Row >= GridMovimentos.FixedRows And Len(Trim(sAlmoxarifado)) <> 0 And iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        If objProduto Is Nothing Then

            Set objProduto = New ClassProduto

            objProduto.sCodigo = sProdutoFormatado

            'Lê o produto no BD para obter UM de estoque
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 41351

            If lErro = 28030 Then gError 41352

        End If

        objAlmoxarifado.sNomeReduzido = sAlmoxarifado

        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25060 Then gError 41353

        If lErro = 25060 Then gError 41354

        objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
        objEstoqueProduto.sProduto = sProdutoFormatado

        'Lê o Estoque Produto correspondente ao Produto e ao Almoxarifado
        lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
        If lErro <> SUCESSO And lErro <> 21306 Then gError 41355

        'Se não encontrou EstoqueProduto no Banco de Dados
        If lErro = 21306 Then

            QuantDisponivel.Caption = Formata_Estoque(0)

        Else
            sUnidadeMed = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col)

            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProduto.sSiglaUMEstoque, sUnidadeMed, dFator)
            If lErro <> SUCESSO Then gError 41356
            
            If StrParaInt(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Benef_Col)) = MARCADO Then
                
                QuantDisponivel.Caption = Formata_Estoque(objEstoqueProduto.dQuantBenef3 * dFator)
            
            Else
                
                QuantDisponivel.Caption = Formata_Estoque(objEstoqueProduto.dQuantDisponivel * dFator)
            
            End If

        End If

    Else

        'Limpa a Quantidade Disponível da Tela
        QuantDisponivel.Caption = ""

    End If


    QuantDisponivel_Calcula = SUCESSO

    Exit Function

Erro_QuantDisponivel_Calcula:

    QuantDisponivel_Calcula = gErr

    Select Case gErr

        Case 41350, 41351, 41353, 41355, 41356

        Case 41352
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 41354
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163770)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_ORDEM_PRODUCAO
    Set Form_Load_Ocx = Me
    Caption = "Ordem de Corte"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "OrdemCorte"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Codigo Then
            Call CodigoOPLabel_Click
        ElseIf Me.ActiveControl Is CclPadrao Then
            Call CclPadraoLabel_Click
        ElseIf Me.ActiveControl Is Maquina Then
            Call BotaoMaquinas_Click
        ElseIf Me.ActiveControl Is AlmoxPadrao Then
            Call AlmoxPadraoLabel_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
        ElseIf Me.ActiveControl Is Almoxarifado Then
            Call BotaoEstoque_Click
        ElseIf Me.ActiveControl Is Ccl Then
            Call BotaoCcls_Click
        ElseIf Me.ActiveControl Is PedidoDeVendaId Then
            Call BotaoPedidoDeVenda_Click
        End If
    End If

End Sub

Private Sub PrioridadePadraoLbl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PrioridadePadraoLbl, Source, X, Y)
End Sub

Private Sub PrioridadePadraoLbl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PrioridadePadraoLbl, Button, Shift, X, Y)
End Sub

Private Sub DestPadraoLbl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DestPadraoLbl, Source, X, Y)
End Sub

Private Sub DestPadraoLbl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DestPadraoLbl, Button, Shift, X, Y)
End Sub

Private Sub DataPrevFimLbl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataPrevFimLbl, Source, X, Y)
End Sub

Private Sub DataPrevFimLbl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataPrevFimLbl, Button, Shift, X, Y)
End Sub

Private Sub DataPrevIniLbl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataPrevIniLbl, Source, X, Y)
End Sub

Private Sub DataPrevIniLbl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataPrevIniLbl, Button, Shift, X, Y)
End Sub

Private Sub AlmoxPadraoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(AlmoxPadraoLabel, Source, X, Y)
End Sub

Private Sub AlmoxPadraoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(AlmoxPadraoLabel, Button, Shift, X, Y)
End Sub

Private Sub CclPadraoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CclPadraoLabel, Source, X, Y)
End Sub

Private Sub CclPadraoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CclPadraoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub QuantDisponivel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantDisponivel, Source, X, Y)
End Sub

Private Sub QuantDisponivel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantDisponivel, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub CodigoOPLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CodigoOPLabel, Source, X, Y)
End Sub

Private Sub CodigoOPLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CodigoOPLabel, Button, Shift, X, Y)
End Sub
Private Sub TabStrip1_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index
        
    End If

End Sub

Private Sub AlmoxPadrao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_AlmoxPadrao_Validate

    'Se almoxarifado estivar vazio sai da rotina
    If Len(Trim(AlmoxPadrao.Text)) = 0 Then Exit Sub

    'Verifica existência do Almoxarifado
    lErro = TP_Almoxarifado_Filial_Le(AlmoxPadrao, objAlmoxarifado, 0)
    If lErro <> SUCESSO And lErro <> 25136 And lErro <> 25143 Then gError 22980

    If lErro = 25136 Then gError 22981

    If lErro = 25143 Then gError 22982

    Exit Sub

Erro_AlmoxPadrao_Validate:

    Cancel = True


    Select Case gErr

        Case 22980

        Case 22981, 22982
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, AlmoxPadrao.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163771)

    End Select

    Exit Sub

End Sub


Private Sub Data_Validate(Cancel As Boolean)
'Critica se a Data da OP está preenchida corretamente
Dim lErro As Long

On Error GoTo Erro_Data_Validate

    If Len(Trim(Data.ClipText)) = 0 Then Exit Sub

    lErro = Data_Critica(Data.Text)
    If lErro <> SUCESSO Then gError 22933

    Exit Sub

Erro_Data_Validate:

    Cancel = True


    Select Case gErr

        Case 22933

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163772)

    End Select

    Exit Sub

End Sub


Private Sub DataFimPadrao_Validate(Cancel As Boolean)
'Critica se a Data Fim Padrao da OP está preenchida corretamente

Dim lErro As Long

On Error GoTo Erro_DataFimPadrao_Validate

    If Len(Trim(DataFimPadrao.ClipText)) = 0 Then Exit Sub

    lErro = Data_Critica(DataFimPadrao.Text)
    If lErro <> SUCESSO Then gError 55247

    Exit Sub

Erro_DataFimPadrao_Validate:

    Cancel = True


    Select Case gErr

        Case 55247

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163773)

    End Select

    Exit Sub

End Sub



Private Sub DataInicioPadrao_Validate(Cancel As Boolean)
'Critica se a Data de Inicio Padrao da OP está preenchida corretamente

Dim lErro As Long

On Error GoTo Erro_DataInicioPadrao_Validate

    If Len(Trim(DataInicioPadrao.ClipText)) = 0 Then Exit Sub

    lErro = Data_Critica(DataInicioPadrao.Text)
    If lErro <> SUCESSO Then gError 55246

    Exit Sub

Erro_DataInicioPadrao_Validate:

    Cancel = True


    Select Case gErr

        Case 55246

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163774)

    End Select

    Exit Sub

End Sub

Private Sub CclPadrao_Validate(Cancel As Boolean)
'verifica existência da Ccl informada

Dim lErro As Long, sCclFormatada As String
Dim objCcl As New ClassCcl

On Error GoTo Erro_CclPadrao_Validate

    'se Ccl não estiver preenchida sai da rotina
    If Len(Trim(CclPadrao.Text)) = 0 Then Exit Sub

    lErro = CF("Ccl_Critica", CclPadrao.Text, sCclFormatada, objCcl)
    If lErro <> SUCESSO And lErro <> 5703 Then gError 31558

    If lErro = 5703 Then gError 31559

    Exit Sub

Erro_CclPadrao_Validate:

    Cancel = True


    Select Case gErr

        Case 31558

        Case 31559
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, CclPadrao.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163775)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objOrdemDeProducao As New ClassOrdemDeProducao
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Codigo_Validate

    'Se houve alteração nos dados da tela
    If (iCodigoAlterado = REGISTRO_ALTERADO) Then

        If Len(Trim(Codigo.Text)) > 0 Then

            'limpa a coleção global
'            Call Limpa_gcolItemOP(gcolItemOP)

            objOrdemDeProducao.sCodigo = Codigo.Text
            objOrdemDeProducao.iFilialEmpresa = giFilialEmpresa

            'tenta ler a OP desejada
            lErro = CF("OrdemProducao_Le", objOrdemDeProducao)
            If lErro <> SUCESSO And lErro <> 30368 And lErro <> 55316 Then gError 41306

            If lErro = SUCESSO And objOrdemDeProducao.iTipo = OP_TIPO_OP Then gError 117660

            'ordem de producao baixada
            If lErro = 55316 Then gError 55319

            'se existir
            If lErro = SUCESSO Then

                vbMsg = Rotina_Aviso(vbYesNo, "AVISO_PREENCHER_TELA")

                If vbMsg = vbNo Then gError 41307

                'traz a OP para a tela
                lErro = Traz_Tela_OrdemDeProducao(objOrdemDeProducao)
                If lErro <> SUCESSO And lErro <> 21966 Then gError 21973

                Call ComandoSeta_Fechar(Me.Name)

            End If

        End If

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 21973
    
        Case 41306
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDENSDEPRODUCAO", gErr)
    
        Case 41307
    
        Case 55319
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_BAIXADA", gErr, Codigo.Text)

        Case 117660
            Call Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO", gErr, Codigo.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163776)

    End Select

    Exit Sub

End Sub

Private Sub GridMovimentos_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGrid)

End Sub

Private Sub Carrega_ComboVersoes(ByVal sProdutoRaiz As String)
    
Dim lErro As Long
Dim objKit As New ClassKit
Dim colKits As New Collection
Dim iPadrao As Integer
Dim iIndice As Integer
    
On Error GoTo Erro_Carrega_ComboVersoes
    
    Versao.Enabled = True
    
    'Limpa a Combo
    Versao.Clear
    
    'Armazena o Produto Raiz do kit
    objKit.sProdutoRaiz = sProdutoRaiz
    
    'Le as Versoes Ativas e a Padrao
    lErro = CF("Kit_Le_Produziveis", objKit, colKits)
    If lErro <> SUCESSO And lErro <> 106333 Then gError 106321
    
    iPadrao = -1
    
    'Carrega a Combo com os Dados da Colecao
    For Each objKit In colKits
    
        Versao.AddItem (objKit.sVersao)
        
        'Se for a padrao -> Armazena
        If objKit.iSituacao = KIT_SITUACAO_PADRAO Then iPadrao = iIndice
        
        iIndice = iIndice + 1
        
    Next
    
    If Len(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Versao_Col)) > 0 Then
    
        Versao.Text = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Versao_Col)
    
    ElseIf iPadrao <> -1 Then
    
        'Seleciona a Padrao na Combo
        Versao.ListIndex = iPadrao
        
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Versao_Col) = Versao.Text

    End If



    Exit Sub
    
Erro_Carrega_ComboVersoes:

    Select Case gErr
    
        Case 106321
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163777)
    
    End Select
    
End Sub

Private Function Saida_Celula_Versao(objGridInt As AdmGrid) As Long
'faz a critica da celula de Versao do grid que está deixando de ser a corrente

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProdutoFilial As New ClassProdutoFilial

On Error GoTo Erro_Saida_Celula_Versao

    Set objGridInt.objControle = Versao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 106338

    Saida_Celula_Versao = SUCESSO

    Exit Function

Erro_Saida_Celula_Versao:

    Saida_Celula_Versao = gErr

    Select Case gErr

        Case 106338
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163778)

    End Select

End Function

Private Function Saida_Celula_Maquina(objGridInt As AdmGrid) As Long
'faz a critica da celula de Equipamento do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objMaquina As New ClassMaquinas

On Error GoTo Erro_Saida_Celula_Maquina

    Set objGridInt.objControle = Maquina

    'Se a Máquina foi especificada => Faz a Validacao da Máquina
    If Len(Trim(Maquina.Text)) > 0 Then
        
        'Verifica sua existencia
        lErro = CF("TP_Maquina_Le", Maquina, objMaquina)
        If lErro <> SUCESSO Then gError 106341
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 106338

    Saida_Celula_Maquina = SUCESSO

    Exit Function

Erro_Saida_Celula_Maquina:

    Saida_Celula_Maquina = gErr

    Select Case gErr

        Case 106338, 106341
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163779)

    End Select

End Function

Private Sub Maquina_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Maquina_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Maquina_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Maquina
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Versao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Versao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Versao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Versao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Versao
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Executa_Relatorio(ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long
'Executa o(s) relatorio(s) de acordo com a selecao no frame de relatorios

Dim lErro As Long, lNumIntRel As Long
Dim objRelatorio1 As New AdmRelatorio, objRelatorio2 As New AdmRelatorio

On Error GoTo Erro_Executa_Relatorio

    'Executa o Relatorio de OP caso o usuário tenha marcado somente o relatório de OP ou ambos
    If OpcaoRelatorio.Item(0).Value = True Or OpcaoRelatorio.Item(2).Value = True Then
    
        lErro = CF("ItensOPRel_Prepara", objOrdemDeProducao, lNumIntRel)
        If lErro <> SUCESSO Then gError 111828
    
        'Imprime o Relatorio de OP
        lErro = objRelatorio1.ExecutarDireto("Ordens de Produção", "OrdemProducao = @TORDPROD", 0, "OPINPAL", "TORDPROD", objOrdemDeProducao.sCodigo, "NNUMINTREL", CStr(lNumIntRel))
        If lErro <> SUCESSO Then gError 106406
    
    End If
    
    'Executa o Relatório de Rótulos caso o usuário tenha marcado somente o relatório de Rótulos ou ambos
    If OpcaoRelatorio.Item(1).Value = True Or OpcaoRelatorio.Item(2).Value = True Then
    
        'Imprime o Relatório de Rótulos
        lErro = objRelatorio2.ExecutarDireto("Rótulos para Ordens de Produção", "OrdemProducao = @TORDPROD", 0, "SOLROTUL", "TORDPROD", objOrdemDeProducao.sCodigo)
        If lErro <> SUCESSO Then gError 106408
        
    End If
    
    Executa_Relatorio = SUCESSO
    
    Exit Function
    
Erro_Executa_Relatorio:

    Executa_Relatorio = gErr
    
    Select Case gErr
    
        Case 106406, 106408, 111828
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163780)
    
    End Select

End Function

Private Function Grid_Possui_Grade() As Boolean

Dim iIndice As Integer

    For iIndice = 1 To objGrid.iLinhasExistentes
        If Left(GridMovimentos.TextMatrix(iIndice, 0), 2) = "# " Then
            Grid_Possui_Grade = True
            Exit Function
        End If
    Next
    
    Grid_Possui_Grade = False
        
    Exit Function
    
End Function

Sub Atualiza_Grid_Movimentos(objItemOP As ClassItemOP)

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim dQuantidade As Double
Dim objItemRomaneioGrade As ClassItemRomaneioGrade
    
    For Each objItemRomaneioGrade In objItemOP.colItensRomaneioGrade
            
        dQuantidade = dQuantidade + objItemRomaneioGrade.dQuantidade
        
    Next

    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col) = Formata_Estoque(dQuantidade)

    objItemOP.dQuantidade = dQuantidade

    Exit Sub

End Sub

Private Sub BotaoImprimirPrevia_Click()

Dim lErro As Long
Dim objCalcNecesProd As New ClassCalcNecesProd
Dim iIndice As Integer
Dim iProdutoPreenchido As Integer
Dim sProduto As String, sCCL As String
Dim sProdutoFormatado As String, bAchou As Boolean
Dim objNecesProdInfo As ClassNecesProdInfo
Dim objItemOP As ClassItemOP
Dim objItemRomaneioGrade As ClassItemRomaneioGrade

On Error GoTo Erro_BotaoImprimirPrevia_Click

    bAchou = False
    
    'Para cada item do grid, guarda em um objeto os dados do grid
    For iIndice = 1 To objGrid.iLinhasExistentes

        If Left(GridMovimentos.TextMatrix(iIndice, 0), 1) <> "#" Then

            Set objNecesProdInfo = New ClassNecesProdInfo
    
            sProduto = GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col)

            'Critica o formato do Produto
            lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 124215
    
            objNecesProdInfo.sProduto = sProdutoFormatado
    
            objNecesProdInfo.sUMNecesInfo = GridMovimentos.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
    
            If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))) > 0 Then
                objNecesProdInfo.dQuantNecesInfo = CDbl(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))
                If objNecesProdInfo.dQuantNecesInfo > 0 Then bAchou = True
            Else
                objNecesProdInfo.dQuantNecesInfo = 0
            End If
        
            objCalcNecesProd.colNecesInfProd.Add objNecesProdInfo
    
        Else
    
            Set objItemOP = gobjOP.colItens(iIndice)
            
            For Each objItemRomaneioGrade In objItemOP.colItensRomaneioGrade
            
                Set objNecesProdInfo = New ClassNecesProdInfo
        
                objNecesProdInfo.sProduto = objItemRomaneioGrade.sProduto
        
                objNecesProdInfo.sUMNecesInfo = objItemOP.sSiglaUM
        
                If objItemRomaneioGrade.dQuantidade > 0 Then
                    objNecesProdInfo.dQuantNecesInfo = objItemRomaneioGrade.dQuantidade
                    If objNecesProdInfo.dQuantNecesInfo > 0 Then bAchou = True
                Else
                    objNecesProdInfo.dQuantNecesInfo = 0
                End If
            
                objCalcNecesProd.colNecesInfProd.Add objNecesProdInfo
            
            Next
    
        End If
        
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163781)
     
    End Select
     
    Exit Sub

End Sub


