VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RoteiroSRVOcx 
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
      Caption         =   "J"
      Height          =   5085
      Index           =   2
      Left            =   225
      TabIndex        =   37
      Top             =   720
      Visible         =   0   'False
      Width           =   9060
      Begin VB.Frame Frame2 
         Caption         =   "Peças"
         Height          =   2655
         Index           =   2
         Left            =   90
         TabIndex        =   39
         Top             =   2295
         Visible         =   0   'False
         Width           =   8775
         Begin VB.TextBox MPObs 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   5295
            TabIndex        =   85
            Top             =   390
            Width           =   1245
         End
         Begin VB.CommandButton BotaoAbrirRoteiro 
            Caption         =   "Roteiro de Fabricação"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4295
            TabIndex        =   25
            ToolTipText     =   "Abre a tela de Roteiro de Fabricação para o Insumo"
            Top             =   2145
            Width           =   2400
         End
         Begin VB.TextBox MPOrigem 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   3045
            TabIndex        =   71
            Top             =   1050
            Width           =   375
         End
         Begin VB.ComboBox MPVersao 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "RoteiroSRV.ctx":0000
            Left            =   4590
            List            =   "RoteiroSRV.ctx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   70
            Top             =   1050
            Width           =   855
         End
         Begin VB.ComboBox MPComp 
            Height          =   315
            ItemData        =   "RoteiroSRV.ctx":0004
            Left            =   6255
            List            =   "RoteiroSRV.ctx":000E
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   1050
            Width           =   990
         End
         Begin VB.CommandButton BotaoProdutos 
            Caption         =   "&Produtos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   105
            TabIndex        =   24
            ToolTipText     =   "Abre o Browse de Produtos"
            Top             =   2130
            Width           =   1620
         End
         Begin VB.CommandButton BotaoLimparGrid 
            Caption         =   "&Limpar Grid"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   7005
            TabIndex        =   26
            ToolTipText     =   "Limpa os dados do Grid"
            Top             =   2130
            Width           =   1620
         End
         Begin VB.ComboBox MPUM 
            Height          =   315
            Left            =   3705
            TabIndex        =   41
            Top             =   1035
            Width           =   720
         End
         Begin VB.TextBox MPDesc 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   1200
            TabIndex        =   40
            Top             =   1035
            Width           =   1620
         End
         Begin MSMask.MaskEdBox MPProduto 
            Height          =   315
            Left            =   465
            TabIndex        =   42
            Top             =   600
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MPQtd 
            Height          =   315
            Left            =   2970
            TabIndex        =   43
            Top             =   450
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   15
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridMP 
            Height          =   1500
            Left            =   60
            TabIndex        =   23
            Top             =   300
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   2646
            _Version        =   393216
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Máquinas"
         Height          =   2670
         Index           =   3
         Left            =   90
         TabIndex        =   66
         Top             =   2295
         Visible         =   0   'False
         Width           =   8775
         Begin VB.CommandButton BotaoLimparGrid 
            Caption         =   "&Limpar Grid"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   7005
            TabIndex        =   29
            ToolTipText     =   "Limpa os dados do Grid"
            Top             =   2130
            Width           =   1620
         End
         Begin VB.CommandButton BotaoMaq 
            Caption         =   "Máquinas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   105
            TabIndex        =   28
            Top             =   2130
            Width           =   1620
         End
         Begin VB.TextBox MaqOBS 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   4275
            MaxLength       =   50
            TabIndex        =   82
            Top             =   945
            Width           =   3375
         End
         Begin VB.TextBox MaqCodigo 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   240
            MaxLength       =   20
            TabIndex        =   81
            Top             =   930
            Width           =   2520
         End
         Begin MSMask.MaskEdBox MaqHoras 
            Height          =   315
            Left            =   3000
            TabIndex        =   83
            Top             =   975
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaqQuantidade 
            Height          =   315
            Left            =   2025
            TabIndex        =   84
            Top             =   960
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridMaq 
            Height          =   345
            Left            =   60
            TabIndex        =   27
            Top             =   300
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   609
            _Version        =   393216
            Rows            =   15
            Cols            =   7
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
            AllowUserResizing=   1
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Mão de Obra"
         Height          =   2670
         Index           =   4
         Left            =   90
         TabIndex        =   75
         Top             =   2295
         Visible         =   0   'False
         Width           =   8775
         Begin VB.CommandButton BotaoLimparGrid 
            Caption         =   "&Limpar Grid"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   7005
            TabIndex        =   32
            ToolTipText     =   "Limpa os dados do Grid"
            Top             =   2130
            Width           =   1620
         End
         Begin VB.TextBox MOOBS 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   5295
            MaxLength       =   50
            TabIndex        =   78
            Top             =   720
            Width           =   2715
         End
         Begin VB.TextBox MODescricao 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   1170
            MaxLength       =   50
            TabIndex        =   77
            Top             =   720
            Width           =   2235
         End
         Begin VB.TextBox MOCodigo 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   285
            MaxLength       =   20
            TabIndex        =   76
            Top             =   720
            Width           =   870
         End
         Begin VB.CommandButton BotaoMO 
            Caption         =   "Mão de Obra"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   105
            TabIndex        =   31
            Top             =   2130
            Width           =   1620
         End
         Begin MSMask.MaskEdBox MOQuantidade 
            Height          =   315
            Left            =   3420
            TabIndex        =   79
            Top             =   720
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MOHoras 
            Height          =   315
            Left            =   4320
            TabIndex        =   80
            Top             =   720
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridMO 
            Height          =   315
            Left            =   60
            TabIndex        =   30
            Top             =   300
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   556
            _Version        =   393216
            Rows            =   15
            Cols            =   7
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
            AllowUserResizing=   1
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Operação"
         Height          =   2640
         Index           =   1
         Left            =   90
         TabIndex        =   44
         Top             =   2295
         Width           =   8775
         Begin VB.TextBox Observacao 
            Height          =   1005
            Left            =   1830
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   22
            Top             =   1155
            Width           =   6795
         End
         Begin MSMask.MaskEdBox CTPadrao 
            Height          =   315
            Left            =   1830
            TabIndex        =   21
            Top             =   690
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Competencia 
            Height          =   315
            Left            =   1845
            TabIndex        =   20
            Top             =   240
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label DescricaoCTPadrao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4320
            TabIndex        =   68
            Top             =   690
            Width           =   4305
         End
         Begin VB.Label DescricaoCompetencia 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4320
            TabIndex        =   67
            Top             =   240
            Width           =   4305
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Nível:"
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
            Left            =   1110
            TabIndex        =   65
            Top             =   2295
            Width           =   540
         End
         Begin VB.Label Sequencial 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3615
            TabIndex        =   64
            Top             =   2250
            Width           =   420
         End
         Begin VB.Label Nivel 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1830
            TabIndex        =   63
            Top             =   2250
            Width           =   420
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Sequencial:"
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
            Left            =   2460
            TabIndex        =   62
            Top             =   2295
            Width           =   1020
         End
         Begin VB.Label Label3 
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
            Height          =   330
            Left            =   510
            TabIndex        =   59
            Top             =   1185
            Width           =   1155
         End
         Begin VB.Label CTLabel 
            Caption         =   "CT Padrão:"
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
            Left            =   645
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   61
            Top             =   720
            Width           =   990
         End
         Begin VB.Label CompetenciaLabel 
            Caption         =   "Competência:"
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
            Height          =   330
            Left            =   465
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   60
            Top             =   270
            Width           =   1155
         End
      End
      Begin VB.CommandButton BotaoIncluir 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7605
         Picture         =   "RoteiroSRV.ctx":0022
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Inclui a Operação na Árvore do Roteiro"
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton BotaoRemover 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7605
         Picture         =   "RoteiroSRV.ctx":1870
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Exclui a Operação da Árvore do Roteiro"
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton BotaoAlterar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7605
         Picture         =   "RoteiroSRV.ctx":3196
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Altera a Operação da Árvore do Roteiro"
         Top             =   1305
         Width           =   1335
      End
      Begin MSComctlLib.TabStrip TabStrip2 
         Height          =   3150
         Left            =   15
         TabIndex        =   35
         Top             =   1905
         Width           =   8940
         _ExtentX        =   15769
         _ExtentY        =   5556
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   4
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Detalhe"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Peças"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Máquinas"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Mão de Obra"
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
      Begin VB.Frame Frame4 
         Caption         =   "Roteiro do Serviço"
         Height          =   1785
         Index           =   1
         Left            =   0
         TabIndex        =   38
         Top             =   30
         Width           =   7485
         Begin MSComctlLib.TreeView Roteiro 
            Height          =   1380
            Left            =   150
            TabIndex        =   16
            Top             =   270
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   2434
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   354
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            FullRowSelect   =   -1  'True
            Appearance      =   1
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5040
      Index           =   1
      Left            =   195
      TabIndex        =   36
      Top             =   750
      Width           =   9105
      Begin VB.CommandButton BotaoOnde 
         Caption         =   "Onde é Usado ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   1410
         Picture         =   "RoteiroSRV.ctx":4ABC
         TabIndex        =   9
         ToolTipText     =   "Lista dos Roteiros de Fabricação onde este roteiro é utilizado"
         Top             =   4365
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Frame Frame5 
         Caption         =   "Relatório"
         Height          =   870
         Left            =   6390
         TabIndex        =   72
         Top             =   4170
         Visible         =   0   'False
         Width           =   2625
         Begin VB.CommandButton BotaoRelRoteiro 
            Caption         =   "Roteiro Completo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   1260
            TabIndex        =   11
            ToolTipText     =   "Abre o Relatório de Roteiro de Fabricação"
            Top             =   210
            Width           =   1200
         End
         Begin VB.CheckBox DetalharInsumos 
            Caption         =   "Detalhar"
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
            Left            =   135
            TabIndex        =   10
            Top             =   345
            Width           =   1215
         End
      End
      Begin VB.CommandButton BotaoRoteiro 
         Caption         =   "&Roteiros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   120
         Picture         =   "RoteiroSRV.ctx":4DC6
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Roteiros de Fabricação cadastrados"
         Top             =   4365
         Width           =   1200
      End
      Begin VB.Frame Frame7 
         Caption         =   "Datas"
         Height          =   1245
         Index           =   1
         Left            =   2505
         TabIndex        =   54
         Top             =   2865
         Width           =   6510
         Begin MSComCtl2.UpDown UpDownDataCriacao 
            Height          =   300
            Left            =   3060
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   285
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataCriacao 
            Height          =   315
            Left            =   1890
            TabIndex        =   6
            Top             =   285
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label LabelDataUltModificacao 
            Caption         =   "Última modificação:"
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
            Left            =   150
            TabIndex        =   45
            Top             =   780
            Width           =   1755
         End
         Begin VB.Label LabelAutor 
            Alignment       =   1  'Right Justify
            Caption         =   "Autor:"
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
            Left            =   4110
            TabIndex        =   58
            Top             =   765
            Width           =   660
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Data de criação:"
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
            Left            =   375
            TabIndex        =   57
            Top             =   330
            Width           =   1440
         End
         Begin VB.Label DataUltModificacao 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1890
            TabIndex        =   56
            Top             =   735
            Width           =   1410
         End
         Begin VB.Label Autor 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4830
            TabIndex        =   55
            Top             =   735
            Width           =   1575
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Quantidade"
         Height          =   1245
         Left            =   105
         TabIndex        =   51
         Top             =   2850
         Width           =   2325
         Begin VB.ComboBox UM 
            Height          =   315
            Left            =   1230
            TabIndex        =   5
            Top             =   735
            Width           =   975
         End
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   315
            Left            =   1230
            TabIndex        =   4
            Top             =   300
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin VB.Label LabelQuantidade 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   105
            TabIndex        =   53
            Top             =   330
            Width           =   1035
         End
         Begin VB.Label LabelUM 
            Caption         =   "Un. Medida:"
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
            Height          =   315
            Left            =   75
            TabIndex        =   52
            Top             =   780
            Width           =   1260
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Identificação"
         Height          =   2580
         Left            =   105
         TabIndex        =   46
         Top             =   150
         Width           =   8895
         Begin VB.TextBox Descricao 
            Height          =   1320
            Left            =   1260
            MultiLine       =   -1  'True
            TabIndex        =   3
            Top             =   1140
            Width           =   7500
         End
         Begin MSMask.MaskEdBox Servico 
            Height          =   315
            Left            =   1260
            TabIndex        =   0
            Top             =   270
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Versao 
            Height          =   315
            Left            =   1260
            TabIndex        =   1
            Top             =   705
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Duracao 
            Height          =   315
            Left            =   7170
            TabIndex        =   2
            Top             =   705
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   3
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "dias uteis"
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
            Height          =   315
            Left            =   7800
            TabIndex        =   74
            Top             =   750
            Width           =   915
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Duração estimada:"
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
            Height          =   315
            Left            =   5370
            TabIndex        =   73
            Top             =   750
            Width           =   1755
         End
         Begin VB.Label LabelProduto 
            AutoSize        =   -1  'True
            Caption         =   "Serviço:"
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
            Left            =   450
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   50
            Top             =   330
            Width           =   720
         End
         Begin VB.Label DescricaoProd 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2865
            TabIndex        =   49
            Top             =   285
            Width           =   5880
         End
         Begin VB.Label LabelVersao 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   525
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   48
            Top             =   735
            Width           =   660
         End
         Begin VB.Label LabelDescricao 
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
            Left            =   240
            TabIndex        =   47
            Top             =   1125
            Width           =   930
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   7290
      ScaleHeight     =   480
      ScaleWidth      =   2025
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   105
      Width           =   2085
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "RoteiroSRV.ctx":50D0
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   570
         Picture         =   "RoteiroSRV.ctx":522A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Excluir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1065
         Picture         =   "RoteiroSRV.ctx":53B4
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1545
         Picture         =   "RoteiroSRV.ctx":58E6
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5520
      Left            =   120
      TabIndex        =   34
      Top             =   405
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   9737
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Roteiro"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Operações"
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
Attribute VB_Name = "RoteiroSRVOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iFrameAtual As Integer
Dim iFrameAtualOper As Integer

Dim colComponentes As New Collection
Dim iProxChave As Integer
Dim bOperacaoNova As Boolean

Dim glNumIntCompetAnt As Long
Dim glNumIntCTAnt As Long
Dim gsProdutoAnt As String

'variaveis auxiliares para recalculo de nivel e sequencial
Dim aNivelSequencial(NIVEL_MAXIMO_OPERACOES) As Integer 'para cada nivel guarda o maior sequencial
Dim aSeqPai(NIVEL_MAXIMO_OPERACOES) As Integer 'para cada nivel guarda o SeqPai

Dim iUltimoNivel As Integer

'Formato para quantidades de Produtos
Const FORMATO_ESTOQUE_KIT = "#,##0.0####"

'Grid de OperacaoInsumos
Dim objGridMP As AdmGrid
Dim iGrid_MPProduto_Col As Integer
Dim iGrid_MPDesc_Col As Integer
Dim iGrid_MPOrigem_Col As Integer
Dim iGrid_MPQtd_Col As Integer
Dim iGrid_MPUM_Col As Integer
Dim iGrid_MPVersao_Col As Integer
Dim iGrid_MPComp_Col As Integer
Dim iGrid_MPOBS_Col As Integer

'Grid de máquinas
Dim objGridMaq As AdmGrid
Dim iGrid_MaqCodigo_Col As Integer
Dim iGrid_MaqHoras_Col As Integer
Dim iGrid_MaqQuantidade_Col As Integer
Dim iGrid_MaqOBS_Col As Integer

'Grid de mão de obra
Dim objGridMO As AdmGrid
Dim iGrid_MOCodigo_Col As Integer
Dim iGrid_MODescricao_Col As Integer
Dim iGrid_MOHoras_Col As Integer
Dim iGrid_MOQuantidade_Col As Integer
Dim iGrid_MOOBS_Col As Integer

Private WithEvents objEventoRSRV As AdmEvento
Attribute objEventoRSRV.VB_VarHelpID = -1
Private WithEvents objEventoCompet As AdmEvento
Attribute objEventoCompet.VB_VarHelpID = -1
Private WithEvents objEventoCT As AdmEvento
Attribute objEventoCT.VB_VarHelpID = -1
Private WithEvents objEventoMP As AdmEvento
Attribute objEventoMP.VB_VarHelpID = -1
Private WithEvents objEventoVersao As AdmEvento
Attribute objEventoVersao.VB_VarHelpID = -1
Private WithEvents objEventoServico As AdmEvento
Attribute objEventoServico.VB_VarHelpID = -1
Private WithEvents objEventoMO As AdmEvento
Attribute objEventoMO.VB_VarHelpID = -1
Private WithEvents objEventoMaq As AdmEvento
Attribute objEventoMaq.VB_VarHelpID = -1

Private Const FRAME1_INDICE_IDENTIFICACAO = 1
Private Const FRAME1_INDICE_OPERACAO = 2

Private Const FRAME2_INDICE_DETALHE = 1
Private Const FRAME2_INDICE_MO = 2
Private Const FRAME2_INDICE_MP = 3
Private Const FRAME2_INDICE_MAQ = 4

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Roteiros de Serviços"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RoteiroSRV"

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

Private Sub BotaoAlterar_Click()

Dim lErro As Long
Dim objNode As Node
Dim sChave As String
Dim objOperacoes As New ClassRoteiroSRVOper
Dim objCompetencias As New ClassCompetencias
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim objProduto As New ClassProduto
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer
Dim sTexto As String

On Error GoTo Erro_BotaoAlterar_Click

    Set objNode = Roteiro.SelectedItem

    If objNode Is Nothing Then gError 194934
    If objNode.Selected = False Then gError 194935
    
    If Len(Trim(Competencia.ClipText)) = 0 Then gError 194948
    
    'preenche objOperacoes à partir dos dados da tela
    lErro = Move_Operacoes_Memoria(objOperacoes, objCompetencias, objCentrodeTrabalho)
    If lErro <> SUCESSO Then gError 194949

    sChave = objNode.Tag
        
    'prepara texto que identificará a nova Operação que está sendo incluida
    
    sTexto = objCompetencias.sNomeReduzido
    
    sCodProduto = Servico.Text
        
    'Critica o formato do MPProduto e se existe no BD
    lErro = CF("Produto_Critica", sCodProduto, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 25041 Then gError 194950
            
    sTexto = sTexto & " (" & objProduto.sNomeReduzido

    If Len(Trim(CTPadrao.ClipText)) <> 0 Then
       sTexto = sTexto & " - " & objCentrodeTrabalho.sNomeReduzido
    End If
        
    sTexto = sTexto & ")"

    objNode.Text = sTexto

    colComponentes.Remove (sChave)
    colComponentes.Add objOperacoes, sChave

    Call Recalcula_Nivel_Sequencial

    'Limpa a tab de operacoes
    lErro = Limpa_Operacoes()
    If lErro <> SUCESSO Then gError 194951

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_BotaoAlterar_Click:

    Select Case gErr

        Case 194934, 194935
            Call Rotina_Erro(vbOKOnly, "AVISO_SELECIONAR_ESTRUTURA_ROTEIRO", gErr)

        Case 194949, 194951, 194950
            'erro tratado na rotina chamada
        
        Case 194948
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_COMPETENCIA_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194959)

    End Select

    Exit Sub

End Sub

Private Sub BotaoIncluir_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim sChave As String
Dim sTexto As String
Dim objNode As Node
Dim objNodePai As Node
Dim sChaveTvw As String
Dim iNivel As Integer
Dim objOperacoes As New ClassRoteiroSRVOper
Dim objCompetencias As New ClassCompetencias
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim objProduto As New ClassProduto
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoIncluir_Click

    If Len(Trim(Competencia.ClipText)) = 0 Then gError 194960

    lErro = Move_Operacoes_Memoria(objOperacoes, objCompetencias, objCentrodeTrabalho)
    If lErro <> SUCESSO Then gError 194961

    Set objNodePai = Roteiro.SelectedItem

    If objNodePai Is Nothing Then
        iNivel = 0
    Else
        If objNodePai.Selected = False Then gError 194962
        iNivel = objNodePai.Index + 1
    End If

    'prepara texto que identificará a nova Operação que está sendo incluida
    sTexto = objCompetencias.sNomeReduzido
    
    sCodProduto = Servico.Text

    'Critica o formato do MPProduto e se existe no BD
    lErro = CF("Produto_Critica", sCodProduto, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 25041 Then gError 194963
            
    sTexto = sTexto & " (" & objProduto.sNomeReduzido

    If Len(Trim(CTPadrao.ClipText)) <> 0 Then
       sTexto = sTexto & " - " & objCentrodeTrabalho.sNomeReduzido
    End If
        
    sTexto = sTexto & ")"

    'prepara uma chave para relacionar colComponentes ao node que está sendo incluido
    Call Calcula_Proxima_Chave(sChaveTvw)
    
    sChave = sChaveTvw
    sChaveTvw = sChaveTvw & Competencia.ClipText
    
    'inclui o node na treeview
    If iNivel = 0 Then
        Set objNode = Roteiro.Nodes.Add(, tvwFirst, sChaveTvw, sTexto)
    Else
        Set objNode = Roteiro.Nodes.Add(objNodePai.Index, tvwChild, sChaveTvw, sTexto)
        If Not Roteiro.Nodes.Item(objNodePai.Index).Expanded Then
            Roteiro.Nodes.Item(objNodePai.Index).Expanded = True
        End If
    End If
    
    colComponentes.Add objOperacoes, sChave
    objNode.Tag = sChave

    Call Recalcula_Nivel_Sequencial

    'Limpa as Tabs de Detalhes, Insumos e Produção
    lErro = Limpa_Operacoes()
    If lErro <> SUCESSO Then gError 194976

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_BotaoIncluir_Click:

    Select Case gErr

        Case 194960
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_COMPETENCIA_NAO_PREENCHIDO", gErr)

        Case 194961, 194963, 194976
            'erro tratado na rotina chamada

        Case 194962
            Call Rotina_Erro(vbOKOnly, "AVISO_SELECIONAR_ESTRUTURA_ROTEIRO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194977)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimparGrid_Click(Index As Integer)

Dim lErro As Long

On Error GoTo Erro_BotaoLimparGrid_Click

    Select Case Index
    
        Case FRAME2_INDICE_MAQ
            Call Grid_Limpa(objGridMaq)
        
        Case FRAME2_INDICE_MP
            Call Grid_Limpa(objGridMP)
            
        Case FRAME2_INDICE_MO
            Call Grid_Limpa(objGridMO)
            
    End Select

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_BotaoLimparGrid_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194978)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoRemover_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objNode As Node
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoRemover_Click

    Set objNode = Roteiro.SelectedItem

    If objNode Is Nothing Then gError 194978
    If objNode.Selected = False Then gError 194979

    If objNode.Children > 0 Then

        'Envia aviso perguntando se realmente deseja excluir
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_PRODUTO_TEM_FILHOS")

        If vbMsgRes = vbNo Then gError 194987

        'chama rotina que exclui filhos
        Call Remove_Filhos(objNode.Child)
    
    End If

    colComponentes.Remove (objNode.Tag)
    Roteiro.Nodes.Remove (objNode.Index)
    
    Call Recalcula_Nivel_Sequencial

    'Limpa as Tabs de Detalhes, Insumos e Produção
    lErro = Limpa_Operacoes()
    If lErro <> SUCESSO Then gError 194988

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_BotaoRemover_Click:

    Select Case gErr

        Case 194978, 194979
            Call Rotina_Erro(vbOKOnly, "AVISO_SELECIONAR_ESTRUTURA_ROTEIRO", gErr)

        Case 194987, 194988

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194989)

    End Select

    Exit Sub

End Sub

Private Sub Competencia_GotFocus()
    Call MaskEdBox_TrataGotFocus(Competencia, iAlterado)
End Sub

Private Sub Competencia_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCompetencias As New ClassCompetencias
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho

On Error GoTo Erro_Competencia_Validate

    DescricaoCompetencia.Caption = ""

    'Verifica se Competencia não está preenchida
    If Len(Trim(Competencia.Text)) <> 0 Then
        
        'Verifica sua existencia
        lErro = CF("TP_Competencia_Le", Competencia, objCompetencias)
        If lErro <> SUCESSO Then gError 194920
        
        DescricaoCompetencia.Caption = objCompetencias.sDescricao
        
        If glNumIntCompetAnt <> objCompetencias.lNumIntDoc Then
        
            Observacao.Text = ""
            
            Call Grid_Limpa(objGridMP)
            Call Grid_Limpa(objGridMO)
            Call Grid_Limpa(objGridMaq)
                
            CTPadrao.Text = ""
            DescricaoCTPadrao.Caption = ""
            
            'Verifica se existe CTPadrao cadastrado na Competencia e traz seus dados
            lErro = CF("Competencias_Le_CTPadrao", objCompetencias, objCentrodeTrabalho)
            If lErro <> SUCESSO And lErro <> 134909 Then gError 194921
            
            If lErro = SUCESSO Then
            
               CTPadrao.Text = objCentrodeTrabalho.sNomeReduzido
               Call CTPadrao_Validate(bSGECancelDummy)
            
            End If
        
        End If
       
    End If

    Exit Sub

Erro_Competencia_Validate:

    Cancel = True

    Select Case gErr

        Case 194920, 194921
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194922)

    End Select

    Exit Sub

End Sub

Private Sub Competencia_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CTPadrao_GotFocus()
    Call MaskEdBox_TrataGotFocus(CTPadrao, iAlterado)
End Sub

Private Sub CTPadrao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim objCTCompetencias As New ClassCTCompetencias
Dim objCompetencias As ClassCompetencias
Dim bCompetenciaCadastrada As Boolean

On Error GoTo Erro_CTPadrao_Validate

    DescricaoCTPadrao.Caption = ""

    'Verifica se CTPadrao não está preenchido
    If Len(Trim(CTPadrao.Text)) <> 0 Then
        
        'Procura pela empresa toda
        objCentrodeTrabalho.iFilialEmpresa = EMPRESA_TODA
        
        'Verifica sua existencia
        lErro = CF("TP_CentrodeTrabalho_Le", CTPadrao, objCentrodeTrabalho)
        If lErro <> SUCESSO Then gError 196000
        
        Set objCompetencias = New ClassCompetencias
        
        objCompetencias.sNomeReduzido = Competencia.Text
        
        'Lê a Competencia pelo NomeReduzido para verificar seu NumIntDoc
        lErro = CF("Competencias_Le_NomeReduzido", objCompetencias)
        If lErro <> SUCESSO And lErro <> 134937 Then gError 196001
    
        If lErro <> SUCESSO Then gError 196002
        
        lErro = CF("CentrodeTrabalho_Le_CTCompetencias", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134453 Then gError 196003
    
        bCompetenciaCadastrada = False
        
        For Each objCTCompetencias In objCentrodeTrabalho.colCompetencias
            If objCTCompetencias.lNumIntDocCompet = objCompetencias.lNumIntDoc Then
                bCompetenciaCadastrada = True
                Exit For
            End If
        Next
            
        If bCompetenciaCadastrada = False Then gError 196004
            
        DescricaoCTPadrao.Caption = objCentrodeTrabalho.sDescricao
       
    End If
    
    Exit Sub

Erro_CTPadrao_Validate:

    Cancel = True

    Select Case gErr

        Case 196000, 196003, 196001
            'erro tratado na rotina chamada
        
        Case 196002, 196004
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPETENCIA_NAO_CADASTRADA_CT", gErr, objCentrodeTrabalho.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196005)

    End Select

    Exit Sub

End Sub

Private Sub CTPadrao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CTLabel_Click()

Dim lErro As Long
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim colSelecao As New Collection

On Error GoTo Erro_CTLabel

    'Verifica se o CTPadrao foi preenchido
    If Len(Trim(CTPadrao.Text)) <> 0 Then
    
        objCentrodeTrabalho.sNomeReduzido = CTPadrao.Text
        
        'Verifica o CTPadrao, lendo no BD a partir do NomeReduzido
        lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134941 Then gError 194990
        
    End If

    Call Chama_Tela("CentrodeTrabalhoLista", colSelecao, objCentrodeTrabalho, objEventoCT)

    Exit Sub

Erro_CTLabel:

    Select Case gErr
    
        Case 194990
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194991)

    End Select

    Exit Sub

End Sub

Private Sub CompetenciaLabel_Click()

Dim lErro As Long
Dim objCompetencias As New ClassCompetencias
Dim colSelecao As New Collection

On Error GoTo Erro_CompetenciaLabel_Click

    'Verifica se a Competencia foi preenchida
    If Len(Trim(Competencia.Text)) <> 0 Then
            
        objCompetencias.sNomeReduzido = Competencia.Text

        'Verifica a Competencia no BD a partir do NomeReduzido
        lErro = CF("Competencias_Le_NomeReduzido", objCompetencias)
        If lErro <> SUCESSO And lErro <> 134937 Then gError 194999

    End If

    Call Chama_Tela("CompetenciasLista", colSelecao, objCompetencias, objEventoCompet)

    Exit Sub

Erro_CompetenciaLabel_Click:

    Select Case gErr
    
        Case 194999
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194923)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProdutos_Click()

Dim lErro As Long
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoProdutos_Click

    If Me.ActiveControl Is MPProduto Then
        sProduto = MPProduto.Text
    Else
        'Verifica se tem alguma linha selecionada no Grid
        If GridMP.Row = 0 Then gError 196006
        
        sProduto = GridMP.TextMatrix(GridMP.Row, iGrid_MPProduto_Col)
    End If

    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 196007
    
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""
    
    objProduto.sCodigo = sProdutoFormatado
    
    'Lista de produtos produzíveis
    Call Chama_Tela("ProdutosKitLista", colSelecao, objProduto, objEventoMP)
    
    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr
        
        Case 196006
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 196007

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196008)

    End Select

    Exit Sub

End Sub

Private Sub BotaoRoteiro_Click()

Dim lErro As Long
Dim objRoteiroSRV As New ClassRoteiroSRV
Dim sProdutoFormatado As String
Dim colSelecao As New Collection
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoRoteiro_Click

    lErro = CF("Produto_Formata", Servico.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 196009

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""

    objRoteiroSRV.sServico = sProdutoFormatado
    
    If Len(Trim(Versao.Text)) <> 0 Then
        objRoteiroSRV.sVersao = Versao.Text
    End If

    Call Chama_Tela("RoteiroSRVLista", colSelecao, objRoteiroSRV, objEventoRSRV)

    Exit Sub

Erro_BotaoRoteiro_Click:

    Select Case gErr

        Case 196009

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196010)

    End Select

    Exit Sub

End Sub

Private Sub Duracao_Change()
 iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Duracao_GotFocus()
    Call MaskEdBox_TrataGotFocus(Duracao, iAlterado)
End Sub

Private Sub GridMP_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridMP)
End Sub

Private Sub GridMP_LostFocus()
    Call Grid_Libera_Foco(objGridMP)
End Sub

Private Sub objEventoCT_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objCentrodeTrabalho = obj1

    CTPadrao.Text = objCentrodeTrabalho.sNomeReduzido
    Call CTPadrao_Validate(bSGECancelDummy)
        
    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196011)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCompet_evSelecao(obj1 As Object)

Dim objCompetencias As New ClassCompetencias
Dim lErro As Long

On Error GoTo Erro_objEventoCompetencia_evSelecao

    Set objCompetencias = obj1

    Competencia.Text = objCompetencias.sNomeReduzido
    Call Competencia_Validate(bSGECancelDummy)

    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_objEventoCompetencia_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196012)

    End Select

    Exit Sub

End Sub

Private Sub objEventoMP_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim sProdutoMascarado As String
Dim iLinha As Integer
Dim objProdutoKit As New ClassProdutoKit

On Error GoTo Erro_objEventoMP_evSelecao

    Set objProduto = obj1
        
    lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 196013

    'Verifica se há algum MPProduto repetido no grid
    For iLinha = 1 To objGridMP.iLinhasExistentes
        If iLinha <> GridMP.Row Then
            If GridMP.TextMatrix(iLinha, iGrid_MPProduto_Col) = sProdutoMascarado Then
                MPProduto.PromptInclude = False
                MPProduto.Text = ""
                MPProduto.PromptInclude = True
                gError 196014
            End If
        End If
    Next
        
    MPProduto.PromptInclude = False
    MPProduto.Text = sProdutoMascarado
    MPProduto.PromptInclude = True
        
    If Not (Me.ActiveControl Is MPProduto) Then
        
        GridMP.TextMatrix(GridMP.Row, iGrid_MPProduto_Col) = MPProduto.Text
    
        GridMP.TextMatrix(GridMP.Row, iGrid_MPDesc_Col) = objProduto.sDescricao
        
        If objProduto.iCompras = PRODUTO_COMPRAVEL Then
            GridMP.TextMatrix(GridMP.Row, iGrid_MPOrigem_Col) = INSUMO_COMPRADO
        Else
            GridMP.TextMatrix(GridMP.Row, iGrid_MPOrigem_Col) = INSUMO_PRODUZIDO
        End If
        
        GridMP.TextMatrix(GridMP.Row, iGrid_MPUM_Col) = objProduto.sSiglaUMEstoque
        
        Call Carrega_ComboVersoes(objProduto.sCodigo)
        
        If MPVersao.ListCount > 0 Then
            Call MPVersao_SelecionaPadrao(objProduto.sCodigo)
            GridMP.TextMatrix(GridMP.Row, iGrid_MPVersao_Col) = MPVersao.Text
        End If
    
        Set objProdutoKit = New ClassProdutoKit
    
        objProdutoKit.sProdutoRaiz = GridMP.TextMatrix(GridMP.Row, iGrid_MPProduto_Col)
        objProdutoKit.sVersao = GridMP.TextMatrix(GridMP.Row, iGrid_MPVersao_Col)
        
        'Lê o MPProduto Raiz do Kit para pegar seus dados
        lErro = CF("ProdutoKit_Le_Raiz", objProdutoKit)
        If lErro <> SUCESSO And lErro <> 34875 Then gError 196015
        
        'Se não encontrou é porque não existe esta Versão do Kit
        If lErro = SUCESSO Then
                        
            MPComp.ListIndex = objProdutoKit.iComposicao
            GridMP.TextMatrix(GridMP.Row, iGrid_MPComp_Col) = MPComp.Text
        
        Else
        
            Call Composicao_Seleciona
            GridMP.TextMatrix(GridMP.Row, iGrid_MPComp_Col) = MPComp.Text
        
        End If
        
        'verifica se precisa preencher o grid com uma nova linha
        If GridMP.Row - GridMP.FixedRows = objGridMP.iLinhasExistentes Then
            objGridMP.iLinhasExistentes = objGridMP.iLinhasExistentes + 1
        End If
        
    End If
    
    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoMP_evSelecao:

    Select Case gErr

        Case 196013, 196015
        
        Case 196014
            Call Rotina_Erro(vbOKOnly, "ERRO_Produto_REPETIDO", gErr, sProdutoMascarado, iLinha)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196016)

    End Select

    Exit Sub

End Sub

Private Sub objEventoServico_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim sProduto As String

On Error GoTo Erro_objEventoServico_evSelecao

    Set objProduto = obj1

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then gError 196017

    Servico.PromptInclude = False
    Servico.Text = sProduto
    Servico.PromptInclude = True
    
    Call Servico_Validate(bSGECancelDummy)
    
    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoServico_evSelecao:

    Select Case gErr

        Case 196017
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196018)

    End Select

    Exit Sub

End Sub

Private Sub objEventoVersao_evSelecao(obj1 As Object)

Dim objKit As ClassKit
Dim lErro As Long

On Error GoTo Erro_objEventoVersao_evSelecao

    Set objKit = obj1

    Versao.Text = objKit.sVersao
        
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoVersao_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196019)

    End Select

    Exit Sub
    
End Sub

Private Sub Observacao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MPOrigem_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MPOrigem_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMP)
End Sub

Private Sub MPOrigem_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMP)
End Sub

Private Sub MPOrigem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMP.objControle = MPOrigem
    lErro = Grid_Campo_Libera_Foco(objGridMP)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Servico_GotFocus()
    Call MaskEdBox_TrataGotFocus(Servico, iAlterado)
End Sub

Private Sub Roteiro_NodeClick(ByVal Node As MSComctlLib.Node)

Dim lErro As Long
Dim objOperacoes As New ClassRoteiroSRVOper
Dim objNode As Node

On Error GoTo Erro_Roteiro_NodeClick

    Set objNode = Roteiro.SelectedItem

    Set objOperacoes = colComponentes.Item(objNode.Tag)
    
    lErro = Preenche_Operacoes(objOperacoes)
    If lErro <> SUCESSO Then gError 196020

    bOperacaoNova = False

    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_Roteiro_NodeClick:

    Select Case gErr

        Case 196020

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196021)

    End Select

    Exit Sub

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

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub
    
Public Sub Form_Activate()
    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)
End Sub
    
Public Sub Form_Deactivate()
    gi_ST_SetaIgnoraClick = 1
End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objEventoRSRV = Nothing
    Set objEventoCompet = Nothing
    Set objEventoCT = Nothing
    Set objEventoMP = Nothing
    Set objEventoVersao = Nothing
    Set objEventoServico = Nothing
    Set objEventoMO = Nothing
    Set objEventoMaq = Nothing
    
    Set objGridMP = Nothing
    Set objGridMO = Nothing
    Set objGridMaq = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196022)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    bOperacaoNova = True

    Set objEventoRSRV = New AdmEvento
    Set objEventoCompet = New AdmEvento
    Set objEventoCT = New AdmEvento
    Set objEventoMP = New AdmEvento
    Set objEventoVersao = New AdmEvento
    Set objEventoServico = New AdmEvento
    Set objEventoMO = New AdmEvento
    Set objEventoMaq = New AdmEvento
    
    DataCriacao.PromptInclude = False
    DataCriacao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataCriacao.PromptInclude = True
    
    DataUltModificacao.Caption = ""
    Autor.Caption = ""
    
    'Grid de Matéria Prima
    Set objGridMP = New AdmGrid
    
    'tela em questão
    Set objGridMP.objForm = Me
    
    lErro = Inicializa_GridMP(objGridMP)
    If lErro <> SUCESSO Then gError 196023
    
    'Grid de Mão de Obra
    Set objGridMO = New AdmGrid
    
    'tela em questão
    Set objGridMO.objForm = Me
    
    lErro = Inicializa_GridMO(objGridMO)
    If lErro <> SUCESSO Then gError 196024
    
    'Grid de Máquinas
    Set objGridMaq = New AdmGrid
    
    'tela em questão
    Set objGridMaq.objForm = Me
    
    lErro = Inicializa_GridMaq(objGridMaq)
    If lErro <> SUCESSO Then gError 196025

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", MPProduto)
    If lErro <> SUCESSO Then gError 196026
    
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Servico)
    If lErro <> SUCESSO Then gError 196027
    
    Quantidade.Format = FORMATO_ESTOQUE_KIT
    
    lErro_Chama_Tela = SUCESSO

    iFrameAtual = FRAME1_INDICE_IDENTIFICACAO
    iFrameAtualOper = FRAME2_INDICE_DETALHE
    iAlterado = 0
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 196023 To 196027
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196028)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objRoteiroSRV As ClassRoteiroSRV) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objRoteiroSRV Is Nothing) Then

        lErro = Traz_RoteiroSRV_Tela(objRoteiroSRV)
        If lErro <> SUCESSO And lErro <> 134974 Then gError 196029

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 196029

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196030)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objRoteiroSRV As ClassRoteiroSRV) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objOperacoes As New ClassRoteiroSRVOper

On Error GoTo Erro_Move_Tela_Memoria

    lErro = CF("Produto_Formata", Servico.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 196031
    
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
        objRoteiroSRV.sServico = sProdutoFormatado
    
    End If

    objRoteiroSRV.sVersao = Versao.Text
    objRoteiroSRV.sDescricao = Descricao.Text
    objRoteiroSRV.iDuracao = StrParaInt(Duracao.Text)
    
    objRoteiroSRV.dtDataCriacao = StrParaDate(DataCriacao.Text)

    objRoteiroSRV.dtDataUltModificacao = gdtDataAtual
    
    If Len(Trim(Autor.Caption)) = 0 Then
       objRoteiroSRV.sAutor = gsUsuario
    Else
       objRoteiroSRV.sAutor = Autor.Caption
    End If
    
    objRoteiroSRV.dQuantidade = StrParaDbl(Quantidade.Text)
    objRoteiroSRV.sUM = UM.Text
    
    'preenche a coleção das Operacoes
    For Each objOperacoes In colComponentes
       objRoteiroSRV.colOperacoes.Add objOperacoes
    Next
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 196031

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196032)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objRoteiroSRV As New ClassRoteiroSRV

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "RoteiroSRV"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objRoteiroSRV)
    If lErro <> SUCESSO Then gError 196033

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Servico", objRoteiroSRV.sServico, STRING_PRODUTO, "Servico"
    colCampoValor.Add "Versao", objRoteiroSRV.sVersao, STRING_KIT_VERSAO, "Versao"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 196033

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196034)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objRoteiroSRV As New ClassRoteiroSRV

On Error GoTo Erro_Tela_Preenche

    objRoteiroSRV.sServico = colCampoValor.Item("Servico").vValor
    objRoteiroSRV.sVersao = colCampoValor.Item("Versao").vValor

    If Len(Trim(objRoteiroSRV.sServico)) > 0 And Len(Trim(objRoteiroSRV.sVersao)) > 0 Then
    
        lErro = Traz_RoteiroSRV_Tela(objRoteiroSRV)
        If lErro <> SUCESSO And lErro <> 134974 Then gError 196035
        
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 196035

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196036)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objRoteiroSRV As New ClassRoteiroSRV

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se Servico está preenchido
    If Len(Trim(Servico.Text)) = 0 Then gError 196037
    
    'Verifica se Versao está preenchida
    If Len(Trim(Versao.Text)) = 0 Then gError 196038

    'Verifica se a Quantidade está preenchida
    If Len(Trim(Quantidade.Text)) = 0 Then gError 196039
    
    'Verifica se a U.M. está preenchida
    If Len(Trim(UM.Text)) = 0 Then gError 196040
    
    'Verifica se a Data de Criação está preenchida
    If Len(Trim(DataCriacao.ClipText)) = 0 Then gError 196041
    
    'Verifica se existe pelo menos uma Operação cadastrada
    If colComponentes.Count = 0 Then gError 196042

    'Preenche o objRoteiroSRV
    lErro = Move_Tela_Memoria(objRoteiroSRV)
    If lErro <> SUCESSO Then gError 196043
        
    lErro = Trata_Alteracao(objRoteiroSRV, objRoteiroSRV.sServico, objRoteiroSRV.sVersao)
    If lErro <> SUCESSO Then gError 196044
    
    'Grava o RoteiroSRV no Banco de Dados
    lErro = CF("RoteiroSRV_Grava", objRoteiroSRV)
    If lErro <> SUCESSO Then gError 196045
        
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 196037
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PREENCHIDO", gErr)

        Case 196038
            Call Rotina_Erro(vbOKOnly, "ERRO_VERSAO_ROTEIROSRV_NAO_PREENCHIDO", gErr)

        Case 196039
            Call Rotina_Erro(vbOKOnly, "ERRO_QTD_ROTFABR_NAO_PREENCHIDA", gErr)
        
        Case 196040
            Call Rotina_Erro(vbOKOnly, "ERRO_UM_ROTEIROSDEFABRICACAO_NAO_PREENCHIDA", gErr)

        Case 196041
            Call Rotina_Erro(vbOKOnly, "ERRO_DATACRIACAO_NAO_PREENCHIDA", gErr)

        Case 196042
            Call Rotina_Erro(vbOKOnly, "ERRO_OPERACOES_ROTEIROSDEFABRICACAO_NAO_PREENCHIDA", gErr)
        
        Case 196043, 196044, 196045
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196046)

    End Select

    Exit Function

End Function

Function Limpa_Tela_RoteiroSRV() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_RoteiroSRV
        
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    DescricaoProd.Caption = ""
    
    UM.Clear
    
    Call Composicao_Seleciona
    
    lErro = Limpa_Operacoes()
    If lErro <> SUCESSO Then gError 196047
    
    lErro = Limpa_Arvore_Roteiro()
    If lErro <> SUCESSO Then gError 196048

    DataCriacao.PromptInclude = False
    DataCriacao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataCriacao.PromptInclude = True
    
    DataUltModificacao.Caption = ""
    Autor.Caption = ""
    
    iAlterado = 0

    Limpa_Tela_RoteiroSRV = SUCESSO

    Exit Function

Erro_Limpa_Tela_RoteiroSRV:

    Limpa_Tela_RoteiroSRV = gErr

    Select Case gErr
    
        Case 196047, 196048
            'erro tratado nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196049)

    End Select

    Exit Function

End Function

Function Traz_RoteiroSRV_Tela(objRoteiroSRV As ClassRoteiroSRV) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sCodProduto As String
Dim sProdutoMascarado As String

On Error GoTo Erro_Traz_RoteiroSRV_Tela

    'Limpa Tela
    Call Limpa_Tela_RoteiroSRV
    
    sCodProduto = objRoteiroSRV.sServico
    
    lErro = Mascara_RetornaProdutoTela(sCodProduto, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 196050

    MPProduto.PromptInclude = False
    Servico.Text = sProdutoMascarado
    MPProduto.PromptInclude = True
        
    Versao.Text = objRoteiroSRV.sVersao
    
    'Lê o RoteiroSRV que está sendo Passado
    lErro = CF("RoteiroSRV_Le", objRoteiroSRV)
    If lErro <> SUCESSO And lErro <> 134617 Then gError 196051
    
    If lErro <> SUCESSO Then gError 196052

    Descricao.Text = objRoteiroSRV.sDescricao

    If objRoteiroSRV.dtDataCriacao <> DATA_NULA Then
        DataCriacao.PromptInclude = False
        DataCriacao.Text = Format(objRoteiroSRV.dtDataCriacao, "dd/mm/yy")
        DataCriacao.PromptInclude = True
    End If

    If objRoteiroSRV.dtDataUltModificacao <> DATA_NULA Then
        DataUltModificacao.Caption = Format(objRoteiroSRV.dtDataUltModificacao, "dd/mm/yyyy")
    End If
    
    If Len(objRoteiroSRV.sAutor) <> 0 Then
        Autor.Caption = objRoteiroSRV.sAutor
    End If
    
    If objRoteiroSRV.dQuantidade <> 0 Then Quantidade.Text = CStr(objRoteiroSRV.dQuantidade)
    UM.Text = objRoteiroSRV.sUM
    
    objProduto.sCodigo = objRoteiroSRV.sServico
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 196053
    
    DescricaoProd.Caption = objProduto.sDescricao

    lErro = CarregaComboUM(objProduto.iClasseUM, objRoteiroSRV.sUM)
    If lErro <> SUCESSO Then gError 196054

    lErro = Carrega_Arvore(objRoteiroSRV)
    If lErro <> SUCESSO Then gError 196055
    
    Duracao.PromptInclude = False
    Duracao.Text = CStr(objRoteiroSRV.iDuracao)
    Duracao.PromptInclude = True

    iAlterado = 0

    Traz_RoteiroSRV_Tela = SUCESSO

    Exit Function

Erro_Traz_RoteiroSRV_Tela:

    Traz_RoteiroSRV_Tela = gErr

    Select Case gErr

        Case 196050, 196051, 196053, 196054, 196055
            'erros tratados nas rotinas chamadas
        
        Case 196052
            'sem dados -> erro tratado na rotina chamadora
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196056)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 196057

    'Limpa Tela
    Call Limpa_Tela_RoteiroSRV

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 196057

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196058)

    End Select

    Exit Sub

End Sub

Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196059)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 196060

    Call Limpa_Tela_RoteiroSRV

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 196060

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196061)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objRoteiroSRV As New ClassRoteiroSRV
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Servico.Text)) = 0 Then gError 196062
    If Len(Trim(Versao.Text)) = 0 Then gError 196063

    lErro = CF("Produto_Formata", Servico.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 196064
    
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        objRoteiroSRV.sServico = sProdutoFormatado
    End If

    objRoteiroSRV.sVersao = Versao.Text

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_ROTEIROSRV", objRoteiroSRV.sServico, objRoteiroSRV.sVersao)

    If vbMsgRes = vbNo Then gError 196065

    'Exclui a requisição de consumo
    lErro = CF("RoteiroSRV_Exclui", objRoteiroSRV)
    If lErro <> SUCESSO Then gError 196066

    'Limpa Tela
    Call Limpa_Tela_RoteiroSRV

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 196062
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PREENCHIDO", gErr)

        Case 196063
            Call Rotina_Erro(vbOKOnly, "ERRO_VERSAO_ROTEIROSRV_NAO_PREENCHIDO", gErr)

        Case 196064 To 196066
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196067)

    End Select

    Exit Sub

End Sub

Private Sub Servico_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProduto As String
Dim sProdutoMascarado As String

On Error GoTo Erro_Produto_Validate
   
    If Len(Trim(Servico.ClipText)) <> 0 Then

        sProduto = Servico.Text
    
        'Critica o formato do MPProduto e se existe no BD
        lErro = CF("Produto_Critica", sProduto, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 Then gError 196068
    
        'se o Produto não estiver cadastrado ==> erro
        If lErro = 25041 Then gError 196069
    
        'se o Produto for gerencial, não pode fazer parte de um kit
        If objProduto.iGerencial = GERENCIAL Then gError 196073
        
        If objProduto.iNatureza <> NATUREZA_PROD_SERVICO Then gError 196070
        
        sProdutoMascarado = String(STRING_PRODUTO, 0)
    
        lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 196071
        
        Servico.PromptInclude = False
        Servico.Text = sProdutoMascarado
        Servico.PromptInclude = True
        
        lErro = CarregaComboUM(objProduto.iClasseUM, objProduto.sSiglaUMEstoque)
        If lErro <> SUCESSO Then gError 196072
    
        UM.Text = objProduto.sSiglaUMEstoque
    
        DescricaoProd.Caption = objProduto.sDescricao
        
    Else
        DescricaoProd.Caption = ""
    End If
            
    Exit Sub

Erro_Produto_Validate:

    Cancel = True

    Select Case gErr

        Case 196068, 196071, 196072
            'erros tratados nas rotinas chamadas
                   
        Case 196069
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case 196070
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_SERVICO", gErr, objProduto.sCodigo)

        Case 196073
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196074)

    End Select

    Exit Sub

End Sub

Private Sub Servico_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Versao_GotFocus()
    Call MaskEdBox_TrataGotFocus(Versao, iAlterado)
End Sub

Private Sub Versao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Descricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownDataCriacao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataCriacao_DownClick

    DataCriacao.SetFocus

    If Len(DataCriacao.ClipText) > 0 Then

        sData = DataCriacao.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 196075

        DataCriacao.Text = sData

    End If

    Exit Sub

Erro_UpDownDataCriacao_DownClick:

    Select Case gErr

        Case 196075

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196076)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataCriacao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataCriacao_UpClick

    DataCriacao.SetFocus

    If Len(Trim(DataCriacao.ClipText)) > 0 Then

        sData = DataCriacao.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 196077

        DataCriacao.Text = sData

    End If

    Exit Sub

Erro_UpDownDataCriacao_UpClick:

    Select Case gErr

        Case 196077

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196078)

    End Select

    Exit Sub

End Sub

Private Sub DataCriacao_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataCriacao, iAlterado)
End Sub

Private Sub DataCriacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataCriacao_Validate

    If Len(Trim(DataCriacao.ClipText)) <> 0 Then

        lErro = Data_Critica(DataCriacao.Text)
        If lErro <> SUCESSO Then gError 196079

    End If

    Exit Sub

Erro_DataCriacao_Validate:

    Cancel = True

    Select Case gErr

        Case 196079

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196080)

    End Select

    Exit Sub

End Sub

Private Sub DataCriacao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Quantidade_Validate

    'Veifica se Quantidade está preenchida
    If Len(Trim(Quantidade.Text)) <> 0 Then

       'Critica a Quantidade
       lErro = Valor_Positivo_Critica(Quantidade.Text)
       If lErro <> SUCESSO Then gError 196081
       
       Quantidade.Text = Formata_Estoque(Quantidade.Text)

    End If

    Exit Sub

Erro_Quantidade_Validate:

    Cancel = True

    Select Case gErr

        Case 196081

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196082)

    End Select

    Exit Sub

End Sub

Private Sub Quantidade_GotFocus()
    Call MaskEdBox_TrataGotFocus(Quantidade, iAlterado)
End Sub

Private Sub Quantidade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UM_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoRSRV_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRoteiroSRV As ClassRoteiroSRV

On Error GoTo Erro_objEventoRSRV_evSelecao

    Set objRoteiroSRV = obj1

    'Mostra os dados do RoteiroSRV na tela
    lErro = Traz_RoteiroSRV_Tela(objRoteiroSRV)
    If lErro <> SUCESSO And lErro <> 134974 Then gError 196083

    Me.Show

    Exit Sub

Erro_objEventoRSRV_evSelecao:

    Select Case gErr

        Case 196083

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196084)

    End Select

    Exit Sub

End Sub

Private Sub LabelProduto_Click()

Dim lErro As Long
Dim objKit As New ClassKit
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_LabelProduto_Click

    lErro = CF("Produto_Formata", Servico.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 196085

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""
    
    objProduto.sCodigo = sProdutoFormatado
    
    colSelecao.Add NATUREZA_PROD_SERVICO
    
    'Lista de produtos produzíveis
    Call Chama_Tela("ProdutoLista_Consulta", colSelecao, objProduto, objEventoServico, "Natureza = ?")
    
    Exit Sub

Erro_LabelProduto_Click:

    Select Case gErr

        Case 196085

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196086)

    End Select

    Exit Sub

End Sub

Private Sub LabelVersao_Click()

Dim lErro As Long
Dim objRoteiro As New ClassRoteiroSRV
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_LabelVersao_Click

    lErro = CF("Produto_Formata", Servico.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 196087

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 196088
    
    objRoteiro.sServico = sProdutoFormatado
    If Len(Trim(Versao.ClipText)) > 0 Then objRoteiro.sVersao = Versao.Text
        
    colSelecao.Add sProdutoFormatado
    
    Call Chama_Tela("RoteiroSRVLista", colSelecao, objRoteiro, objEventoRSRV, "Servico = ?")

    Exit Sub

Erro_LabelVersao_Click:

    Select Case gErr

        Case 196087
        
        Case 196088
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_SERVICO_NAO_PREENCHIDO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196089)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_GridMP(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Peça")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("C/P")
    objGrid.colColuna.Add ("UM")
    objGrid.colColuna.Add ("Qtd")
    objGrid.colColuna.Add ("Versão")
    objGrid.colColuna.Add ("Compos.")
    objGrid.colColuna.Add ("Obs")

    'Controles que participam do Grid
    objGrid.colCampo.Add (MPProduto.Name)
    objGrid.colCampo.Add (MPDesc.Name)
    objGrid.colCampo.Add (MPOrigem.Name)
    objGrid.colCampo.Add (MPUM.Name)
    objGrid.colCampo.Add (MPQtd.Name)
    objGrid.colCampo.Add (MPVersao.Name)
    objGrid.colCampo.Add (MPComp.Name)
    objGrid.colCampo.Add (MPObs.Name)

    'Colunas do Grid
    iGrid_MPProduto_Col = 1
    iGrid_MPDesc_Col = 2
    iGrid_MPOrigem_Col = 3
    iGrid_MPUM_Col = 4
    iGrid_MPQtd_Col = 5
    iGrid_MPVersao_Col = 6
    iGrid_MPComp_Col = 7
    iGrid_MPOBS_Col = 8

    objGrid.objGrid = GridMP

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_ITENS + 1 'NUM_MAX_ITENS_MOV_ESTOQUE

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 4

    'Largura da primeira coluna
    GridMP.ColWidth(0) = 250

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    Call Grid_Inicializa(objGrid)

    Inicializa_GridMP = SUCESSO

End Function

Private Sub GridMP_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridMP, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMP, iAlterado)
    End If

End Sub

Private Sub GridMP_GotFocus()
    
    Call Grid_Recebe_Foco(objGridMP)

End Sub

Private Sub GridMP_EnterCell()

    Call Grid_Entrada_Celula(objGridMP, iAlterado)

End Sub

Private Sub GridMP_LeaveCell()
    
    Call Saida_Celula(objGridMP)

End Sub

Private Sub GridMP_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridMP, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMP, iAlterado)
    End If

End Sub

Private Sub GridMP_RowColChange()

    Call Grid_RowColChange(objGridMP)

End Sub

Private Sub GridMP_Scroll()

    Call Grid_Scroll(objGridMP)

End Sub

Private Sub MPDesc_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MPDesc_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMP)

End Sub

Private Sub MPDesc_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMP)

End Sub

Private Sub MPDesc_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMP.objControle = MPDesc
    lErro = Grid_Campo_Libera_Foco(objGridMP)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MPQtd_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MPQtd_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMP)

End Sub

Private Sub MPQtd_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMP)

End Sub

Private Sub MPQtd_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMP.objControle = MPQtd
    lErro = Grid_Campo_Libera_Foco(objGridMP)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MPUM_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MPUM_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMP)

End Sub

Private Sub MPUM_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMP)

End Sub

Private Sub MPUM_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMP.objControle = MPUM
    lErro = Grid_Campo_Libera_Foco(objGridMP)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
        
        'OperacaoInsumos
        If objGridInt.objGrid.Name = GridMP.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col

                Case iGrid_MPProduto_Col

                    lErro = Saida_Celula_MPProduto(objGridInt)
                    If lErro <> SUCESSO Then gError 196090

                Case iGrid_MPDesc_Col

                    lErro = Saida_Celula_Padrao(objGridInt, MPDesc)
                    If lErro <> SUCESSO Then gError 196091

                Case iGrid_MPQtd_Col

                    lErro = Saida_Celula_MPQtd(objGridInt)
                    If lErro <> SUCESSO Then gError 196092

                Case iGrid_MPUM_Col

                    lErro = Saida_Celula_Padrao(objGridInt, MPUM)
                    If lErro <> SUCESSO Then gError 196093

                Case iGrid_MPVersao_Col

                    lErro = Saida_Celula_Padrao(objGridInt, MPVersao)
                    If lErro <> SUCESSO Then gError 196094
                
                Case iGrid_MPComp_Col

                    lErro = Saida_Celula_Padrao(objGridInt, MPComp)
                    If lErro <> SUCESSO Then gError 196095
                
                Case iGrid_MPOBS_Col

                    lErro = Saida_Celula_Padrao(objGridInt, MPObs)
                    If lErro <> SUCESSO Then gError 196096

            End Select
            
         ElseIf objGridInt.objGrid.Name = GridMO.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_MOCodigo_Col
                
                    lErro = Saida_Celula_MOCodigo(objGridInt)
                    If lErro <> SUCESSO Then gError 196097

                Case iGrid_MOQuantidade_Col
                
                    lErro = Saida_Celula_MOQuantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 196098

                Case iGrid_MOHoras_Col
                
                    lErro = Saida_Celula_MOHoras(objGridInt)
                    If lErro <> SUCESSO Then gError 196099

                Case iGrid_MOOBS_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, MOOBS)
                    If lErro <> SUCESSO Then gError 196100
        
            End Select
            
        ElseIf objGridInt.objGrid.Name = GridMaq.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_MaqCodigo_Col
                
                    lErro = Saida_Celula_MaqCodigo(objGridInt)
                    If lErro <> SUCESSO Then gError 196101

                Case iGrid_MaqQuantidade_Col
                
                    lErro = Saida_Celula_MaqQuantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 196102

                Case iGrid_MaqHoras_Col
                
                    lErro = Saida_Celula_MaqHoras(objGridInt)
                    If lErro <> SUCESSO Then gError 196103

                Case iGrid_MaqOBS_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, MaqOBS)
                    If lErro <> SUCESSO Then gError 196104
                    
            End Select
      
                    
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 196105

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 196090 To 196104

        Case 196105
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196106)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim iIndice As Integer
Dim objProdutos As New ClassProduto
Dim objClasseUM As New ClassClasseUM
Dim objUnidadeDeMedida As New ClassUnidadeDeMedida
Dim colSiglas As New Collection
Dim sUnidadeMed As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iMaquinaPreenchida As Integer
Dim iTipoMaoDeObraPreenchida As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    lErro = CF("Produto_Formata", GridMP.TextMatrix(GridMP.Row, iGrid_MPProduto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 196107
    
    If Len(Trim(GridMO.TextMatrix(GridMO.Row, iGrid_MOCodigo_Col))) > 0 Then
        iTipoMaoDeObraPreenchida = MARCADO
    Else
        iTipoMaoDeObraPreenchida = DESMARCADO
    End If
    
    If Len(Trim(GridMaq.TextMatrix(GridMaq.Row, iGrid_MaqCodigo_Col))) > 0 Then
        iMaquinaPreenchida = MARCADO
    Else
        iMaquinaPreenchida = DESMARCADO
    End If
    
    Select Case objControl.Name
    
        Case MPProduto.Name
        
            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
        
        Case MPOrigem.Name

            objControl.Enabled = False
        
        Case MPQtd.Name, MPDesc.Name, MPComp.Name, MPObs.Name
        
            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
        
        Case MPUM.Name
        
            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
                objControl.Enabled = True
    
                Set objProdutos = New ClassProduto
    
                objProdutos.sCodigo = sProdutoFormatado
    
                lErro = CF("Produto_Le", objProdutos)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 196108
    
                Set objClasseUM = New ClassClasseUM
                
                objClasseUM.iClasse = objProdutos.iClasseUM
    
                'Preenche a List da Combo UnidadeMed com as UM's do MPProduto
                lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                If lErro <> SUCESSO Then gError 196109
    
                'Se tem algum valor para MPUM do Grid
                If Len(GridMP.TextMatrix(GridMP.Row, iGrid_MPUM_Col)) > 0 Then
                    'Guardo o valor da MPUM da Linha
                    sUnidadeMed = GridMP.TextMatrix(GridMP.Row, iGrid_MPUM_Col)
                Else
                    'Senão coloco o do MPProduto em estoque
                    sUnidadeMed = objProdutos.sSiglaUMEstoque
                End If
                
                'Limpar as Unidades utilizadas anteriormente
                MPUM.Clear
    
                For Each objUnidadeDeMedida In colSiglas
                    MPUM.AddItem objUnidadeDeMedida.sSigla
                Next
    
                'MPUM.AddItem ""
    
                'Tento selecionar na Combo a Unidade anterior
                If MPUM.ListCount <> 0 Then
    
                    For iIndice = 0 To MPUM.ListCount - 1
    
                        If MPUM.List(iIndice) = sUnidadeMed Then
                            MPUM.ListIndex = iIndice
                            Exit For
                        End If
                    Next
                End If
                
            Else
                objControl.Enabled = False
            End If
        
        Case MPVersao.Name
        
            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                objControl.Enabled = True
                Call Carrega_ComboVersoes(sProdutoFormatado)
                Call MPVersao_Seleciona(GridMP.TextMatrix(GridMP.Row, iGrid_MPVersao_Col))
            Else
                objControl.Enabled = False
            End If
            
        Case MOCodigo.Name
        
            If iTipoMaoDeObraPreenchida = MARCADO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case MOQuantidade.Name, MOHoras.Name, MOOBS.Name

            If iTipoMaoDeObraPreenchida <> MARCADO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case MaqCodigo.Name
        
            If iMaquinaPreenchida = MARCADO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case MaqQuantidade.Name, MaqHoras.Name, MaqOBS.Name

            If iMaquinaPreenchida <> MARCADO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case Else
            objControl.Enabled = False
    
    End Select
        
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr
    
        Case 196107 To 196109

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 196110)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_MPProduto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sCodProduto As String
Dim iLinha As Integer
Dim objProdutos As ClassProduto
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String
Dim iProdutoPreenchido As Integer
Dim objProdutoKit As ClassProdutoKit

On Error GoTo Erro_Saida_Celula_MPProduto

    Set objGridInt.objControle = MPProduto
                
    sCodProduto = MPProduto.Text
        
    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 196111
    
    'Se o campo foi preenchido
    If Len(sProdutoFormatado) > 0 Then

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        lErro = Mascara_RetornaProdutoTela(sProdutoFormatado, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 196112
        
        MPProduto.PromptInclude = False
        MPProduto.Text = sProdutoMascarado
        MPProduto.PromptInclude = True
        
        'Verifica se há algum MPProduto repetido no grid
        For iLinha = 1 To objGridInt.iLinhasExistentes
            If iLinha <> GridMP.Row Then
                If GridMP.TextMatrix(iLinha, iGrid_MPProduto_Col) = sProdutoMascarado Then
                    MPProduto.PromptInclude = False
                    MPProduto.Text = ""
                    MPProduto.PromptInclude = True
                    gError 196113
                End If
            End If
        Next
        
        Set objProdutos = New ClassProduto

        objProdutos.sCodigo = sProdutoFormatado

        lErro = CF("Produto_Le", objProdutos)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 196114
        
        'Verifica se o MPProduto pode compor um Kit
        If objProdutos.iAtivo <> 0 And objProdutos.iGerencial <> 0 And _
            objProdutos.iKitBasico <> 1 And objProdutos.iKitInt <> 1 Then gError 196115
                
        GridMP.TextMatrix(GridMP.Row, iGrid_MPDesc_Col) = objProdutos.sDescricao
        
        If objProdutos.iCompras = PRODUTO_COMPRAVEL Then
            GridMP.TextMatrix(GridMP.Row, iGrid_MPOrigem_Col) = INSUMO_COMPRADO
        Else
            GridMP.TextMatrix(GridMP.Row, iGrid_MPOrigem_Col) = INSUMO_PRODUZIDO
        End If

        If Len(GridMP.TextMatrix(GridMP.Row, iGrid_MPUM_Col)) = 0 Then
            GridMP.TextMatrix(GridMP.Row, iGrid_MPUM_Col) = objProdutos.sSiglaUMEstoque
        End If

        Call Carrega_ComboVersoes(objProdutos.sCodigo)

        If MPVersao.ListCount > 0 Then

            If Len(GridMP.TextMatrix(GridMP.Row, iGrid_MPVersao_Col)) = 0 Then
                Call MPVersao_SelecionaPadrao(objProdutos.sCodigo)
                GridMP.TextMatrix(GridMP.Row, iGrid_MPVersao_Col) = MPVersao.Text
            End If

        End If

        Set objProdutoKit = New ClassProdutoKit

        objProdutoKit.sProdutoRaiz = GridMP.TextMatrix(GridMP.Row, iGrid_MPProduto_Col)
        objProdutoKit.sVersao = GridMP.TextMatrix(GridMP.Row, iGrid_MPVersao_Col)

        'Lê o MPProduto Raiz do Kit para pegar seus dados
        lErro = CF("ProdutoKit_Le_Raiz", objProdutoKit)
        If lErro <> SUCESSO And lErro <> 34875 Then gError 196116

        'Se não encontrou é porque não existe esta Versão do Kit
        If lErro = SUCESSO Then

            If Len(GridMP.TextMatrix(GridMP.Row, iGrid_MPComp_Col)) = 0 Then
                Call Combo_Seleciona_ItemData(MPComp, objProdutoKit.iComposicao)
                GridMP.TextMatrix(GridMP.Row, iGrid_MPComp_Col) = MPComp.Text
            End If

        Else

            If Len(GridMP.TextMatrix(GridMP.Row, iGrid_MPComp_Col)) = 0 Then
                Call Composicao_Seleciona
                GridMP.TextMatrix(GridMP.Row, iGrid_MPComp_Col) = MPComp.Text
            End If

        End If

        'verifica se precisa preencher o grid com uma nova linha
        If GridMP.Row - GridMP.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 196117

    Saida_Celula_MPProduto = SUCESSO

    Exit Function

Erro_Saida_Celula_MPProduto:

    Saida_Celula_MPProduto = gErr

    Select Case gErr

        Case 196111, 196112, 196114, 196116, 196117
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 196113
            Call Rotina_Erro(vbOKOnly, "ERRO_Produto_REPETIDO", gErr, sProdutoMascarado, iLinha)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 196115
            Call Rotina_Erro(vbOKOnly, "ERRO_Produto_NAO_PODE_COMPOR_KIT", gErr, sProdutoMascarado, iLinha)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 196118)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MPQtd(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MPQtd

    Set objGridInt.objControle = MPQtd
    
    'Se o campo foi preenchido
    If Len(Trim(MPQtd.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(MPQtd.Text)
        If lErro <> SUCESSO Then gError 196119
        
        MPQtd.Text = Formata_Estoque(MPQtd.Text)
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 196120

    Saida_Celula_MPQtd = SUCESSO

    Exit Function

Erro_Saida_Celula_MPQtd:

    Saida_Celula_MPQtd = gErr

    Select Case gErr

        Case 196119, 196120
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 196121)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MPUM(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MPUM

    Set objGridInt.objControle = MPUM
    
    'Se o campo foi preenchido
    If Len(Trim(MPUM.Text)) > 0 Then
    
        GridMP.TextMatrix(GridMP.Row, iGrid_MPUM_Col) = MPUM.Text
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 196122

    Saida_Celula_MPUM = SUCESSO

    Exit Function

Erro_Saida_Celula_MPUM:

    Saida_Celula_MPUM = gErr

    Select Case gErr

        Case 196122
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 196123)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function CarregaComboUM(ByVal iClasseUM As Integer, ByVal sUM As String) As Long

Dim lErro As Long
Dim objClasseUM As ClassClasseUM
Dim objUnidadeDeMedida As ClassUnidadeDeMedida
Dim colSiglas As New Collection
Dim sUnidadeMed As String
Dim iIndice As Integer

On Error GoTo Erro_CarregaComboUM

    Set objClasseUM = New ClassClasseUM
    
    objClasseUM.iClasse = iClasseUM
    
    'Preenche a List da Combo UnidadeMed com as UM's da Competencia
    lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
    If lErro <> SUCESSO Then gError 196124

    'Se tem algum valor para UM na Tela
    If Len(UM.Text) > 0 Then
        'Guardo o valor da UM da Tela
        sUnidadeMed = UM.Text
    Else
        'Senão coloco a do Servico no Kit
        sUnidadeMed = sUM
    End If
    
    'Limpar as Unidades utilizadas anteriormente
    UM.Clear

    For Each objUnidadeDeMedida In colSiglas
        UM.AddItem objUnidadeDeMedida.sSigla
    Next

    'Tento selecionar na Combo a Unidade anterior
    If UM.ListCount <> 0 Then

        For iIndice = 0 To UM.ListCount - 1

            If UM.List(iIndice) = sUnidadeMed Then
                UM.ListIndex = iIndice
                Exit For
            End If
        Next
    End If
    
    CarregaComboUM = SUCESSO
    
    Exit Function

Erro_CarregaComboUM:

    CarregaComboUM = gErr

    Select Case gErr

        Case 196124
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196125)

    End Select

    Exit Function

End Function

Private Function Preenche_GridMP(objOperacoes As ClassRoteiroSRVOper) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objProdutos As ClassProduto
Dim sProdutoMascarado As String

On Error GoTo Erro_Preenche_GridMP
    
    Call Grid_Limpa(objGridMP)
    
    'Exibe os dados da coleção na tela
    For iIndice = 1 To objOperacoes.colMP.Count
        
        Set objProdutos = New ClassProduto
        
        objProdutos.sCodigo = objOperacoes.colMP.Item(iIndice).sProduto
        
        lErro = CF("Produto_Le", objProdutos)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 196126
        
        lErro = Mascara_RetornaProdutoTela(objProdutos.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 196127
                
        'Insere no GridMP
        GridMP.TextMatrix(iIndice, iGrid_MPProduto_Col) = sProdutoMascarado
        GridMP.TextMatrix(iIndice, iGrid_MPDesc_Col) = objProdutos.sDescricao
        
        If objProdutos.iCompras = PRODUTO_COMPRAVEL Then
            GridMP.TextMatrix(iIndice, iGrid_MPOrigem_Col) = INSUMO_COMPRADO
        Else
            GridMP.TextMatrix(iIndice, iGrid_MPOrigem_Col) = INSUMO_PRODUZIDO
        End If
        
        If objOperacoes.colMP.Item(iIndice).dQuantidade > 0 Then
            GridMP.TextMatrix(iIndice, iGrid_MPQtd_Col) = Formata_Estoque(objOperacoes.colMP.Item(iIndice).dQuantidade)
        End If
        GridMP.TextMatrix(iIndice, iGrid_MPUM_Col) = objOperacoes.colMP.Item(iIndice).sUM
        GridMP.TextMatrix(iIndice, iGrid_MPVersao_Col) = objOperacoes.colMP.Item(iIndice).sVersao
        Call Combo_Seleciona_ItemData(MPComp, objOperacoes.colMP.Item(iIndice).iComposicao)
        GridMP.TextMatrix(iIndice, iGrid_MPComp_Col) = MPComp.Text
        GridMP.TextMatrix(iIndice, iGrid_MPOBS_Col) = objOperacoes.colMP.Item(iIndice).sObs
        
    Next

    objGridMP.iLinhasExistentes = objOperacoes.colMP.Count
    
    Preenche_GridMP = SUCESSO
    
    Exit Function

Erro_Preenche_GridMP:

    Preenche_GridMP = gErr

    Select Case gErr

        Case 196126, 196127

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196128)

    End Select

    Exit Function

End Function

Private Function Move_OperPecas_Memoria(objOperacoes As ClassRoteiroSRVOper) As Long

Dim lErro As Long
Dim objRotSRVOperMP As ClassRoteiroSRVOperMP
Dim objProdutos As ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iIndice As Integer
Dim objKit As ClassKit

On Error GoTo Erro_Move_OperPecas_Memoria

    'Ir preenchendo a colecao no objOperacoes com todas as linhas "existentes" do grid
    For iIndice = 1 To objGridMP.iLinhasExistentes

        If Len(Trim(GridMP.TextMatrix(iIndice, iGrid_MPProduto_Col))) <> 0 Then
        
            'Verifica se Quantidade está preenchida
            If Len(Trim(GridMP.TextMatrix(iIndice, iGrid_MPQtd_Col))) = 0 Then gError 196129
    
            'Verifica se MPUM está preenchida
            If Len(Trim(GridMP.TextMatrix(iIndice, iGrid_MPUM_Col))) = 0 Then gError 196130
    
            'Verifica se Composicao está preenchida
            If Len(Trim(GridMP.TextMatrix(iIndice, iGrid_MPComp_Col))) = 0 Then gError 196131
            
            Set objProdutos = New ClassProduto
            
            lErro = CF("Produto_Formata", GridMP.TextMatrix(iIndice, iGrid_MPProduto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 196132
            
            objProdutos.sCodigo = sProdutoFormatado
            
            lErro = CF("Produto_Le", objProdutos)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 196133
            
            Set objRotSRVOperMP = New ClassRoteiroSRVOperMP
            
            objRotSRVOperMP.lNumIntDocOper = objOperacoes.lNumIntDoc
            objRotSRVOperMP.sProduto = objProdutos.sCodigo
            objRotSRVOperMP.dQuantidade = StrParaDbl(GridMP.TextMatrix(iIndice, iGrid_MPQtd_Col))
            objRotSRVOperMP.sUM = GridMP.TextMatrix(iIndice, iGrid_MPUM_Col)
            objRotSRVOperMP.iComposicao = Composicao_Extrai(GridMP.TextMatrix(iIndice, iGrid_MPComp_Col))
            objRotSRVOperMP.sObs = GridMP.TextMatrix(iIndice, iGrid_MPOBS_Col)
            objRotSRVOperMP.sVersao = GridMP.TextMatrix(iIndice, iGrid_MPVersao_Col)
            
            If objProdutos.iCompras = PRODUTO_PRODUZIVEL Then
            
                If Len(objRotSRVOperMP.sVersao) = 0 Then
                
                    Set objKit = New ClassKit
            
                    objKit.sProdutoRaiz = objProdutos.sCodigo
                
                    'Le as Versoes Ativas e a Padrao
                    lErro = CF("Kit_Le_Padrao", objKit)
                    If lErro <> SUCESSO And lErro <> 106304 Then gError 196134
                    
                    If lErro <> SUCESSO Then gError 196135
                    
                    objRotSRVOperMP.sVersao = objKit.sVersao
                
                End If
            
            End If
            
            objOperacoes.colMP.Add objRotSRVOperMP
            
        End If
    
    Next

    Move_OperPecas_Memoria = SUCESSO

    Exit Function

Erro_Move_OperPecas_Memoria:

    Move_OperPecas_Memoria = gErr

    Select Case gErr
    
        Case 196129
            Call Rotina_Erro(vbOKOnly, "ERRO_OPERACAO_QUANTIDADE_NAO_PREENCHIDA", gErr)
        
        Case 196130
            Call Rotina_Erro(vbOKOnly, "ERRO_OPERACAO_UMPRODUTO_NAO_PREENCHIDA", gErr)
        
        Case 196131
            Call Rotina_Erro(vbOKOnly, "ERRO_OPERACAO_COMPOSICAO_NAO_PREENCHIDA", gErr)
        
        Case 196132, 196133, 196134
            'erros tratados nas rotinas chamadas
            
        Case 196135
            Call Rotina_Erro(vbOKOnly, "ERRO_VERSAO_PADRAO_NAO_LOCALIZADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196136)

    End Select

    Exit Function

End Function

Private Sub TabStrip2_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip2.SelectedItem.Index <> iFrameAtualOper Then

        If TabStrip_PodeTrocarTab(iFrameAtualOper, TabStrip2, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame2(TabStrip2.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame2(iFrameAtualOper).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtualOper = TabStrip2.SelectedItem.Index
        
    End If

End Sub

Private Sub Composicao_Seleciona()

Dim iIndice As Integer

    MPComp.ListIndex = -1
    For iIndice = 0 To MPComp.ListCount - 1
        If MPComp.List(iIndice) = COMPOSICAO_VARIAVEL Then
            MPComp.ListIndex = iIndice
            Exit For
        End If
    Next

End Sub

Private Function Composicao_Extrai(sComposicaoGrid As String) As Integer

Dim iIndice As Integer

    For iIndice = 0 To MPComp.ListCount - 1
        If MPComp.List(iIndice) = sComposicaoGrid Then
            Composicao_Extrai = iIndice
            Exit For
        End If
    Next

End Function

Private Sub MPVersao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MPVersao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMP)

End Sub

Private Sub MPVersao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMP)

End Sub

Private Sub MPVersao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMP.objControle = MPVersao
    lErro = Grid_Campo_Libera_Foco(objGridMP)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MPComp_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MPComp_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMP)

End Sub

Private Sub MPComp_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMP)

End Sub

Private Sub MPComp_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMP.objControle = MPComp
    lErro = Grid_Campo_Libera_Foco(objGridMP)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MPProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MPProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMP)

End Sub

Private Sub MPProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMP)

End Sub

Private Sub MPProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMP.objControle = MPProduto
    lErro = Grid_Campo_Libera_Foco(objGridMP)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Sub Calcula_Proxima_Chave(sChave As String)

Dim iNumero As Integer

    iNumero = iProxChave
    iProxChave = iProxChave + 1
    sChave = "X" & right$(CStr(100000 + iNumero), 5)

End Sub

Sub Recalcula_Nivel_Sequencial()
'(re)calcula niveis e sequencias de toda a estrutura
'deve ser chamada apos a remocao de algum node

Dim iIndice As Integer

    If Roteiro.Nodes.Count = 0 Then Exit Sub

    For iIndice = LBound(aNivelSequencial) To UBound(aNivelSequencial)
        aNivelSequencial(iIndice) = 0
    Next

    iUltimoNivel = 0

    'chamar rotina que recalcula recursivamente os campos nivel e sequencial (Nivel e SeqArvore)
    Call Calcula_Nivel_Sequencial(Roteiro.Nodes.Item(1), 0, 0)

End Sub

Sub Calcula_Nivel_Sequencial(objNode As Node, iNivel As Integer, iPosicaoAtual As Integer)
'parte recursiva do recalculo de nivel e sequencial, atuando a partir do node passado
'iNivel informa o nivel deste node

Dim objOperacoes As New ClassRoteiroSRVOper
Dim sChave1 As String

    sChave1 = objNode.Tag

    Set objOperacoes = colComponentes.Item(sChave1)

    aNivelSequencial(iNivel) = aNivelSequencial(iNivel) + 1

    iPosicaoAtual = iPosicaoAtual + 1
    aSeqPai(iNivel) = iPosicaoAtual

    objOperacoes.iSeqArvore = aNivelSequencial(iNivel)

    If iNivel > 0 Then
        objOperacoes.iSeqPai = aSeqPai(iNivel - 1)
    Else
        objOperacoes.iSeqPai = 0
    End If
    
    objOperacoes.iSeq = iPosicaoAtual

    objOperacoes.iNivel = iNivel
    
    colComponentes.Remove sChave1
    colComponentes.Add objOperacoes, sChave1

    If objNode.Children > 0 Then
        Call Calcula_Nivel_Sequencial(objNode.Child, iNivel + 1, iPosicaoAtual)
    End If

    If objNode.Index <> objNode.LastSibling.Index Then Call Calcula_Nivel_Sequencial(objNode.Next, iNivel, iPosicaoAtual)

    If iNivel > iUltimoNivel Then iUltimoNivel = iNivel
   
End Sub

Function Limpa_Operacoes() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Operacoes
        
    Nivel.Caption = ""
    Sequencial.Caption = ""
    
    Competencia.PromptInclude = False
    Competencia.Text = ""
    Competencia.PromptInclude = True
    
    DescricaoCompetencia.Caption = ""
    
    CTPadrao.PromptInclude = False
    CTPadrao.Text = ""
    CTPadrao.PromptInclude = True
    
    DescricaoCTPadrao.Caption = ""
    
    bOperacaoNova = True
    
    Observacao.Text = ""
    
    Call Grid_Limpa(objGridMP)
    Call Grid_Limpa(objGridMaq)
    Call Grid_Limpa(objGridMO)
    
    iAlterado = 0

    Limpa_Operacoes = SUCESSO

    Exit Function

Erro_Limpa_Operacoes:

    Limpa_Operacoes = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196137)

    End Select

    Exit Function

End Function

Private Function Move_Operacoes_Memoria(ByVal objOperacoes As ClassRoteiroSRVOper, ByVal objCompetencias As ClassCompetencias, ByVal objCentrodeTrabalho As ClassCentrodeTrabalho) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Operacoes_Memoria
        
    objCompetencias.sNomeReduzido = Competencia.Text
    
    'Verifica a Competencia no BD a partir do Código
    lErro = CF("Competencias_Le_NomeReduzido", objCompetencias)
    If lErro <> SUCESSO And lErro <> 134937 Then gError 196138

    objOperacoes.lNumIntDocCompet = objCompetencias.lNumIntDoc
    
    If Len(Trim(CTPadrao.Text)) <> 0 Then
            
        objCentrodeTrabalho.sNomeReduzido = CTPadrao.Text
        
        'Lê o CentrodeTrabalho que está sendo Passado
        lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134941 Then gError 196139
        
        objOperacoes.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc
    
    End If
    
    If Len(Trim(Observacao.Text)) <> 0 Then objOperacoes.sObservacao = Observacao.Text
        
    lErro = Move_OperPecas_Memoria(objOperacoes)
    If lErro <> SUCESSO Then gError 196140
    
    lErro = Move_OperMaq_Memoria(objOperacoes)
    If lErro <> SUCESSO Then gError 196141
    
    lErro = Move_OperMO_Memoria(objOperacoes)
    If lErro <> SUCESSO Then gError 196142

    Move_Operacoes_Memoria = SUCESSO

    Exit Function

Erro_Move_Operacoes_Memoria:

    Move_Operacoes_Memoria = gErr

    Select Case gErr

        Case 196138 To 196142
            'erros tratados nas rotinas chamadas
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196143)

    End Select

    Exit Function

End Function

Sub Remove_Filhos(objNode As Node)
'remove objNode, seus irmaos e filhos de colComponentes

        colComponentes.Remove (objNode.Tag)

        If objNode.Children > 0 Then

            Call Remove_Filhos(objNode.Child)

        End If

        If objNode <> objNode.LastSibling Then Call Remove_Filhos(objNode.Next)

    Exit Sub

End Sub

Function Preenche_Operacoes(objOperacoes As ClassRoteiroSRVOper) As Long
'preenche as tabs de Detalhes, Insumos e Produção à partir dos dados de objOperacoes

Dim lErro As Long
Dim objCompetencias As ClassCompetencias
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_Preenche_Operacoes

    'Limpa as Tabs de Detalhes, Insumos e Produção
    lErro = Limpa_Operacoes()
    If lErro <> SUCESSO Then gError 196144

    Nivel.Caption = objOperacoes.iNivel
    Sequencial.Caption = objOperacoes.iSeqArvore
    
    Set objCompetencias = New ClassCompetencias
    
    objCompetencias.lNumIntDoc = objOperacoes.lNumIntDocCompet
    
    lErro = CF("Competencias_Le_NumIntDoc", objCompetencias)
    If lErro <> SUCESSO And lErro <> 134336 Then gError 196145
    
    Competencia.PromptInclude = False
    Competencia.Text = objCompetencias.sNomeReduzido
    Competencia.PromptInclude = True
    
    DescricaoCompetencia.Caption = objCompetencias.sDescricao
    
    If objOperacoes.lNumIntDocCT <> 0 Then
        
        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
        
        objCentrodeTrabalho.lNumIntDoc = objOperacoes.lNumIntDocCT
        
        lErro = CF("CentroDeTrabalho_Le_NumIntDoc", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134590 Then gError 196146
        
        CTPadrao.PromptInclude = False
        CTPadrao.Text = objCentrodeTrabalho.sNomeReduzido
        CTPadrao.PromptInclude = True
        
        DescricaoCTPadrao.Caption = objCentrodeTrabalho.sDescricao
    
    End If
    
    Observacao.Text = objOperacoes.sObservacao

    lErro = Preenche_GridMP(objOperacoes)
    If lErro <> SUCESSO Then gError 196147

    lErro = Preenche_GridMaq(objOperacoes)
    If lErro <> SUCESSO Then gError 196148

    lErro = Preenche_GridMO(objOperacoes)
    If lErro <> SUCESSO Then gError 196149

    iAlterado = 0

    Preenche_Operacoes = SUCESSO

    Exit Function

Erro_Preenche_Operacoes:

    Preenche_Operacoes = gErr

    Select Case gErr
    
        Case 196144 To 196149

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196150)

    End Select

    Exit Function

End Function

Private Sub Roteiro_Collapse(ByVal Node As MSComctlLib.Node)
    Roteiro_NodeClick Node
End Sub

Private Sub Carrega_ComboVersoes(ByVal sServico As String)
    
Dim lErro As Long
Dim objKit As New ClassKit
Dim colKits As New Collection
    
On Error GoTo Erro_Carrega_ComboVersoes
    
    MPVersao.Enabled = True
    
    'Limpa a Combo
    MPVersao.Clear
    
    'Armazena o MPProduto Raiz do kit
    objKit.sProdutoRaiz = sServico
    
    'Le as Versoes Ativas e a Padrao
    lErro = CF("Kit_Le_Produziveis", objKit, colKits)
    If lErro <> SUCESSO And lErro <> 106333 Then gError 196151
    
    MPVersao.AddItem ""
    
    'Carrega a Combo com os Dados da Colecao
    For Each objKit In colKits
    
        MPVersao.AddItem (objKit.sVersao)
        
    Next
    
    Exit Sub
    
Erro_Carrega_ComboVersoes:

    Select Case gErr
    
        Case 196151
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196152)
    
    End Select
    
End Sub

Private Sub MPVersao_Seleciona(sVersao As String)
Dim iIndice As Integer

    MPVersao.ListIndex = -1
    For iIndice = 0 To MPVersao.ListCount - 1
        If MPVersao.List(iIndice) = sVersao Then
            MPVersao.ListIndex = iIndice
            Exit For
        End If
    Next

End Sub

Private Function MPVersao_SelecionaPadrao(sProduto As String)

Dim lErro As Long
Dim objKit As New ClassKit
    
On Error GoTo Erro_MPVersao_SelecionaPadrao
    
    'Armazena o MPProduto Raiz do kit
    objKit.sProdutoRaiz = sProduto
    
    'Le as Versoes Ativas e a Padrao
    lErro = CF("Kit_Le_Padrao", objKit)
    If lErro <> SUCESSO And lErro <> 106304 Then gError 196153
        
    Call MPVersao_Seleciona(objKit.sVersao)
    
    MPVersao_SelecionaPadrao = SUCESSO
    
    Exit Function

Erro_MPVersao_SelecionaPadrao:

    MPVersao_SelecionaPadrao = gErr
    
    Select Case gErr

        Case 196153
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196154)

    End Select

    Exit Function
    
End Function

Function Limpa_Arvore_Roteiro() As Long
'Limpa a Arvore do Roteiro

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Arvore_Roteiro

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then gError 196155

    Roteiro.Nodes.Clear
    Set colComponentes = New Collection
    
    iProxChave = 1

    Limpa_Arvore_Roteiro = SUCESSO

    Exit Function

Erro_Limpa_Arvore_Roteiro:

    Limpa_Arvore_Roteiro = gErr
    
    Select Case gErr

        Case 196155

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196156)

    End Select

    Exit Function

End Function

Function Carrega_Arvore(objRoteiroSRV As ClassRoteiroSRV) As Long
'preenche a treeview Roteiro com a composicao de objRoteiroSRV
   
Dim objNode As Node
Dim lErro As Long, sChave As String, sChaveTvw As String
Dim iIndice As Integer
Dim sTexto As String
Dim objOperacoes As ClassRoteiroSRVOper
Dim objCompetencias As ClassCompetencias
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_Carrega_Arvore
    
    'Critica o formato do MPProduto e se existe no BD
    lErro = CF("Produto_Critica", Servico.Text, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 25041 Then gError 196157
    
    For Each objOperacoes In objRoteiroSRV.colOperacoes

        Set objCompetencias = New ClassCompetencias
        
        objCompetencias.lNumIntDoc = objOperacoes.lNumIntDocCompet
        
        lErro = CF("Competencias_Le_NumIntDoc", objCompetencias)
        If lErro <> SUCESSO And lErro <> 134336 Then gError 196158

        'prepara texto que identificará a nova Operação que está sendo incluida
        sTexto = objCompetencias.sNomeReduzido
        
        sTexto = sTexto & " (" & objProduto.sNomeReduzido
        
        If objOperacoes.lNumIntDocCT > 0 Then
        
            Set objCentrodeTrabalho = New ClassCentrodeTrabalho
            
            objCentrodeTrabalho.lNumIntDoc = objOperacoes.lNumIntDocCT
            
            lErro = CF("CentroDeTrabalho_Le_NumIntDoc", objCentrodeTrabalho)
            If lErro <> SUCESSO And lErro <> 134590 Then gError 196159
        
            sTexto = sTexto & " - " & objCentrodeTrabalho.sNomeReduzido
           
        End If
        
        sTexto = sTexto & ")"
        
        'prepara uma chave para relacionar colComponentes ao node que está sendo incluido
        Call Calcula_Proxima_Chave(sChaveTvw)
        
        sChave = sChaveTvw
        sChaveTvw = sChaveTvw & objCompetencias.lCodigo

        If objOperacoes.iNivel = 0 Then

            Set objNode = Roteiro.Nodes.Add(, tvwFirst, sChaveTvw, sTexto)

        Else

            Set objNode = Roteiro.Nodes.Add(objOperacoes.iSeqPai, tvwChild, sChaveTvw, sTexto)

        End If
                
        Roteiro.Nodes.Item(objNode.Index).Expanded = True
        
        colComponentes.Add objOperacoes, sChave
        
        objNode.Tag = sChave
        
    Next

    'se houver árvore ...
    If Roteiro.Nodes.Count > 0 Then
        
        'selecionar a raiz
        Set Roteiro.SelectedItem = Roteiro.Nodes.Item(1)
        Roteiro.SelectedItem.Selected = True
        
        'e carregar as operações pertinentes
        Call Roteiro_NodeClick(Roteiro.Nodes.Item(1))
        
        bOperacaoNova = False
        
    End If

    Carrega_Arvore = SUCESSO

    Exit Function

Erro_Carrega_Arvore:

    Carrega_Arvore = gErr

    Select Case gErr

        Case 196157, 196158, 196159
            'erro tratado nas rotinas chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196160)

    End Select

    Exit Function

End Function

Private Sub BotaoRelRoteiro_Click()

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim lNumIntRel As Long
Dim objRoteiro As New ClassRoteiroSRV
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sTsk As String

On Error GoTo Erro_BotaoRelRoteiro_Click

    lErro = CF("Produto_Formata", Servico.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 196161

    If iProdutoPreenchido = PRODUTO_PREENCHIDO And Len(Trim(Versao.Text)) <> 0 Then

        objRoteiro.sServico = sProdutoFormatado
        objRoteiro.sVersao = Versao.Text

        lErro = CF("RelRoteiroServico_Prepara", objRoteiro, lNumIntRel)
        If lErro <> SUCESSO Then gError 196162
        
        If DetalharInsumos.Value = vbChecked Then
            sTsk = "RotSRVD"
        Else
            sTsk = "RotSRV"
        End If
                
        lErro = objRelatorio.ExecutarDireto("Roteiro de Serviço", "", 0, sTsk, "NNUMINTREL", CStr(lNumIntRel))
        If lErro <> SUCESSO Then gError 196163
    
    End If
    
    Exit Sub
    
Erro_BotaoRelRoteiro_Click:
    
    Select Case gErr
    
        Case 196161 To 196163
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196164)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoAbrirRoteiro_Click()

Dim lErro As Long
Dim objRoteiro As New ClassRoteirosDeFabricacao
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoAbrirRoteiro_Click

    'Se não tiver linha selecionada => Erro
    If GridMP.Row = 0 Then gError 196165

    lErro = CF("Produto_Formata", GridMP.TextMatrix(GridMP.Row, iGrid_MPProduto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 196166
    
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 196167
    
    objProduto.sCodigo = sProdutoFormatado
    
    'Lê o MPProduto Componente do Kit
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 196168
    
    'se o MPProduto não estiver cadastrado... erro
    If lErro <> SUCESSO Then gError 196169
    
    If objProduto.iCompras <> PRODUTO_PRODUZIVEL Then gError 196170

    If Len(Trim(GridMP.TextMatrix(GridMP.Row, iGrid_MPVersao_Col))) = 0 Then gError 196171

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 196172

    objRoteiro.sProdutoRaiz = objProduto.sCodigo
    objRoteiro.sVersao = GridMP.TextMatrix(GridMP.Row, iGrid_MPVersao_Col)

    Call Chama_Tela("RoteirosDeFabricacao", objRoteiro)

    Exit Sub
    
Erro_BotaoAbrirRoteiro_Click:
    
    Select Case gErr
    
        Case 196165
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    
        Case 196166, 196168, 196172

        Case 196167
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PREENCHIDO", gErr)
        
        Case 196169
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
    
        Case 196170
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_SERVICO", gErr, objProduto.sCodigo)
    
        Case 196171
            Call Rotina_Erro(vbOKOnly, "ERRO_VERSAO_ROTEIROSRV_NAO_PREENCHIDO", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196173)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoOnde_Click()

'PS-> POR ENQUANTO NÃO TEM SUB-ROTEIROS

'Dim lErro As Long
'Dim colSelecao As New Collection
'Dim sProdutoFormatado As String
'Dim iProdutoPreenchido As Integer
'Dim objRoteiro As New ClassRoteiroSRV
'Dim sFiltro As String
'
'On Error GoTo Erro_BotaoOnde_Click
'
'    lErro = CF("Produto_Formata", Servico.Text, sProdutoFormatado, iProdutoPreenchido)
'    If lErro <> SUCESSO Then gError 196174
'
'    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 196175
'
'    If Len(Trim(Versao.Text)) = 0 Then gError 196176
'
'    colSelecao.Add sProdutoFormatado
'    colSelecao.Add Versao.Text
'
'    sFiltro = "EXISTS (SELECT R.Servico FROM RoteiroSRV AS R, RoteiroSRVOper AS O, RoteiroSRVOperMP AS I WHERE R.NumIntDoc = O.NumIntDocRotSRV AND O.NumIntDoc = I.NumIntDocOper AND R.Servico = RoteiroSRV.Servico AND R.Versao = RoteiroSRV.Versao AND I.Produto = ? AND I.Versao = ?)"
'
'    Call Chama_Tela("RoteiroSRVLista", colSelecao, objRoteiro, objEventoRSRV, sFiltro)
'
'    Exit Sub
'
'Erro_BotaoOnde_Click:
'
'    Select Case gErr
'
'        Case 196174
'
'        Case 196175
'            Call Rotina_Erro(vbOKOnly, "ERRO_Servico_ROTFABR_NAO_PREENCHIDO", gErr)
'
'        Case 196176
'            Call Rotina_Erro(vbOKOnly, "ERRO_VERSAO_ROTEIROSRV_NAO_PREENCHIDO", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196177)
'
'    End Select
'
'    Exit Sub
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Competencia Then Call CompetenciaLabel_Click
        If Me.ActiveControl Is CTPadrao Then Call CTLabel_Click
        If Me.ActiveControl Is MPProduto Then Call BotaoProdutos_Click
        If Me.ActiveControl Is MaqCodigo Then Call BotaoMaq_Click
        If Me.ActiveControl Is MOCodigo Then Call BotaoMO_Click
        If Me.ActiveControl Is Servico Then Call LabelProduto_Click
        If Me.ActiveControl Is Versao Then Call LabelVersao_Click
    
    End If
    
End Sub

Private Function Inicializa_GridMaq(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Máquina")
    objGrid.colColuna.Add ("Qtde")
    objGrid.colColuna.Add ("Horas")
    objGrid.colColuna.Add ("Observação")
    
    'Controles que participam do Grid
    objGrid.colCampo.Add (MaqCodigo.Name)
    objGrid.colCampo.Add (MaqQuantidade.Name)
    objGrid.colCampo.Add (MaqHoras.Name)
    objGrid.colCampo.Add (MaqOBS.Name)
    
    'Colunas do Grid
    iGrid_MaqCodigo_Col = 1
    iGrid_MaqQuantidade_Col = 2
    iGrid_MaqHoras_Col = 3
    iGrid_MaqOBS_Col = 4
    
    objGrid.objGrid = GridMaq

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 4
    
    'Largura da primeira coluna
    GridMaq.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridMaq = SUCESSO

End Function

Private Sub GridMaq_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridMaq, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMaq, iAlterado)
    End If

End Sub

Private Sub GridMaq_GotFocus()
    Call Grid_Recebe_Foco(objGridMaq)
End Sub

Private Sub GridMaq_EnterCell()
    Call Grid_Entrada_Celula(objGridMaq, iAlterado)
End Sub

Private Sub GridMaq_LeaveCell()
    Call Saida_Celula(objGridMaq)
End Sub

Private Sub GridMaq_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridMaq, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMaq, iAlterado)
    End If

End Sub

Private Sub GridMaq_RowColChange()
    Call Grid_RowColChange(objGridMaq)
End Sub

Private Sub GridMaq_Scroll()
    Call Grid_Scroll(objGridMaq)
End Sub

Private Sub GridMaq_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridMaq)
End Sub

Private Sub GridMaq_LostFocus()
    Call Grid_Libera_Foco(objGridMaq)
End Sub

Private Function Inicializa_GridMO(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Código")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("Qtde")
    objGrid.colColuna.Add ("Horas")
    objGrid.colColuna.Add ("Observação")
    
    'Controles que participam do Grid
    objGrid.colCampo.Add (MOCodigo.Name)
    objGrid.colCampo.Add (MODescricao.Name)
    objGrid.colCampo.Add (MOQuantidade.Name)
    objGrid.colCampo.Add (MOHoras.Name)
    objGrid.colCampo.Add (MOOBS.Name)
    
    'Colunas do Grid
    iGrid_MOCodigo_Col = 1
    iGrid_MODescricao_Col = 2
    iGrid_MOQuantidade_Col = 3
    iGrid_MOHoras_Col = 4
    iGrid_MOOBS_Col = 5
    
    objGrid.objGrid = GridMO

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 4

    'Largura da primeira coluna
    GridMO.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridMO = SUCESSO

End Function

Private Sub GridMO_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridMO, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMO, iAlterado)
    End If

End Sub

Private Sub GridMO_GotFocus()
    Call Grid_Recebe_Foco(objGridMO)
End Sub

Private Sub GridMO_EnterCell()
    Call Grid_Entrada_Celula(objGridMO, iAlterado)
End Sub

Private Sub GridMO_LeaveCell()
    Call Saida_Celula(objGridMO)
End Sub

Private Sub GridMO_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridMO, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMO, iAlterado)
    End If

End Sub

Private Sub GridMO_RowColChange()
    Call Grid_RowColChange(objGridMO)
End Sub

Private Sub GridMO_Scroll()
    Call Grid_Scroll(objGridMO)
End Sub

Private Sub GridMO_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridMO)
End Sub

Private Sub GridMO_LostFocus()
    Call Grid_Libera_Foco(objGridMO)
End Sub

Private Sub MOCodigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MOCodigo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMO)
End Sub

Private Sub MOCodigo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMO)
End Sub

Private Sub MOCodigo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMO.objControle = MOCodigo
    lErro = Grid_Campo_Libera_Foco(objGridMO)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MOQuantidade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MOQuantidade_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMO)
End Sub

Private Sub MOQuantidade_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMO)
End Sub

Private Sub MOQuantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMO.objControle = MOQuantidade
    lErro = Grid_Campo_Libera_Foco(objGridMO)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MODescricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MODescricao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMO)
End Sub

Private Sub MODescricao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMO)
End Sub

Private Sub MODescricao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMO.objControle = MODescricao
    lErro = Grid_Campo_Libera_Foco(objGridMO)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MOHoras_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MOHoras_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMO)
End Sub

Private Sub MOHoras_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMO)
End Sub

Private Sub MOHoras_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMO.objControle = MOHoras
    lErro = Grid_Campo_Libera_Foco(objGridMO)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_MOCodigo(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objTiposDeMaodeObra As New ClassTiposDeMaodeObra

On Error GoTo Erro_Saida_Celula_MOCodigo

    Set objGridInt.objControle = MOCodigo
    
    'Se o campo foi preenchido
    If Len(MOCodigo.Text) > 0 Then
        
        objTiposDeMaodeObra.iCodigo = StrParaInt(MOCodigo.Text)
        
        'Lê o TiposDeMaodeObra que está sendo Passado
        lErro = CF("TiposDeMaodeObra_Le", objTiposDeMaodeObra)
        If lErro <> SUCESSO And lErro <> 137598 Then gError 196178
    
        If lErro <> SUCESSO Then gError 196179

        GridMO.TextMatrix(GridMO.Row, iGrid_MODescricao_Col) = objTiposDeMaodeObra.sDescricao
       
        'verifica se precisa preencher o grid com uma nova linha
        If GridMO.Row - GridMO.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 196180

    Saida_Celula_MOCodigo = SUCESSO

    Exit Function

Erro_Saida_Celula_MOCodigo:

    Saida_Celula_MOCodigo = gErr

    Select Case gErr
    
        Case 196178, 196180
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 196179
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOSDEMAODEOBRA_NAO_CADASTRADO", gErr, objTiposDeMaodeObra.iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196181)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MOQuantidade(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MOQuantidade

    Set objGridInt.objControle = MOQuantidade
    
    'se a quantidade foi preenchida
    If Len(MOQuantidade.ClipText) > 0 Then

        lErro = Valor_Inteiro_Critica(MOQuantidade.Text)
        If lErro <> SUCESSO Then gError 196182
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 196183

    Saida_Celula_MOQuantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_MOQuantidade:

    Saida_Celula_MOQuantidade = gErr

    Select Case gErr
    
        Case 196182
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 196183
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196184)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MOHoras(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MOHoras

    Set objGridInt.objControle = MOHoras
    
    'se a quantidade foi preenchida
    If Len(MOHoras.ClipText) > 0 Then

        lErro = Valor_Positivo_Critica(MOHoras.Text)
        If lErro <> SUCESSO Then gError 196185
    
        MOHoras.Text = Formata_Estoque(MOHoras.Text)
           
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 196186

    Saida_Celula_MOHoras = SUCESSO

    Exit Function

Erro_Saida_Celula_MOHoras:

    Saida_Celula_MOHoras = gErr

    Select Case gErr
    
        Case 196185
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 196186
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196187)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Sub MaqCodigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MaqCodigo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMaq)
End Sub

Private Sub MaqCodigo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaq)
End Sub

Private Sub MaqCodigo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaq.objControle = MaqCodigo
    lErro = Grid_Campo_Libera_Foco(objGridMaq)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MaqQuantidade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MaqQuantidade_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMaq)
End Sub

Private Sub MaqQuantidade_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaq)
End Sub

Private Sub MaqQuantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaq.objControle = MaqQuantidade
    lErro = Grid_Campo_Libera_Foco(objGridMaq)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MaqHoras_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MaqHoras_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMaq)
End Sub

Private Sub MaqHoras_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaq)
End Sub

Private Sub MaqHoras_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaq.objControle = MaqHoras
    lErro = Grid_Campo_Libera_Foco(objGridMaq)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_MaqCodigo(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objMaquinas As New ClassMaquinas

On Error GoTo Erro_Saida_Celula_MaqCodigo

    Set objGridInt.objControle = MaqCodigo
    
    'Se o campo foi preenchido
    If Len(Trim(MaqCodigo.Text)) > 0 Then
    
        Set objMaquinas = New ClassMaquinas
    
        'Verifica sua existencia
        lErro = CF("TP_Maquina_Le", MaqCodigo, objMaquinas)
        If lErro <> SUCESSO Then gError 196188
        
        'verifica se precisa preencher o grid com uma nova linha
        If GridMaq.Row - GridMaq.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 196189

    Saida_Celula_MaqCodigo = SUCESSO

    Exit Function

Erro_Saida_Celula_MaqCodigo:

    Saida_Celula_MaqCodigo = gErr

    Select Case gErr
    
        Case 196188, 196189
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196190)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MaqQuantidade(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MaqQuantidade

    Set objGridInt.objControle = MaqQuantidade
    
    'se a quantidade foi preenchida
    If Len(MaqQuantidade.ClipText) > 0 Then

        lErro = Valor_Inteiro_Critica(MaqQuantidade.Text)
        If lErro <> SUCESSO Then gError 196191
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 196192

    Saida_Celula_MaqQuantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_MaqQuantidade:

    Saida_Celula_MaqQuantidade = gErr

    Select Case gErr
           
        Case 196191
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 196192
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196193)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MaqHoras(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MaqHoras

    Set objGridInt.objControle = MaqHoras
    
    'se a quantidade foi preenchida
    If Len(MaqHoras.ClipText) > 0 Then

        lErro = Valor_Positivo_Critica(MaqHoras.Text)
        If lErro <> SUCESSO Then gError 196194
    
        MaqHoras.Text = Formata_Estoque(MaqHoras.Text)
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 196195

    Saida_Celula_MaqHoras = SUCESSO

    Exit Function

Erro_Saida_Celula_MaqHoras:

    Saida_Celula_MaqHoras = gErr

    Select Case gErr

        Case 196194
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 196195
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196196)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Sub BotaoMO_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objTiposDeMaodeObras As New ClassTiposDeMaodeObra

On Error GoTo Erro_BotaoMO_Click

    If Me.ActiveControl Is MOCodigo Then
        objTiposDeMaodeObras.iCodigo = StrParaInt(MOCodigo)
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridMO.Row = 0 Then gError 196197

        objTiposDeMaodeObras.iCodigo = StrParaInt(GridMO.TextMatrix(GridMO.Row, iGrid_MOCodigo_Col))
        
    End If

    Call Chama_Tela("TiposDeMaodeObraLista", colSelecao, objTiposDeMaodeObras, objEventoMO)

    Exit Sub

Erro_BotaoMO_Click:

    Select Case gErr
        
        Case 196197
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196198)

    End Select

    Exit Sub

End Sub

Private Sub objEventoMO_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTiposDeMaodeObra As ClassTiposDeMaodeObra
Dim iLinha As Integer

On Error GoTo Erro_objEventoMO_evSelecao

    Set objTiposDeMaodeObra = obj1
    
    MOCodigo.Text = CStr(objTiposDeMaodeObra.iCodigo)
    
    If Not (Me.ActiveControl Is MOCodigo) Then
    
        GridMO.TextMatrix(GridMO.Row, iGrid_MOCodigo_Col) = CStr(objTiposDeMaodeObra.iCodigo)
        GridMO.TextMatrix(GridMO.Row, iGrid_MODescricao_Col) = objTiposDeMaodeObra.sDescricao
    
        'verifica se precisa preencher o grid com uma nova linha
        If GridMO.Row - GridMO.FixedRows = objGridMO.iLinhasExistentes Then
            objGridMO.iLinhasExistentes = objGridMO.iLinhasExistentes + 1
        End If
        
    End If

    iAlterado = REGISTRO_ALTERADO
    
    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoMO_evSelecao:

    Select Case gErr
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196199)

    End Select

    Exit Sub

End Sub

Private Sub BotaoMaq_Click()

Dim lErro As Long
Dim objMaquinas As ClassMaquinas
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoMaq_Click

    Set objMaquinas = New ClassMaquinas

    If Me.ActiveControl Is MaqCodigo Then
        objMaquinas.sNomeReduzido = MaqCodigo.Text
    Else
        'Verifica se tem alguma linha selecionada no Grid
        If GridMaq.Row = 0 Then gError 196200
        objMaquinas.sNomeReduzido = GridMaq.TextMatrix(GridMaq.Row, iGrid_MaqCodigo_Col)
    End If
    
    'Le a Máquina no BD a partir do NomeReduzido
    lErro = CF("Maquinas_Le_NomeReduzido", objMaquinas)
    If lErro <> SUCESSO And lErro <> 103100 Then gError 196201
    
    Call Chama_Tela("MaquinasLista", colSelecao, objMaquinas, objEventoMaq, , "Nome Reduzido")

    Exit Sub

Erro_BotaoMaq_Click:

    Select Case gErr

        Case 196200
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 196201

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196202)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoMaq_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objMaquinas As ClassMaquinas

On Error GoTo Erro_objEventoMaq_evSelecao

    Set objMaquinas = obj1

    'Lê o Maquinas
    lErro = CF("TP_Maquina_Le", MaqCodigo, objMaquinas)
    If lErro <> SUCESSO Then gError 196202
    
    'Mostra os dados do Maquinas na tela
    MaqCodigo.Text = objMaquinas.sNomeReduzido
    
    If Not (Me.ActiveControl Is MaqCodigo) Then
        GridMaq.TextMatrix(GridMaq.Row, iGrid_MaqCodigo_Col) = objMaquinas.sNomeReduzido
    End If
    
    'verifica se precisa preencher o grid com uma nova linha
    If GridMaq.Row - GridMaq.FixedRows = objGridMaq.iLinhasExistentes Then
        objGridMaq.iLinhasExistentes = objGridMaq.iLinhasExistentes + 1
    End If
    
    iAlterado = REGISTRO_ALTERADO
    
    Me.Show

    Exit Sub

Erro_objEventoMaq_evSelecao:

    Select Case gErr

        Case 196202
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196203)

    End Select

    Exit Sub

End Sub

Function Move_OperMO_Memoria(ByVal objOper As ClassRoteiroSRVOper) As Long

Dim lErro As Long
Dim iTipo As Integer
Dim iIndice As Integer
Dim objRoteiroSRVOperMO As ClassRoteiroSRVOperMO

On Error GoTo Erro_Move_OperMO_Memoria

    For iIndice = 1 To objGridMO.iLinhasExistentes
    
        Set objRoteiroSRVOperMO = New ClassRoteiroSRVOperMO
        
        objRoteiroSRVOperMO.iCodMO = StrParaInt(GridMO.TextMatrix(iIndice, iGrid_MOCodigo_Col))
        objRoteiroSRVOperMO.iQtd = StrParaInt(GridMO.TextMatrix(iIndice, iGrid_MOQuantidade_Col))
        objRoteiroSRVOperMO.dHoras = StrParaDbl(GridMO.TextMatrix(iIndice, iGrid_MOHoras_Col))
        objRoteiroSRVOperMO.iSeq = iIndice
        objRoteiroSRVOperMO.sObs = GridMO.TextMatrix(iIndice, iGrid_MOOBS_Col)
        
        objOper.colMO.Add objRoteiroSRVOperMO
    
    Next

    Move_OperMO_Memoria = SUCESSO

    Exit Function

Erro_Move_OperMO_Memoria:

    Move_OperMO_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196204)

    End Select

    Exit Function

End Function

Function Move_OperMaq_Memoria(ByVal objOper As ClassRoteiroSRVOper) As Long

Dim lErro As Long
Dim iTipo As Integer
Dim iIndice As Integer
Dim objRoteiroSRVOperMaq As ClassRoteiroSRVOperMaq
Dim objMaquina As ClassMaquinas

On Error GoTo Erro_Move_OperMaq_Memoria

    For iIndice = 1 To objGridMaq.iLinhasExistentes
    
        Set objRoteiroSRVOperMaq = New ClassRoteiroSRVOperMaq
        Set objMaquina = New ClassMaquinas
        
        objMaquina.sNomeReduzido = GridMaq.TextMatrix(iIndice, iGrid_MaqCodigo_Col)
        
        lErro = CF("Maquinas_Le_NomeReduzido", objMaquina)
        If lErro <> SUCESSO And lErro <> 103100 Then gError 196205
        
        objRoteiroSRVOperMaq.iCodMaq = objMaquina.iCodigo
        objRoteiroSRVOperMaq.iFilialEmpMaq = objMaquina.iFilialEmpresa
        objRoteiroSRVOperMaq.iQtd = StrParaInt(GridMaq.TextMatrix(iIndice, iGrid_MaqQuantidade_Col))
        objRoteiroSRVOperMaq.dHoras = StrParaDbl(GridMaq.TextMatrix(iIndice, iGrid_MaqHoras_Col))
        objRoteiroSRVOperMaq.iSeq = iIndice
        objRoteiroSRVOperMaq.sObs = GridMaq.TextMatrix(iIndice, iGrid_MaqOBS_Col)
        
        objOper.colMaq.Add objRoteiroSRVOperMaq
    
    Next

    Move_OperMaq_Memoria = SUCESSO

    Exit Function

Erro_Move_OperMaq_Memoria:

    Move_OperMaq_Memoria = gErr

    Select Case gErr
    
        Case 196205

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196206)

    End Select

    Exit Function

End Function

Function Preenche_GridMO(ByVal objOper As ClassRoteiroSRVOper) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objRoteiroSRVOperMO As ClassRoteiroSRVOperMO
Dim objMO As ClassTiposDeMaodeObra

On Error GoTo Erro_Preenche_GridMO
    
    'Exibe os dados da coleção de Competencias na tela
    For Each objRoteiroSRVOperMO In objOper.colMO
        
        iLinha = iLinha + 1
        
        Set objMO = New ClassTiposDeMaodeObra
        
        objMO.iCodigo = objRoteiroSRVOperMO.iCodMO
        
        'Lê o TiposDeMaodeObra que está sendo Passado
        lErro = CF("TiposDeMaodeObra_Le", objMO)
        If lErro <> SUCESSO And lErro <> 137598 Then gError 196207
        
        GridMO.TextMatrix(iLinha, iGrid_MODescricao_Col) = objMO.sDescricao
        GridMO.TextMatrix(iLinha, iGrid_MOCodigo_Col) = CStr(objRoteiroSRVOperMO.iCodMO)
        GridMO.TextMatrix(iLinha, iGrid_MOQuantidade_Col) = CStr(objRoteiroSRVOperMO.iQtd)
        GridMO.TextMatrix(iLinha, iGrid_MOHoras_Col) = Formata_Estoque(objRoteiroSRVOperMO.dHoras)
        GridMO.TextMatrix(iLinha, iGrid_MOOBS_Col) = objRoteiroSRVOperMO.sObs

    Next

    objGridMO.iLinhasExistentes = iLinha

    Preenche_GridMO = SUCESSO

    Exit Function

Erro_Preenche_GridMO:

    Preenche_GridMO = gErr

    Select Case gErr
    
        Case 196207

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196208)

    End Select

    Exit Function

End Function

Function Preenche_GridMaq(ByVal objOper As ClassRoteiroSRVOper) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objRoteiroSRVOperMaq As ClassRoteiroSRVOperMaq
Dim objMaquina As ClassMaquinas

On Error GoTo Erro_Preenche_GridMaq
    
    'Exibe os dados da coleção de Competencias na tela
    For Each objRoteiroSRVOperMaq In objOper.colMaq
        
        iLinha = iLinha + 1
        
        Set objMaquina = New ClassMaquinas
        
        objMaquina.iCodigo = objRoteiroSRVOperMaq.iCodMaq
        objMaquina.iFilialEmpresa = objRoteiroSRVOperMaq.iFilialEmpMaq
        
        lErro = CF("Maquinas_Le", objMaquina)
        If lErro <> SUCESSO And lErro <> 103090 Then gError 196209
        
        GridMaq.TextMatrix(iLinha, iGrid_MaqCodigo_Col) = objMaquina.sNomeReduzido
        GridMaq.TextMatrix(iLinha, iGrid_MaqQuantidade_Col) = objRoteiroSRVOperMaq.iQtd
        GridMaq.TextMatrix(iLinha, iGrid_MaqHoras_Col) = Formata_Estoque(objRoteiroSRVOperMaq.dHoras)
        GridMaq.TextMatrix(iLinha, iGrid_MaqOBS_Col) = objRoteiroSRVOperMaq.sObs
        
    Next

    objGridMaq.iLinhasExistentes = iLinha

    Preenche_GridMaq = SUCESSO

    Exit Function

Erro_Preenche_GridMaq:

    Preenche_GridMaq = gErr

    Select Case gErr
    
        Case 196209

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196210)

    End Select

    Exit Function

End Function

Private Sub MaqOBS_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MaqOBS_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMaq)
End Sub

Private Sub MaqOBS_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaq)
End Sub

Private Sub MaqOBS_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaq.objControle = MaqOBS
    lErro = Grid_Campo_Libera_Foco(objGridMaq)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MOOBS_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MOOBS_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMO)
End Sub

Private Sub MOOBS_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMO)
End Sub

Private Sub MOOBS_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMO.objControle = MOOBS
    lErro = Grid_Campo_Libera_Foco(objGridMO)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MPOBS_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MPOBS_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMP)
End Sub

Private Sub MPOBS_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMP)
End Sub

Private Sub MPOBS_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMP.objControle = MPObs
    lErro = Grid_Campo_Libera_Foco(objGridMP)
    If lErro <> SUCESSO Then Cancel = True

End Sub
