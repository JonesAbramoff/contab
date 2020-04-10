VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.UserControl TecladoProduto 
   ClientHeight    =   6180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10575
   ScaleHeight     =   6180
   ScaleWidth      =   10575
   Begin VB.CheckBox RemoveBuracos 
      Caption         =   "Ordenar botões ativos antes dos inativos ao gravar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6105
      TabIndex        =   66
      Top             =   750
      Width           =   4290
   End
   Begin VB.CommandButton BotaoRemover 
      Caption         =   "Remover Submenu"
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
      Left            =   9045
      TabIndex        =   65
      ToolTipText     =   "Exibe lista de produtos"
      Top             =   2160
      Width           =   1395
   End
   Begin VB.CommandButton BotaoSubmenu 
      Caption         =   "Criar Submenu"
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
      Left            =   7605
      TabIndex        =   64
      ToolTipText     =   "Exibe lista de produtos"
      Top             =   2160
      Width           =   1380
   End
   Begin MSComctlLib.TreeView TvwMenu 
      Height          =   3345
      Left            =   7605
      TabIndex        =   63
      Top             =   2670
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   5900
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
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
      Height          =   240
      Left            =   6105
      TabIndex        =   48
      ToolTipText     =   "Indica se o teclado é o padrão da tela de vendas"
      Top             =   315
      Width           =   945
   End
   Begin VB.Frame FrameBotaoModelo 
      Caption         =   "Tecla Modelo"
      Height          =   1635
      Left            =   210
      TabIndex        =   46
      Top             =   825
      Width           =   2070
      Begin VB.CommandButton BotaoModelo 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   525
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Exibe como as  alterações vão refletir no botão"
         Top             =   525
         Width           =   930
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   8265
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "TecladoProduto.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "TecladoProduto.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "TecladoProduto.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "TecladoProduto.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame FrameProdutos 
      Caption         =   "Produtos"
      Height          =   3465
      Left            =   195
      TabIndex        =   5
      ToolTipText     =   "Teclado de Produtos"
      Top             =   2580
      Width           =   7260
      Begin VB.CommandButton Produto 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   0
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   1
         Left            =   1210
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   240
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   2
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   255
         Width           =   945
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   3
         Left            =   3105
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   255
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   34
         Left            =   5955
         TabIndex        =   36
         Top             =   2595
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   4
         Left            =   4064
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   255
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   5
         Left            =   5010
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   255
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   6
         Left            =   5955
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   255
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   12
         Left            =   5010
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   840
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   13
         Left            =   5955
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   840
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   7
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   840
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   8
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   840
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   9
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   840
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   10
         Left            =   3105
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   840
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   11
         Left            =   4065
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   840
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   18
         Left            =   4065
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1425
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   19
         Left            =   5010
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1425
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   20
         Left            =   5955
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1425
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   14
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1425
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   15
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1425
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   16
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1425
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   17
         Left            =   3105
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1425
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   26
         Left            =   5010
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2010
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   21
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2010
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   22
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2010
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   23
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2010
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   24
         Left            =   3105
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2010
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   25
         Left            =   4065
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2010
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   27
         Left            =   5955
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2010
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   33
         Left            =   5010
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2595
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   32
         Left            =   4065
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2595
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   31
         Left            =   3105
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2595
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   30
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2595
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   29
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2595
         Width           =   950
      End
      Begin VB.CommandButton Produto 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   28
         Left            =   255
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2595
         UseMaskColor    =   -1  'True
         Width           =   950
      End
   End
   Begin VB.CommandButton BotaoAplicar 
      Caption         =   "Aplicar"
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
      Left            =   7785
      TabIndex        =   4
      ToolTipText     =   "Passa as alterações feitas no botão modelo para o botão no teclado"
      Top             =   1215
      Width           =   1140
   End
   Begin VB.ComboBox ComboTeclado 
      Height          =   315
      Left            =   3285
      TabIndex        =   3
      ToolTipText     =   "Qual é o tipo de Teclado para qual essa configuração foi criada"
      Top             =   720
      Width           =   2610
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   300
      Left            =   2055
      Picture         =   "TecladoProduto.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Numeração Automática"
      Top             =   255
      Width           =   300
   End
   Begin VB.CommandButton BotaoLimparTecla 
      Height          =   315
      Left            =   4875
      Picture         =   "TecladoProduto.ctx":0A7E
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Limpar"
      Top             =   2250
      Width           =   420
   End
   Begin TelasLoja.ColorBrowser Fundo 
      Height          =   300
      Left            =   6615
      TabIndex        =   1
      ToolTipText     =   "Fundo do Botão de Produto"
      Top             =   1800
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   529
   End
   Begin MSMask.MaskEdBox Titulo 
      Height          =   315
      Left            =   3285
      TabIndex        =   49
      ToolTipText     =   "Título do Botão de Produto"
      Top             =   1785
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   556
      _Version        =   393216
      AllowPrompt     =   -1  'True
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1230
      TabIndex        =   50
      ToolTipText     =   "Código do Teclado"
      Top             =   240
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   556
      _Version        =   393216
      AllowPrompt     =   -1  'True
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Descricao 
      Height          =   315
      Left            =   3285
      TabIndex        =   51
      ToolTipText     =   "Descrição para o teclado"
      Top             =   240
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   556
      _Version        =   393216
      AllowPrompt     =   -1  'True
      MaxLength       =   50
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Tecla 
      Height          =   315
      Left            =   3780
      TabIndex        =   52
      ToolTipText     =   "Tecla relacionada ao produto"
      Top             =   2250
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   556
      _Version        =   393216
      AllowPrompt     =   -1  'True
      MaxLength       =   7
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox CodProduto 
      Height          =   315
      Left            =   3285
      TabIndex        =   53
      ToolTipText     =   "Produto que será representado pelo botão"
      Top             =   1245
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      _Version        =   393216
      AllowPrompt     =   -1  'True
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Desc:"
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
      Left            =   2760
      TabIndex        =   62
      ToolTipText     =   "Descrição para o teclado"
      Top             =   300
      Width           =   510
   End
   Begin VB.Label LabelCodigo 
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
      Left            =   465
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   61
      ToolTipText     =   "Código do Teclado"
      Top             =   285
      Width           =   660
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Título:"
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
      Left            =   2685
      TabIndex        =   60
      ToolTipText     =   "Título do Botão de Produto"
      Top             =   1830
      Width           =   585
   End
   Begin VB.Label DescProduto 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   4650
      TabIndex        =   59
      ToolTipText     =   "Produto que será representado pelo botão"
      Top             =   1245
      Width           =   2940
   End
   Begin VB.Label LabelProduto 
      Alignment       =   1  'Right Justify
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
      Left            =   2535
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   58
      ToolTipText     =   "Produto que será representado pelo botão"
      Top             =   1305
      Width           =   735
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Fundo:"
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
      Left            =   5985
      TabIndex        =   57
      ToolTipText     =   "Fundo do Botão de Produto"
      Top             =   1815
      Width           =   600
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Tecla:"
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
      Left            =   2715
      TabIndex        =   56
      ToolTipText     =   "Tecla relacionada ao produto"
      Top             =   2295
      Width           =   555
   End
   Begin VB.Label LabelTeclado 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Teclado:"
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
      Left            =   2505
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   55
      ToolTipText     =   "Qual é o tipo de Teclado para qual essa configuração foi criada"
      Top             =   780
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Ctrl +"
      Height          =   195
      Left            =   3330
      TabIndex        =   54
      ToolTipText     =   "Tecla relacionada ao produto"
      Top             =   2310
      Width           =   360
   End
End
Attribute VB_Name = "TecladoProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'NA PRIMEIRA VERSÃO ESSA TELA NÃO SERÁ IMPLEMENTADA

'Em todos os botões do teclado e no Espelho coloca como DragIcon uma maozinha.

'O botaoOutros é permanente, não pode ser excluído e fica sempre no canto inferior direito
'do teclado.

'Com qualquer botao que já está no teclado faz a mesma coisa
'que faz com o botao modelo só que deixando pegar todos os
'lados da borda. Se teclar a tecla Del, elimina o botão.

'Property Variables:
Dim m_Caption As String
Event Unload()

'Variáveis globais a tela
Dim iAlterado As Integer
Dim iProdutoAlterado As Integer
Dim iDadosBotaoAlterado As Integer
Dim gColTecladoProdutoItens As New Collection

Private WithEvents objEventoTecladoProduto As AdmEvento
Attribute objEventoTecladoProduto.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoTeclado As AdmEvento
Attribute objEventoTeclado.VB_VarHelpID = -1

'Private Sub Desmembrar_Click()
'
'Dim objLog As New ClassLog
'Dim objTecladoProduto As ClassTecladoProduto
'
'    Call Limpa_Tela_TecladoProduto
'
'    Call Log_Le(objLog)
'
'    Call TecladoProduto_Desmembra_Log(objTecladoProduto, objLog)
'
'    Call Traz_TecladoProduto_Tela(objTecladoProduto)
'
'End Sub

'Function Log_Le(objLog As ClassLog) As Long
'
'Dim lErro As Long
'Dim tLog As typeLog
'Dim lComando As Long
'
'On Error GoTo Erro_Log_Le
'
'    'Abre o comando
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 104197
'
'    'Inicializa o Buffer da Variáveis String
'    tLog.sLog1 = String(STRING_CONCATENACAO, 0)
'    tLog.sLog2 = String(STRING_CONCATENACAO, 0)
'    tLog.sLog3 = String(STRING_CONCATENACAO, 0)
'    tLog.sLog4 = String(STRING_CONCATENACAO, 0)
'
'    'Seleciona código e nome dos meios de pagamentos da tabela AdmMeioPagto
'    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Operacao, Log1, Log2, Log3, Log4 , Data , Hora FROM Log", tLog.lNumIntDoc, tLog.iOperacao, tLog.sLog1, tLog.sLog2, tLog.sLog3, tLog.sLog4, tLog.dtData, tLog.dHora)
'    If lErro <> SUCESSO Then gError 104198
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104199
'
'
'    If lErro = AD_SQL_SUCESSO Then
'
'        'Carrega o objLog com as Infromações de bonco de dados
'        objLog.lNumIntDoc = tLog.lNumIntDoc
'        objLog.iOperacao = tLog.iOperacao
'        objLog.sLog = tLog.sLog1 & tLog.sLog2 & tLog.sLog3 & tLog.sLog4
'        objLog.dtData = tLog.dtData
'        objLog.dHora = tLog.dHora
'
'    End If
'
'    If lErro = AD_SQL_SEM_DADOS Then gError 104202
'
'    Log_Le = SUCESSO
'
'    'Fecha o comando
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'Erro_Log_Le:
'
'    Log_Le = gErr
'
'   Select Case gErr
'
'    Case gErr
'
'        Case 104198, 104199
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOG", gErr)
'
'        Case 104202
'            Call Rotina_Erro(vbOKOnly, "ERRO_LOG_NAO_EXISTENTE", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174557)
'
'        End Select
'
'    'Fecha o comando
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function

Private Sub BotaoLimparTecla_Click()
    Tecla.Text = ""
End Sub

Private Sub BotaoRemover_Click()

Dim objNode As Node
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objTecladoProdutoItens As ClassTecladoProdutoItem
Dim iCount As Integer
Dim lErro As Long

On Error GoTo Erro_BotaoRemover_Click

    Set objNode = TvwMenu.SelectedItem

    If Not objNode Is Nothing Then
        
        If objNode.Text = "<Vazio>" Then
            
            TvwMenu.Nodes.Remove (objNode.Key)
            
            For iIndice = 0 To Produto.Count - 1
                Produto(iIndice).Caption = ""
                Produto(iIndice).Tag = ""
                Produto(iIndice).BackColor = COR_DEFAULT
            Next
            
            BotaoModelo.Caption = ""
            BotaoModelo.Tag = ""
            BotaoModelo.BackColor = COR_DEFAULT
        
            DescProduto.Caption = ""
            Fundo.SelectedColor = COR_DEFAULT
        
            CodProduto.PromptInclude = False
            CodProduto.Text = ""
            CodProduto.PromptInclude = True
            
            Tecla.Text = ""
            Titulo.Text = ""
            
            If TvwMenu.Nodes.Count > 0 Then
                Call TvwMenu_NodeClick(TvwMenu.Nodes.Item(1).FirstSibling)
                TvwMenu.SelectedItem = TvwMenu.Nodes.Item(1).FirstSibling
            Else
                TvwMenu.SelectedItem = Nothing
            End If
            
        Else
        
            For iIndice = gColTecladoProdutoItens.Count To 1 Step -1
                Set objTecladoProdutoItens = gColTecladoProdutoItens(iIndice)
                If Left(objTecladoProdutoItens.sArvoreKey, Len(objNode.Key)) = objNode.Key Then
                    iCount = iCount + 1
                End If
            Next
            
            If iCount > 1 Then
                'vbMsgBox = MsgBox("este elemento contem filhos. Confirma a remoção?", vbYesNo)
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ELEMENTO_COM_FILHOS", objNode.Text)
                
            ElseIf iCount = 1 Then
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ELEMENTO_REMOCAO", objNode.Text)
            Else
                'vbMsgBox = MsgBox("este elemento nao esta cadastrado")
                gError 214917
            End If
            
            If vbMsgRes = vbYes Then
            
                TvwMenu.Nodes.Remove (objNode.Key)
                For iIndice = gColTecladoProdutoItens.Count To 1 Step -1
                    Set objTecladoProdutoItens = gColTecladoProdutoItens(iIndice)
                    If Left(objTecladoProdutoItens.sArvoreKey, Len(objNode.Key)) = objNode.Key Then
                        gColTecladoProdutoItens.Remove (objTecladoProdutoItens.sArvoreKey)
                    End If
                Next
                
                For iIndice = 0 To Produto.Count - 1
                    Produto(iIndice).Caption = ""
                    Produto(iIndice).Tag = ""
                    Produto(iIndice).BackColor = COR_DEFAULT
                Next
            
                BotaoModelo.Caption = ""
                BotaoModelo.Tag = ""
                BotaoModelo.BackColor = COR_DEFAULT
            
                DescProduto.Caption = ""
                Fundo.SelectedColor = COR_DEFAULT
            
                CodProduto.PromptInclude = False
                CodProduto.Text = ""
                CodProduto.PromptInclude = True
                
                Tecla.Text = ""
                Titulo.Text = ""
                
                If TvwMenu.Nodes.Count > 0 Then
                    Call TvwMenu_NodeClick(TvwMenu.Nodes.Item(1).FirstSibling)
                    TvwMenu.SelectedItem = TvwMenu.Nodes.Item(1).FirstSibling
                Else
                    TvwMenu.SelectedItem = Nothing
                End If
                
            End If
        End If

    Else
    
        gError 214923

    End If

    Exit Sub
    
Erro_BotaoRemover_Click:

    Select Case gErr
        
        Case 214917
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ELEMENTO_NAO_CADASTRADO", gErr)
        
                
        Case 214923
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ELEMENTO_TEM_QUE_ESTAR_SELECIONADO", gErr)
                
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 214918)

    End Select

    Exit Sub

End Sub

Private Sub BotaoSubmenu_Click()

Dim objNode As Node
Dim iIndice As Integer
Dim lErro As Long
Dim objTecladoProdutoItem As ClassTecladoProdutoItem
Dim iAchou As Integer

On Error GoTo Erro_BotaoSubmenu_Click

    Set objNode = TvwMenu.SelectedItem
    
    If Not objNode Is Nothing Then
    
        '**** mario  *****
        If Len(objNode.Key) = 11 Then gError 214920
        
        iAchou = 0
        
        For Each objTecladoProdutoItem In gColTecladoProdutoItens
            If objTecladoProdutoItem.sArvoreKey = objNode.Key Then
                iAchou = 1
                Exit For
            End If
        Next
        
        'se o nó pai ainda não está criado ==> nao pode ter submenu
        If iAchou = 0 Then gError 214921
    
        'se o botao pai esta associado a um codigo de produto ==> nao pode ter submenu
        If Len(objTecladoProdutoItem.sProduto) <> 0 Then gError 214922
    
        If objNode.Children = 0 Then
            Set objNode = TvwMenu.Nodes.Add(objNode.Key, tvwChild, objNode.Key & "01", "<Vazio>")
'            TreeView1.SelectedItem = TreeView1.Nodes.Item(objNode.Key & "01")
            TvwMenu.SelectedItem = objNode
            
            For iIndice = 0 To Produto.Count - 1
                Produto(iIndice).Caption = ""
                Produto(iIndice).Tag = ""
                Produto(iIndice).BackColor = COR_DEFAULT
            Next
        
            BotaoModelo.Caption = ""
            BotaoModelo.Tag = ""
            BotaoModelo.BackColor = COR_DEFAULT
        
            DescProduto.Caption = ""
            Fundo.SelectedColor = COR_DEFAULT
        
            CodProduto.PromptInclude = False
            CodProduto.Text = ""
            CodProduto.PromptInclude = True
            
            Tecla.Text = ""
            Titulo.Text = ""
        
        Else
        
'            MsgBox ("esse elemento ja tem um submenu associado")
            gError 214914
        
        End If
    Else
    
'        MsgBox ("um elemento da arvore tem q estar selecionado")
        gError 214915
        
    End If
    
    Exit Sub
    
Erro_BotaoSubmenu_Click:

    Select Case gErr
        
        Case 214914
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ELEMENTO_JA_TEM_SUBMENU", gErr)
        
        Case 214915
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ELEMENTO_TEM_QUE_ESTAR_SELECIONADO", gErr)
        
        Case 214920
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NIVEL_MAX_SUBMENU_ATINGIDO", gErr)
        
        Case 214921
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NO_VAZIO_NAO_PODE_TER_SUBNIVEL", gErr)
        
        Case 214922
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NO_FOLHA_NAO_PODE_TER_SUBNIVEL", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 214916)

    End Select

    Exit Sub

End Sub


Private Sub LabelCodigo_Click()

Dim objTecladoProduto As New ClassTecladoProduto
Dim colSelecao As Collection
    
    If Len(Trim(Codigo.Text)) > 0 Then objTecladoProduto.iCodigo = StrParaInt(Codigo.Text)
    
    Call Chama_Tela("TecladoProdutoLista", colSelecao, objTecladoProduto, objEventoTecladoProduto)

    Exit Sub

End Sub

Private Sub LabelProduto_Click()

Dim sProduto As String
Dim iPreenchido As Integer
Dim lErro As Long
Dim colSelecao As New Collection
Dim sProduto1 As String
Dim objProduto As New ClassProduto
Dim sSelecaoSQL As String

On Error GoTo Erro_LabelProduto_Click
    
    'Se não tem botão selecionado -->erro
    If Len(Trim(BotaoModelo.Tag)) = 0 Then gError 99574
    
    If Len(Trim(CodProduto.Text)) > 0 Then
    
        sProduto1 = CodProduto.Text
    
        'Formata o produto contido na variável se estiver preenchida
        lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 99488
        
        'Se não estiver --> limpa a variável
        If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""
        
        objProduto.sCodigo = sProduto
    
    End If
    
    
    'Passagem da data no último parâmetro do chama_tela
    'Chama a tela de browse
    Call Chama_Tela("ProdutosLojaLista", colSelecao, objProduto, objEventoProduto)

    Exit Sub
        
Erro_LabelProduto_Click:
    
    Select Case gErr
        
        Case 99488
        
        Case 99574
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BOTAO_NAO_SELECIONADO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174560)

    End Select

    Exit Sub

End Sub

Private Sub objEventotecladoProduto_evSelecao(obj1 As Object)

Dim objTecladoProduto As New ClassTecladoProduto
Dim lErro As Long

On Error GoTo Erro_objEventotecladoProduto_evSelecao

    Set objTecladoProduto = obj1
    
    If objTecladoProduto.iCodigo > 0 Then
                       
        objTecladoProduto.iFilialEmpresa = giFilialEmpresa
        
        'Exibe os dados do teclado
        lErro = Traz_TecladoProduto_Tela(objTecladoProduto)
        If lErro <> SUCESSO Then gError 99485
            
    End If
    
    Me.Show
        
    Exit Sub

Erro_objEventotecladoProduto_evSelecao:
    
    Select Case gErr
        
        Case 99485
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174558)

    End Select
    
    Exit Sub

End Sub

Private Sub LabelTeclado_Click()

Dim objTeclado As New ClassTeclado
Dim colSelecao As Collection
    
    If ComboTeclado.ListIndex <> -1 Then objTeclado.iCodigo = ComboTeclado.ItemData(ComboTeclado.ListIndex)
    
    Call Chama_Tela("TecladoLista", colSelecao, objTeclado, objEventoTeclado)

    Exit Sub

End Sub

Private Sub objEventoteclado_evSelecao(obj1 As Object)

Dim objTeclado As ClassTeclado
Dim lErro As Long
Dim iIndex As Integer

On Error GoTo Erro_objEventoteclado_evSelecao

    Set objTeclado = obj1
    
    If objTeclado.iCodigo > 0 Then
        'Bus0ca na combo o TECLADO
        For iIndex = 0 To ComboTeclado.ListCount - 1
            'Quando achar, seleciona este item na combo
            If ComboTeclado.ItemData(iIndex) = objTeclado.iCodigo Then
                ComboTeclado.ListIndex = iIndex
                Exit For
            End If
        Next
    End If
    
    Me.Show
        
    Exit Sub

Erro_objEventoteclado_evSelecao:
    
    Select Case gErr
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174559)

    End Select
    
    Exit Sub

End Sub

'Public Sub BotaoProdutos_Click()
''Chama o browser do ProdutoLojaLista
''So traz produtos onde codigo de barras ou referencia está preenchida
'
'Dim sProduto As String
'Dim iPreenchido As Integer
'Dim lErro As Long
'Dim colSelecao As New Collection
'Dim sProduto1 As String
'Dim objProduto As New ClassProduto
'Dim sSelecaoSQL As String
'
'On Error GoTo Erro_BotaoProdutos_Click
'
'    'Se não tem botão selecionado -->erro
'    If Len(Trim(BotaoModelo.Tag)) = 0 Then gError 99574
'
'    If Len(Trim(CodProduto.Text)) > 0 Then
'
'        sProduto1 = CodProduto.Text
'
'        'Formata o produto contido na variável se estiver preenchida
'        lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
'        If lErro <> SUCESSO Then gError 99488
'
'        'Se não estiver --> limpa a variável
'        If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""
'
'        objProduto.sCodigo = sProduto
'
'    End If
'
'
'    'Passagem da data no último parâmetro do chama_tela
'    'Chama a tela de browse
'    Call Chama_Tela("ProdutosLojaLista", colSelecao, objProduto, objEventoProduto)
'
'    Exit Sub
'
'Erro_BotaoProdutos_Click:
'
'    Select Case gErr
'
'        Case 99488
'
'        Case 99574
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_BOTAO_NAO_SELECIONADO", gErr)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174560)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim sProduto As String
Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iIndice As Integer
Dim sProduto1 As String
Dim iPreenchido As Integer

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1
    
    If Len(Trim(objProduto.sCodigo)) > 0 Then
        
        CodProduto.PromptInclude = False
        CodProduto.Text = objProduto.sCodigo
        CodProduto.PromptInclude = True
        
        Call CodProduto_Validate(False)
            
    End If
    
    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174561)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objTecladoProduto As ClassTecladoProduto) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se houver TecladoProduto passado como parâmetro, exibe seus dados
    If Not (objTecladoProduto Is Nothing) Then

        objTecladoProduto.iFilialEmpresa = giFilialEmpresa

        If objTecladoProduto.iCodigo > 0 Then
            
            'Le o TecladoProduto
            lErro = CF("TecladoProduto_Le", objTecladoProduto)
            If lErro <> SUCESSO And lErro <> 99526 Then gError 99520
            
            'Se não encontrou
            If lErro = 99526 Then gError 99521
                        
            'Exibe os dados do teclado
            lErro = Traz_TecladoProduto_Tela(objTecladoProduto)
            If lErro <> SUCESSO Then gError 99487
            
        End If
    
    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr
     
    Select Case gErr
    
        Case 99487, 99520
        
        Case 99521
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TECLADOPRODUTO_NAO_ENCONTRADO", gErr, objTecladoProduto.iCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174562)

    End Select

    Exit Function

End Function

Function Move_Tela_Memoria(objTecladoProduto As ClassTecladoProduto) As Long
'Move os dados da tela para o objTecladoProduto
Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria
         
    'Move a FilialEmpresa que esta sendo Referenciada para a Memória
    objTecladoProduto.iFilialEmpresa = giFilialEmpresa
    
    'Move o Codigo Para Memoria
    objTecladoProduto.iCodigo = StrParaInt(Codigo.Text)
    
    'Move o descricao Para Memoria
    objTecladoProduto.sDescricao = Descricao.Text
    
    'Move o Teclado Para Memoria
    If ComboTeclado.ListIndex <> -1 Then
        objTecladoProduto.iTeclado = ComboTeclado.ItemData(ComboTeclado.ListIndex)
    End If
    
    'Move o Padrão Para Memoria
    objTecladoProduto.iPadrao = Padrao.Value
    
    Call Preenche_ColTecladoProdutoItens(objTecladoProduto)
    
    Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174563)
        
        End Select

    Exit Function
    
End Function

Function Traz_TecladoProduto_Tela(objTecladoProduto As ClassTecladoProduto) As Long
'Função que Traz as Informações do BD para o TecladoProduto
Dim lErro As Long
Dim iIndex As Integer

On Error GoTo Erro_Traz_TecladoProduto_Tela

    Call Limpa_Tela_TecladoProduto
    
    'Le os seus itens
    lErro = CF("TecladoProdutoItens_Le", objTecladoProduto)
    If lErro <> SUCESSO Then gError 99522
            
    'Traz o Codigo para a Tela
    Codigo.Text = objTecladoProduto.iCodigo
    
    'Traz o descricao para a Tela
    Descricao.Text = objTecladoProduto.sDescricao
        
    'Traz o Padrão para a Tela
    Padrao.Value = objTecladoProduto.iPadrao
    
    'Traz o Teclado para a Tela
    If objTecladoProduto.iTeclado > 0 Then
        'Busca na combo o TECLADO
        For iIndex = 0 To ComboTeclado.ListCount - 1
            'Quando achar, seleciona este item na combo
            If ComboTeclado.ItemData(iIndex) = objTecladoProduto.iTeclado Then
                ComboTeclado.ListIndex = iIndex
                Exit For
            End If
        Next
    End If
    
    'Preenche o array de botões da tela
    Call Traz_ColTecladoProdutoItens(objTecladoProduto.colTecladoProdutoItem)
    
    'Demonstra que não Houve Alteração na Tela
    iAlterado = 0
    
    Traz_TecladoProduto_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_TecladoProduto_Tela:

    Traz_TecladoProduto_Tela = gErr

    Select Case gErr
        
        Case 99522
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174564)
        
        End Select
        
        Exit Function
        
End Function

Public Sub Form_Load()
    
Dim colTeclado As New Collection
Dim objTeclado As New ClassTeclado
Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objEventoProduto = New AdmEvento
    Set objEventoTecladoProduto = New AdmEvento
    Set objEventoTeclado = New AdmEvento
    Set gColTecladoProdutoItens = New Collection
            
    'Mascara o produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", CodProduto)
    If lErro <> SUCESSO Then gError 99482
    
    'Inicializando combo de teclado
    lErro = CF("Teclado_Le_Todos", colTeclado)
    If lErro <> SUCESSO And lErro <> 99514 Then gError 99483
    
    If lErro = 99514 Then gError 86257
    
    
    For Each objTeclado In colTeclado
    'Adiciona o item na combo de Teclado e preenche o itemdata
        ComboTeclado.AddItem objTeclado.iCodigo & SEPARADOR & objTeclado.sDescricao
        ComboTeclado.ItemData(ComboTeclado.NewIndex) = objTeclado.iCodigo
    Next
    'Inicializa os botões do teclado com vazio
    Call Inicializa_ColTecladoProdutoItens
    
    Fundo.SelectedColor = COR_DEFAULT
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:
    
    Select Case gErr
    
        Case 86257
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_TECLADO_CADASTRADO", gErr)
        
        Case 99482, 99483
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174565)

    End Select
    
    Exit Sub

End Sub

Sub Inicializa_ColTecladoProdutoItens()

Dim iIndice As Integer
Dim objTecladoProdutoItens As ClassTecladoProdutoItem

    Set gColTecladoProdutoItens = New Collection

    For iIndice = 0 To 34

        Produto(iIndice).BackColor = COR_DEFAULT
        Produto(iIndice).Caption = ""
        Produto(iIndice).Tag = ""
        
   Next

End Sub

Private Sub BotaoProxNum_Click()
'Gera um novo número disponível para código do TecladoProduto

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click
    
    'Chama a função que gera o sequencial do Código Automático para o novo teclado
    lErro = CF("Config_Obter_Inteiro_Automatico", "LojaConfig", "NUM_PROXIMO_TECLADOPRODUTO", "TecladoProduto", "Codigo", iCodigo)
    
    If lErro <> SUCESSO Then gError 99516

    'Exibe o novo código na tela
    Codigo.Text = CStr(iCodigo)
        
    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 99516
            'Erro tratado na rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174566)
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Função que Inicializa a Gravação de Novo Registro

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chamada da Função Gravar Registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 99499
    
    'Limpa a Tela
    Call Limpa_Tela_TecladoProduto
     
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoGravar_Click:
    
    Select Case gErr
            
        Case 99499
            'Erro Tratada dentro da Função Chamadora
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174567)

    End Select

    Exit Sub
    
End Sub
             
Function Gravar_Registro() As Long
'Função de Gravação de TecladoProduto

Dim objTecladoProduto As New ClassTecladoProduto
Dim lErro As Long

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o campo Código esta preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 99505
    
    'Verifica se o campo descricao esta preenchido
    If Len(Trim(Descricao.Text)) = 0 Then gError 99506
        
    'Verifica se existe algum teclado selecionado
    If ComboTeclado.ListIndex = -1 Then gError 99507
    
    'Move para a memória os campos da Tela
    lErro = Move_Tela_Memoria(objTecladoProduto)
    If lErro <> SUCESSO Then gError 99508
    
    'Se não tem nenhum produto
    If objTecladoProduto.colTecladoProdutoItem.Count = 0 Then gError 99576
    
    'Utilização para incluir FilialEmpresa como parâmetro
    lErro = Trata_Alteracao(objTecladoProduto, objTecladoProduto.iCodigo, objTecladoProduto.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 99509
    
    If objTecladoProduto.colTecladoProdutoItem.Count = 0 Then gError 99576
    
    'Chama a Função que Grava TecladoProduto na Tabela
    lErro = CF("TecladoProduto_Grava", objTecladoProduto)
    If lErro <> SUCESSO Then gError 99510
        
    Gravar_Registro = SUCESSO
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function
    
Erro_Gravar_Registro:
   
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = gErr
        
        Select Case gErr
            
            Case 99505
                lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
                
            Case 99506
                lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)
            
            Case 99507
                lErro = Rotina_Erro(vbOKOnly, "ERRO_TECLADO_NAO_SELECIONADO", gErr)
            
            Case 99508, 99509, 99510
                'Erro Tratado Dentro da Função
                    
            Case 99576
                lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_ESCOLHIDO", gErr)
                    
            Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174568)

        End Select
        
    Exit Function
    
End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objTecladoProduto As New ClassTecladoProduto
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se o Codigo Está Preenchido senão Erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 99500
    
    'Para Saber qual é a FilialEmpresa que Esta sendo Referenciada
    objTecladoProduto.iFilialEmpresa = giFilialEmpresa
    
    'Passa o codigo para a leitura no banco de dados
    objTecladoProduto.iCodigo = Codigo.Text
    
    'Lê a TecladoProduto no Banco e Trazer o objTecladoProduto
    lErro = CF("TecladoProduto_Le", objTecladoProduto)
    If lErro <> SUCESSO And lErro <> 99526 Then gError 99501
    
    'Se não for encontrado a Teclado no Bd
    If lErro = 99526 Then gError 99502
    
    'Envia aviso perguntando se realmente deseja excluir TecladoProduto
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_TECLADOPRODUTO", objTecladoProduto.iCodigo)

    If vbMsgRes = vbYes Then
        
        'Se o teclado a ser excluído é o padrão -->erro
        If objTecladoProduto.iPadrao = TECLADO_PADRAO Then
                   
            'Verifica se exite outro tecladoproduto para esta marca de teclado
            lErro = CF("Teclado_Verifica", objTecladoProduto)
            If lErro <> SUCESSO And lErro <> 109820 Then gError 109821
            
            'Se existe --> não pode ser excluído
            If lErro = 109820 Then gError 99543
            
        End If
        
        'Exclui Teclado
        lErro = CF("Tecladoproduto_Exclui", objTecladoProduto)
        If lErro <> SUCESSO Then gError 99503
        
    End If
    
    'Limpa a Tela
    Call Limpa_Tela_TecladoProduto
    
    'Fechar o Comando de Setas
    Call ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
    Exit Sub
        
Erro_BotaoExcluir_Click:

    Select Case gErr
    
        Case 99500
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 99501, 99503, 109821
            'Erro Tratado Dentro da Função Chamadora
        
        Case 99502
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TECLADOPRODUTO_NAO_ENCONTRADO", gErr, objTecladoProduto.iCodigo)
    
        Case 99543
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TECLADO_PADRAO", gErr, objTecladoProduto.iCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174569)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoLimpar_Click()
'Função que tem as chamadas para as Funções que limpam a tela
Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click
                    
    'Verifica se existe algo para ser salvo antes de limpar a tela
    Call Teste_Salva(Me, iAlterado)
    
    Call Limpa_Tela_TecladoProduto
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then gError 99504
    
    iAlterado = 0
    
    Exit Sub
        
Erro_Botaolimpar_Click:

    Select Case gErr
    
        Case 99504
            'Erro Tratado dentro da Função Chamadora
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174570)

    End Select
    
    Exit Sub
        
End Sub

Sub Limpa_Tela_TecladoProduto()
    
    'Limpa Tela
    Call Limpa_Tela(Me)
    
    DescProduto.Caption = ""
    ComboTeclado.ListIndex = -1
    BotaoModelo.Caption = ""
    BotaoModelo.BackColor = COR_DEFAULT
    BotaoModelo.Tag = ""
    Fundo.SelectedColor = COR_DEFAULT
    Padrao.Value = TECLADO_COMUM
    TvwMenu.Nodes.Clear
    
    RemoveBuracos.Value = vbUnchecked
        
    'Limpa os botões do array
    Call Inicializa_ColTecladoProdutoItens
    
    
End Sub
        
Private Sub BotaoModelo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Se fizer mouse down e estiver com MousePointer de Size, começa um Drag, usando método Drag e através do evento
'DragOver e conforme o tipo de MousePointer vai alterando o tamanho do botão (largura, altura ou os dois) até


'o evento DragDrop acontecer quando encerra o processo. Faz o espelho acompanhar as alterações de tamanho.

'Se fizer mouse down e MousePointer estiver normal, começa um Drag, com método Drag e através do evento
'DragOver vai arrastar o BotaoEspelho seguindo o movimento do mouse até o evento DragDrop acontecer.
'SE estiver dentro do teclado, copia o botão para lá. Em qualquer caso volta com o Espelho para trás do
'original.

End Sub


Private Sub BotaoOrganiza_Click()
'Arruma os botoes do teclado para ficarem com o tamanho do último que foi inserido ou alterado o tamanho.
'Arruma as posicoes para ficarem igualmente espacados verticalmente e horizontalmente. Os botoes devem ser
'arrumados por linha horizontal baixando pouco a pouco. Faz varreduras horizontais com o espaçamento entre
'varreduras sendo METADE da altura dos botoes. Pega só os que tiverem com seu centro ACIMA da varredura.

End Sub

Private Sub FrameBotaoModelo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Nesse MouseMove verifica se está na borda do BotaoModelo. Se estiver muda MousePointer para SizeEW ou SizeNS ou Size conforme
'esteja nos lados ou nas bordas de topo e de baixo. Quando sair da borda volta a Mouse Pointer normal.

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Libera a referência da tela e fecha o comando das setas se estiver aberto
    Set objEventoProduto = Nothing
    Set objEventoTecladoProduto = Nothing
    Set objEventoTeclado = Nothing
    Set gColTecladoProdutoItens = Nothing
    
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub


Private Sub Produto_Click(index As Integer)
    
Dim objTecladoProdutoItens As New ClassTecladoProdutoItem
    
    'Preenche o botão modelo com as características do produto selecionado
    BotaoModelo.BackColor = Produto(index).BackColor
    BotaoModelo.Caption = Produto(index).Caption
    BotaoModelo.Tag = index
    Fundo.SelectedColor = Produto(index).BackColor
    
    'se esta relacionado com item da arvore, portanto o botao ja tiver sido iniciado
    If Len(Produto(index).Tag) > 0 Then
    
        TvwMenu.Nodes.Item(Produto(index).Tag).Selected = True
    
        'inicializa o obj com os dados armazenado na váriavel global para este índice
        Set objTecladoProdutoItens = gColTecladoProdutoItens.Item(Produto(index).Tag)
    
        'Jona na tela os dados do produto selecionado
        If objTecladoProdutoItens.iTecla <> 0 Then
            Tecla.Text = CStr(objTecladoProdutoItens.iTecla)
        Else
            Tecla.Text = ""
        End If
    
        CodProduto.PromptInclude = False
        CodProduto.Text = objTecladoProdutoItens.sProduto
        CodProduto.PromptInclude = True
    
        Call CodProduto_Validate(False)
    
        Titulo.Text = objTecladoProdutoItens.sTitulo
    
    Else
    
        CodProduto.PromptInclude = False
        CodProduto.Text = ""
        CodProduto.PromptInclude = True
        Titulo.Text = ""
        Tecla.Text = ""
        DescProduto.Caption = ""
    
    End If
    
End Sub

Private Sub BotaoAplicar_Click()
    
Dim objTecladoProdutoItens As ClassTecladoProdutoItem
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim bConfigura As Boolean
Dim sTecla As String

On Error GoTo Erro_BotaoAplicar_Click
    
    'Se não tem botão selecionado -->erro
    If Len(Trim(BotaoModelo.Tag)) = 0 Then gError 99575
    
'    bConfigura = False
    
    'se o titulo nao esta preenchido
    If Len(Trim(Titulo.Text)) = 0 Then
    
        'se o botao nao tem uma key associada
        If Len(Trim(Produto(BotaoModelo.Tag).Tag)) = 0 Then
            gError 99517
    
        Else
        
            'manda aviso dizendo que o botão vai ser desconfigurado
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DESCONFIGURAR_TECLADO", Produto(BotaoModelo.Tag).Caption)
'            'Se não deseja -->sair
'            If vbMsgRes = vbNo Then Exit Sub
    
            Call BotaoRemover_Click
    
'            CodProduto.PromptInclude = False
'            CodProduto.Text = ""
'            CodProduto.PromptInclude = True
'            Call CodProduto_Validate(False)
'
'            Tecla.Text = ""
'            Titulo.Text = ""
'            Fundo.SelectedColor = COR_DEFAULT
'
'            Produto(BotaoModelo.Tag).Caption = ""
'            Produto(BotaoModelo.Tag).BackColor = COR_DEFAULT
'
'            gColTecladoProdutoItens.Remove (Produto(BotaoModelo.Tag).Tag)
'            TvwMenu.Nodes.Remove (Produto(BotaoModelo.Tag).Tag)
'            Produto(BotaoModelo.Tag).Tag = ""
'
'            'limpo o botão modelo
'            BotaoModelo.BackColor = COR_DEFAULT
'            BotaoModelo.Caption = ""
'            BotaoModelo.Tag = ""
    
        End If
    
    Else
    
        'se o titulo esta preenchido
        'atualiza ou cria um novo botao
        lErro = Atualiza_gColTecladoProdutoItens(gColTecladoProdutoItens, BotaoModelo.Tag)
        If lErro <> SUCESSO Then gError 99519
        
        'Preenche o produto com as características do botão selecionado
        Produto(BotaoModelo.Tag).BackColor = Fundo.SelectedColor
        
    '    Produto(BotaoModelo.Tag).Caption = Titulo.Text
            
        Set objTecladoProdutoItens = gColTecladoProdutoItens(Produto(BotaoModelo.Tag).Tag)
            
        If objTecladoProdutoItens.sTitulo <> "" Then
                
            If objTecladoProdutoItens.iTecla <> 0 Then
                If objTecladoProdutoItens.iTecla >= vbKeyF2 Then
                    Call Acha_Tecla(objTecladoProdutoItens.iTecla, sTecla)
                Else
                    sTecla = Chr(objTecladoProdutoItens.iTecla)
                End If
                Produto(BotaoModelo.Tag).Caption = objTecladoProdutoItens.sTitulo & "(" & sTecla & ")"
            Else
                Produto(BotaoModelo.Tag).Caption = objTecladoProdutoItens.sTitulo
            End If
        Else
            Produto(BotaoModelo.Tag).Caption = ""
        End If
        
    End If
        
    iDadosBotaoAlterado = 0
        
    Exit Sub
        
Erro_BotaoAplicar_Click:

    Select Case gErr
        
        Case 99517
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULO_NAO_PREENCHIDO1", gErr)
            
        Case 99519
            'Erro Tratado dentro da Função Chamadora
        
        Case 99575
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BOTAO_NAO_SELECIONADO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174571)

    End Select
    
    Exit Sub
        
End Sub

Function Atualiza_gColTecladoProdutoItens(ColTecladoProdutoItens As Collection, ByVal index As Integer) As Long

Dim objTecladoProdutoItens As ClassTecladoProdutoItem
Dim iIndice As Integer
Dim sServ As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long
Dim objNode As Node
Dim sKey As String
Dim iRelative As Integer

On Error GoTo Erro_Atualiza_gColTecladoProdutoItens
    
    If CodProduto.ClipText <> "" Then
        
        lErro = CF("Produto_Formata", CodProduto.Text, sServ, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 99571

    End If

    If StrParaInt(Tecla.Text) = vbKeyF2 Or StrParaInt(Tecla.Text) = vbKeyF3 Or StrParaInt(Tecla.Text) = vbKeyF4 _
    Or StrParaInt(Tecla.Text) = vbKeyF5 Or StrParaInt(Tecla.Text) = vbKeyF6 Or StrParaInt(Tecla.Text) = vbKeyF7 _
    Or StrParaInt(Tecla.Text) = vbKeyF8 Or StrParaInt(Tecla.Text) = vbKeyF9 Or StrParaInt(Tecla.Text) = vbKeyF10 _
    Or StrParaInt(Tecla.Text) = vbKeyF11 Or StrParaInt(Tecla.Text) = vbKeyEscape Or StrParaInt(Tecla.Text) = vbKeyF1 Then gError 112534


    sKey = Produto(index).Tag

    'se for uma alteracao
    If Len(sKey) > 0 Then

        
        lErro = Valida_Elemento_Duplicado(sKey, ColTecladoProdutoItens)
        If lErro <> SUCESSO Then gError 214910
        
        Set objTecladoProdutoItens = ColTecladoProdutoItens(sKey)
        
        'Joga no obj os novos dados do produto selecionado
        objTecladoProdutoItens.iTecla = StrParaInt(Tecla.Text)
        objTecladoProdutoItens.sProduto = sServ
        objTecladoProdutoItens.sTitulo = Titulo.Text
        objTecladoProdutoItens.lColor = Fundo.SelectedColor
        
        TvwMenu.Nodes.Item(sKey).Text = Titulo.Text
        Produto(index).Caption = Titulo.Text


    Else
    

        'se o elemento ainda nao testa cadastrado
        
        Set objTecladoProdutoItens = New ClassTecladoProdutoItem
        
        'Joga no obj os novos dados do produto selecionado
        objTecladoProdutoItens.iTecla = StrParaInt(Tecla.Text)
        objTecladoProdutoItens.sProduto = sServ
        objTecladoProdutoItens.sTitulo = Titulo.Text
        objTecladoProdutoItens.lColor = Fundo.SelectedColor
        
        Set objNode = TvwMenu.SelectedItem
        
        'se a arvore esta vazia
        If objNode Is Nothing Then
            objTecladoProdutoItens.sArvoreKey = "x" & Format(index + 1, "00")
            
            lErro = Valida_Elemento_Duplicado(objTecladoProdutoItens.sArvoreKey, ColTecladoProdutoItens)
            If lErro <> SUCESSO Then gError 214911
            
            Set objNode = TvwMenu.Nodes.Add(, , objTecladoProdutoItens.sArvoreKey, Titulo.Text)
            TvwMenu.SelectedItem = objNode
        Else
            If Len(objNode.Key) > 3 Then
                objTecladoProdutoItens.sArvoreKey = Left(objNode.Key, Len(objNode.Key) - 2) & Format(index + 1, "00")
                
                lErro = Valida_Elemento_Duplicado(objTecladoProdutoItens.sArvoreKey, ColTecladoProdutoItens)
                If lErro <> SUCESSO Then gError 214912
                
                If objNode.Text = "<Vazio>" Then TvwMenu.Nodes.Remove (objNode.Key)
                    
                Call Procura_Irmao(objTecladoProdutoItens.sArvoreKey, objNode, iRelative)
                    
                Set objNode = TvwMenu.Nodes.Add(objNode, iRelative, objTecladoProdutoItens.sArvoreKey, Titulo.Text)
                TvwMenu.SelectedItem = objNode
                    
                    
            Else
                objTecladoProdutoItens.sArvoreKey = "x" & Format(index + 1, "00")
                    
                lErro = Valida_Elemento_Duplicado(objTecladoProdutoItens.sArvoreKey, ColTecladoProdutoItens)
                If lErro <> SUCESSO Then gError 214913
                    
                Call Procura_Irmao(objTecladoProdutoItens.sArvoreKey, objNode, iRelative)
                    
                Set objNode = TvwMenu.Nodes.Add(objNode, iRelative, objTecladoProdutoItens.sArvoreKey, Titulo.Text)
                TvwMenu.SelectedItem = objNode
            
            End If
            
        End If

        ColTecladoProdutoItens.Add objTecladoProdutoItens, objTecladoProdutoItens.sArvoreKey

        'Joga no obj os novos dados do produto selecionado
        objTecladoProdutoItens.iTecla = StrParaInt(Tecla.Text)
        objTecladoProdutoItens.sProduto = sServ
        objTecladoProdutoItens.sTitulo = Titulo.Text
        objTecladoProdutoItens.lColor = Fundo.SelectedColor
        
        Produto(index).Tag = objTecladoProdutoItens.sArvoreKey


   End If
  
    Atualiza_gColTecladoProdutoItens = SUCESSO
    
    Exit Function
        
Erro_Atualiza_gColTecladoProdutoItens:
    
    Atualiza_gColTecladoProdutoItens = gErr
    
    Select Case gErr
        
        Case 99518
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_SELECIONADO", gErr)
            
        Case 99571
        
        Case 214910, 214911, 214912, 214913
            Call Rotina_Erro(vbOKOnly, "ERRO_TECLA_JA_UTILIZADA", gErr, objTecladoProdutoItens.iTecla, objTecladoProdutoItens.sTitulo)
        
        Case 112534
            Call Rotina_Erro(vbOKOnly, "ERRO_TECLA_NAO_DEVE_UTILIZAR", gErr, Tecla.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174572)

    End Select
    
End Function

Sub Preenche_ColTecladoProdutoItens(objTecladoProduto As ClassTecladoProduto)

Dim objTecladoProdutoItens As ClassTecladoProdutoItem
Dim iIndice As Integer
Dim sServ As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long
Dim colSaida As New Collection, colCampos As New Collection
    
On Error GoTo Erro_Preenche_ColTecladoProdutoItens

    If RemoveBuracos.Value = vbChecked Then

        colCampos.Add "sArvoreKey"
    
        Call Ordena_Colecao(gColTecladoProdutoItens, colSaida, colCampos)
    
        lErro = Acerta_ArvoreKey(colSaida)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    Else
        Set colSaida = gColTecladoProdutoItens
    End If

    For iIndice = 1 To colSaida.Count
        Set objTecladoProdutoItens = colSaida(iIndice)
            
'        If Len(Trim(objTecladoProdutoItens.sProduto)) <> 0 Then
        objTecladoProduto.colTecladoProdutoItem.Add objTecladoProdutoItens
'        End If
    Next
    
     Exit Sub
    
Erro_Preenche_ColTecladoProdutoItens:
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174572)

    End Select
    
    Exit Sub
    
End Sub

Private Function Acerta_ArvoreKey(ByVal colItens As Collection) As Long

Dim Indice As Integer
Dim objItem As ClassTecladoProdutoItem, objItemAnt As ClassTecladoProdutoItem
Dim iLv As Integer, iLvAnt As Integer, sArvoreKeyIrmao As String
    
On Error GoTo Erro_Acerta_ArvoreKey

    Indice = 0
    For Each objItem In colItens
        Indice = Indice + 1
        iLv = (Len(objItem.sArvoreKey) - 1) / 2
        
        If Indice = 1 Then
            objItem.sArvoreKey = "x01"
        Else
            iLvAnt = (Len(objItemAnt.sArvoreKey) - 1) / 2
            
            If iLvAnt < iLv Then 'Filho do anterior
                objItem.sArvoreKey = objItemAnt.sArvoreKey & "01"
            ElseIf iLvAnt = iLv Then 'Irmão do anterior
                objItem.sArvoreKey = Left(objItemAnt.sArvoreKey, Len(objItemAnt.sArvoreKey) - 2) & Format(StrParaInt(Right(objItemAnt.sArvoreKey, 2)) + 1, "00")
            Else 'Desconhecido, pode ser irmão do pai ou do avó ou qualquer ascendente
                sArvoreKeyIrmao = Left(objItemAnt.sArvoreKey, 1 + (iLv * 2)) 'Obtém irmão
                objItem.sArvoreKey = Left(sArvoreKeyIrmao, Len(sArvoreKeyIrmao) - 2) & Format(StrParaInt(Right(sArvoreKeyIrmao, 2)) + 1, "00")
            End If
        
        End If
        
        Set objItemAnt = objItem
    Next
    
    Acerta_ArvoreKey = SUCESSO
    
    Exit Function
    
Erro_Acerta_ArvoreKey:
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174572)

    End Select
    
    Exit Function
    
End Function

Sub Traz_gColTecladoProdutoItens_Nivel1()

Dim objTecladoProdutoItens As ClassTecladoProdutoItem
Dim iIndice As Integer
Dim sTecla As String
    
    'Pra cada item do TecladoProduto eu jogo seu correnspondente na tela
    For Each objTecladoProdutoItens In gColTecladoProdutoItens
    
        If Len(objTecladoProdutoItens.sArvoreKey) = 3 Then
            
            iIndice = StrParaInt(Right(objTecladoProdutoItens.sArvoreKey, 2)) - 1
    
            Produto(iIndice).BackColor = objTecladoProdutoItens.lColor
            Produto(iIndice).Tag = objTecladoProdutoItens.sArvoreKey
                
            If objTecladoProdutoItens.sTitulo <> "" Then
                If objTecladoProdutoItens.iTecla <> 0 Then
                    If objTecladoProdutoItens.iTecla >= vbKeyF2 Then
                        Call Acha_Tecla(objTecladoProdutoItens.iTecla, sTecla)
                    Else
                        sTecla = Chr(objTecladoProdutoItens.iTecla)
                    End If
                    Produto(iIndice).Caption = objTecladoProdutoItens.sTitulo & "(" & sTecla & ")"
                Else
                    Produto(iIndice).Caption = objTecladoProdutoItens.sTitulo
                End If
            Else
                Produto(iIndice).Caption = ""
            End If
                
        End If
                
    Next
    
End Sub

Sub Traz_ColTecladoProdutoItens(ColTecladoProdutoItens As Collection)

Dim objTecladoProdutoItens As ClassTecladoProdutoItem
Dim objTecladoProdutoItens1 As ClassTecladoProdutoItem
Dim iIndice As Integer

    
    
    
    'Inicializa a col global com os dados default
    Call Inicializa_ColTecladoProdutoItens
    
'    For iIndice = 1 To ColTecladoProdutoItens.Count
'
'        Set objTecladoProdutoItens = ColTecladoProdutoItens(iIndice)
'
'        Set objTecladoProdutoItens1 = gColTecladoProdutoItens(objTecladoProdutoItens.iIndice + 1)
'
'        objTecladoProdutoItens1.sProduto = objTecladoProdutoItens.sProduto
'        objTecladoProdutoItens1.iIndice = objTecladoProdutoItens.iIndice
'        objTecladoProdutoItens1.iTecla = objTecladoProdutoItens.iTecla
'        objTecladoProdutoItens1.lColor = objTecladoProdutoItens.lColor
'        objTecladoProdutoItens1.sTitulo = objTecladoProdutoItens.sTitulo
'
'    Next
    
    Set gColTecladoProdutoItens = ColTecladoProdutoItens
    
    Call Traz_gColTecladoProdutoItens_Nivel1
    
    Call Carrega_Arvore(ColTecladoProdutoItens)
    
End Sub

Private Sub Tecla_Change()
        
    iAlterado = REGISTRO_ALTERADO
    
End Sub
    
Private Sub Fundo_Change(NewColor As Long)

    BotaoModelo.BackColor = NewColor
    iDadosBotaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Tecla_KeyDown(KeyCode As Integer, Shift As Integer)

'Dim vbMsgBox As VbMsgBoxResult
'
'    If BotaoModelo.Tag = "" Then
'        'Envia aviso para selecionar um produto
'        vbMsgBox = Rotina_Aviso(vbOK, "AVISO_PRODUTO_SELECIONADO")
'        Exit Sub
'    End If
'
'    'Joga na tela a tecla acionada
'    Tecla.Text = KeyCode

End Sub

Private Sub Tecla_KeyPress(KeyAscii As Integer)

'Dim vbMsgBox As VbMsgBoxResult
'
'    'Se não tem botão selecionado --> Aviso
'   If BotaoModelo.Tag = "" Then
'        'Envia aviso para selecionar um produto
'        vbMsgBox = Rotina_Aviso(vbOK, "AVISO_PRODUTO_SELECIONADO")
'        Exit Sub
'    End If
'
'    'Joga na tela a tecla acionada
'    Tecla.Text = KeyAscii

End Sub

Private Sub Tecla_KeyUp(KeyCode As Integer, Shift As Integer)

    'Joga na tela a tecla acionada
    Tecla.Text = KeyCode

End Sub

Private Sub Tecla_Validate(Cancel As Boolean)
    
Dim vbMsgBox As VbMsgBoxResult
Dim iTecla As Integer
Dim sTecla As String

    'Se não tem botão selecionado --> Aviso
   If Len(Trim(BotaoModelo.Tag)) = 0 Then
        'Envia aviso para selecionar um produto
        vbMsgBox = Rotina_Aviso(vbOK, "AVISO_PRODUTO_SELECIONADO")
        Tecla.Text = ""
        Exit Sub
    End If

    
    If IsNumeric(Tecla.Text) Then
        If StrParaInt(Tecla.Text) = 0 Then Tecla.Text = ""
    Else
        Tecla.Text = ""
    End If
    
    iTecla = StrParaInt(Tecla.Text)
    
    If Len(Titulo.Text) > 0 Then
        If iTecla <> 0 Then
        
            If iTecla >= vbKeyF2 Then
                Call Acha_Tecla(iTecla, sTecla)
            Else
                sTecla = Chr(iTecla)
            End If
        
            BotaoModelo.Caption = Titulo.Text & "(" & sTecla & ")"
        Else
            BotaoModelo.Caption = Titulo.Text
        End If
    Else
        BotaoModelo.Caption = ""
    End If
    
End Sub

Private Sub Titulo_Change()
'Determina se Houve Mudança

    iAlterado = REGISTRO_ALTERADO
    iDadosBotaoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Titulo_Validate(Cancel As Boolean)

Dim vbMsgBox As VbMsgBoxResult
Dim objTecladoProdutoItens As ClassTecladoProdutoItem
Dim sTecla As String


    'Se não tem botão selecionado -->sai
    If Len(Trim(BotaoModelo.Tag)) = 0 And Len(Trim(Titulo.Text)) > 0 Then
        'Envia aviso para selecionar um produto
        vbMsgBox = Rotina_Aviso(vbOK, "AVISO_PRODUTO_SELECIONADO")
        Titulo.Text = ""
        Exit Sub
    End If
    
    If Len(Titulo.Text) > 0 Then
        If Len(Tecla.Text) <> 0 Then
        
            If StrParaInt(Tecla.Text) >= vbKeyF2 Then
                Call Acha_Tecla(StrParaInt(Tecla.Text), sTecla)
            Else
                sTecla = Chr(StrParaInt(Tecla.Text))
            End If
        
            BotaoModelo.Caption = Titulo.Text & "(" & sTecla & ")"
        Else
            BotaoModelo.Caption = Titulo.Text
        End If
    Else
        BotaoModelo.Caption = ""
    End If
    
    
End Sub

Private Sub Codigo_Change()
'Determina se Houve Mudança

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Descricao_Change()
'Determina se Houve Mudança

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Padrao_Click()
'Determina se Houve Mudança

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComboTeclado_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objTeclado As New ClassTeclado

On Error GoTo Erro_ComboPais_Validate
    
    'Verifica se foi preenchida a Combo
    If Len(Trim(ComboTeclado.Text)) = 0 Then Exit Sub
    
    'Verifica se está preenchida com o ítem selecionado na Combo
    If ComboTeclado.Text = ComboTeclado.List(ComboTeclado.ListIndex) Then Exit Sub
    
    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(ComboTeclado, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 99496

    'Nao existe o item com o CODIGO na List da ComboBox
    If lErro = 6730 Then

        objTeclado.iCodigo = iCodigo

        'Tenta ler Teclado com esse codigo no BD
        lErro = CF("Teclado_Le", objTeclado)
        If lErro <> SUCESSO And lErro <> 99459 Then gError 99497
        
        'Se não achou
        If lErro <> SUCESSO Then
            
            'pergunta se deseja cadastrar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TECLADO", objTeclado.iCodigo)
            
            'Se confirma
            If vbMsgRes = vbYes Then
                Call Chama_Tela("Teclado", objTeclado)
            End If

        End If
        
        'Joga o Teclado na tela
        ComboTeclado.Text = CStr(iCodigo) & SEPARADOR & objTeclado.sDescricao

    End If

    'Nao existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 99498

    Exit Sub

Erro_ComboPais_Validate:

    Cancel = True

    Select Case gErr

        Case 99496, 99497

        Case 99498
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TECLADO_NAO_CADASTRADO", gErr, ComboTeclado.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174573)

    End Select

    Exit Sub

End Sub

Private Sub CodProduto_Change()

    iAlterado = REGISTRO_ALTERADO
    iProdutoAlterado = REGISTRO_ALTERADO
    iDadosBotaoAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub CodProduto_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodProduto, iAlterado)
    iProdutoAlterado = 0

End Sub

Private Sub CodProduto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sServ As String
Dim vbMsgBox As VbMsgBoxResult

On Error GoTo Erro_CodProduto_Validate
    
    'Se não tem botão selecionado -->sai
    If Len(Trim(BotaoModelo.Tag)) = 0 And Len(Trim(CodProduto.ClipText)) > 0 Then

        CodProduto.PromptInclude = False
        CodProduto.Text = ""
        CodProduto.PromptInclude = True

        DescProduto.Caption = ""

        'Envia aviso para selecionar um produto
        vbMsgBox = Rotina_Aviso(vbOK, "AVISO_PRODUTO_SELECIONADO")

        Exit Sub

    End If
    
    'Se o Produto não foi alterado
    If iProdutoAlterado = 0 Then Exit Sub
    
    If Len(Trim(CodProduto.ClipText)) = 0 Then
        DescProduto.Caption = ""
        Exit Sub
    End If
    
    lErro = CF("Produto_Formata", CodProduto.Text, sServ, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 99491
    
    objProduto.sCodigo = sServ
    
    'Verifica se o produto existe
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 99492
        
    'Produto não existente
    If lErro = 28030 Then gError 99493
    
    'Verifica se é de Faturamento, se não for --> Erro
    If objProduto.iFaturamento = PRODUTO_NAO_VENDAVEL Then gError 99494
            
    'Verifica se é gerencial, se for --> Erro
    If objProduto.iGerencial = PRODUTO_GERENCIAL And Len(Trim(objProduto.sGrade)) = 0 And objProduto.iKitVendaComp <> MARCADO Then gError 99495
    
    'Coloca a descrição na tela
    DescProduto.Caption = objProduto.sDescricao
        
    iProdutoAlterado = 0
    
    Exit Sub

Erro_CodProduto_Validate:

    Cancel = True

    Select Case gErr
        
        Case 99491, 99492
        
        Case 99493
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 99494
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_SER_VENDIDO2", gErr, objProduto.sCodigo)
        
        Case 99495
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174574)
                        
    End Select
   
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Private Sub BotaoFechar_Click()
'Função que Fecha a Tela

    Unload Me

End Sub

'Function TecladoProduto_Desmembra_Log(objTecladoProduto As ClassTecladoProduto, objLog As ClassLog) As Long
''Função que informações do banco de Dados e Carrega no Obj
'
'Dim lErro As Long
'Dim iPosicao1 As Integer
'Dim iPosicao2 As Integer
'Dim iPosicao3 As Integer
'Dim iPosicao4 As Integer
'Dim iPosicao5 As Integer
'Dim stecladoproduto As String
'Dim iIndice As Integer
'Dim bFim As Boolean
'Dim objTecladoProdutoItens As ClassTecladoProdutoItem
'
'On Error GoTo Erro_tecladoproduto_Desmembra_Log
'
'    'iPosicao1 Guarda a posição do Primeiro Control
'    iPosicao1 = InStr(1, objLog.sLog, Chr(vbKeyControl))
'
'    'String que Guarda as Propriedades do Objtecladoproduto
'    stecladoproduto = Mid(objLog.sLog, 1, iPosicao1 - 1)
'
'    'iPosicao4 Guarda o Final da String
'    iPosicao4 = InStr(1, objLog.sLog, Chr(vbKeyEnd))
'
'    'Variável booleana que funcionará como Flag
'    bFim = True
'    'Inicilalização do objtecladoproduto
'    Set objTecladoProduto = New ClassTecladoProduto
'
'    'Primeira Posição
'    iPosicao3 = 1
'
'    'Procura o Primeiro Escape dentro da String stecladoproduto e Armazena a Posição
'    iPosicao2 = (InStr(iPosicao3, stecladoproduto, Chr(vbKeyEscape)))
'
'    iIndice = 0
'
'    stecladoproduto = stecladoproduto & Chr(vbKeyEscape)
'
'    Do While iPosicao2 <> 0
'
'       iIndice = iIndice + 1
'        'Recolhe os Dados do Banco de Dados e Coloca no objtecladoproduto
'        Select Case iIndice
'            Case 1: objTecladoProduto.iCodigo = StrParaInt(Mid(stecladoproduto, iPosicao3, iPosicao2 - iPosicao3))
'            Case 2: objTecladoProduto.iFilialEmpresa = StrParaInt(Mid(stecladoproduto, iPosicao3, iPosicao2 - iPosicao3))
'            Case 3: objTecladoProduto.iPadrao = StrParaLong(Mid(stecladoproduto, iPosicao3, iPosicao2 - iPosicao3))
'            Case 4: objTecladoProduto.iTeclado = StrParaInt(Mid(stecladoproduto, iPosicao3, iPosicao2 - iPosicao3))
'            Case 5: objTecladoProduto.sDescricao = Mid(stecladoproduto, iPosicao3, iPosicao2 - iPosicao3)
'            Case 6: Exit Do
'
'        End Select
'        'Atualiza as Posições
'        iPosicao3 = iPosicao2 + 1
'        iPosicao2 = (InStr(iPosicao3, stecladoproduto, Chr(vbKeyEscape)))
'
'    Loop
'
'    iPosicao3 = iPosicao1 + 1
'
'
'
'    Do While bFim <> False
'
'        'iPosicao1 Guarda a posição do Control Ponto Inicial
'        iPosicao1 = InStr(iPosicao3, objLog.sLog, Chr(vbKeyControl))
'
'        If iPosicao1 = 0 Then iPosicao1 = iPosicao4
'
'        'Atualiza as Posições
'        iPosicao2 = (InStr(iPosicao3, objLog.sLog, Chr(vbKeyEscape)))
'
'        'Atualiza o valor de Indice
'        iIndice = 0
'
'        'inicia o objtecladoprodutoCondPagto para receber os dados do Banco de Dados
'        Set objTecladoProdutoItens = New ClassTecladoProdutoItem
'
'        Do While iPosicao2 > iPosicao3
'
'            iIndice = iIndice + 1
'
'            'Recolhe os Dados do Banco de Dados e Coloca no objtecladoprodutoitens
'            Select Case iIndice
'
'                Case 1: objTecladoProdutoItens.iIndice = StrParaInt(Mid(objLog.sLog, iPosicao3, iPosicao2 - (iPosicao3)))
'                Case 2: objTecladoProdutoItens.iTecla = StrParaInt(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
'                Case 3: objTecladoProdutoItens.lColor = StrParaLong(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
'                Case 4: objTecladoProdutoItens.sProduto = Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3)
'                Case 5: objTecladoProdutoItens.sTitulo = Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3)
'                Case 6: Exit Do
'
'            End Select
'
'            'Atualiza as Posições
'            iPosicao3 = iPosicao2 + 1
'            iPosicao2 = (InStr(iPosicao3, objLog.sLog, Chr(vbKeyEscape)))
'
'            'Verfica se a Posição que é atualizada no decorrer da função é maior que a posição da Flag End
'            If (iPosicao2 > iPosicao1) Or iPosicao2 = 0 Then
'                'A flag Fim Recebe False
'                iPosicao2 = iPosicao1
'            End If
'
'        Loop
'
'        objTecladoProduto.colTecladoProdutoItem.Add objTecladoProdutoItens
'
'        'Verfica se a Posição que é atualizada no decorrer da função é maior que a posição da Flag End
'        If iPosicao3 > iPosicao4 Then
'            'A flag Fim Recebe False
'            bFim = False
'        End If
'
'    Loop
'
'    TecladoProduto_Desmembra_Log = SUCESSO
'
'    Exit Function
'
'Erro_tecladoproduto_Desmembra_Log:
'
'    TecladoProduto_Desmembra_Log = gErr
'
'   Select Case gErr
'
'        Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174575)
'
'        End Select
'
'
'    Exit Function
'
'End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD
Dim lErro As Long
Dim objTecladoProduto As New ClassTecladoProduto

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TecladoProduto"

    'Le os dados da Tela TecladoProduto
    lErro = Move_Tela_Memoria(objTecladoProduto)
    If lErro <> SUCESSO Then gError 99489

    'Preenche a coleção colCampoValor, com descricao do campo,
    colCampoValor.Add "Codigo", objTecladoProduto.iCodigo, 0, "Codigo"
    colCampoValor.Add "descricao", objTecladoProduto.sDescricao, STRING_TECLADOPRODUTO_DESCRICAO, "Descricao"
    colCampoValor.Add "FilialEmpresa", objTecladoProduto.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "Teclado", objTecladoProduto.iTeclado, 0, "Teclado"
    colCampoValor.Add "Padrao", objTecladoProduto.iPadrao, 0, "Padrao"
    
    'Utilizado na hora de passar o parâmetro FilialEmpresa
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr
        'Erro tratado na rotina chamadora
        Case 99489
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174576)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD
Dim lErro As Long
Dim objTecladoProduto As New ClassTecladoProduto

On Error GoTo Erro_Tela_Preenche

    objTecladoProduto.iCodigo = colCampoValor.Item("Codigo").vValor
            
    If objTecladoProduto.iCodigo > 0 Then
        
        'Carrega o TecladoProduto com os dados passados em colCampoValor
        objTecladoProduto.sDescricao = colCampoValor.Item("Descricao").vValor
        objTecladoProduto.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
        objTecladoProduto.iTeclado = colCampoValor.Item("Teclado").vValor
        objTecladoProduto.iPadrao = colCampoValor.Item("Padrao").vValor
        
        'Traz dados de Teclados para a Tela
        lErro = Traz_TecladoProduto_Tela(objTecladoProduto)
        If lErro <> SUCESSO Then gError 99490
        
    End If
        
    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 99490
        'Erro tratado na rotina chamadora
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174577)

    End Select
    
    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'Clique em F2
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    'Clique em F3
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is CodProduto Then Call LabelProduto_Click
        If Me.ActiveControl Is ComboTeclado Then Call LabelTeclado_Click
        If Me.ActiveControl Is Codigo Then Call LabelCodigo_Click
        
    End If

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Teclado Produto"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "TecladoProduto"
    
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

Public Sub Acha_Tecla(iTecla As Integer, sTecla As String)

    Select Case iTecla
    
        Case vbKeyF2
            sTecla = "F2"
        Case vbKeyF3
            sTecla = "F3"
        Case vbKeyF4
            sTecla = "F4"
        Case vbKeyF5
            sTecla = "F5"
        Case vbKeyF6
            sTecla = "F6"
        Case vbKeyF7
            sTecla = "F7"
        Case vbKeyF8
            sTecla = "F8"
        Case vbKeyF9
            sTecla = "F9"
        Case vbKeyF10
            sTecla = "F10"
        Case vbKeyF11
            sTecla = "F11"
        Case vbKeyF12
            sTecla = "F12"
        Case vbKeyF13
            sTecla = "F13"
        Case vbKeyF14
            sTecla = "F14"
        Case vbKeyF15
            sTecla = "F15"
        Case vbKeyF16
            sTecla = "F16"
           
    End Select
    
End Sub

Sub Procura_Irmao(ByVal sKey As String, objNode As Node, iRelative As Integer)

Dim objNode1 As Node

    iRelative = 0

    If TvwMenu.Nodes.Count = 0 Then
        Set objNode = Nothing
    Else

        If Len(sKey) = 3 Then
            Set objNode1 = TvwMenu.Nodes.Item(1).FirstSibling
            
            Do While objNode1.Key < sKey
                If objNode1.Next Is Nothing Then
                    iRelative = tvwNext
                    Exit Do
                End If
                Set objNode1 = objNode1.Next
            Loop
            
            If iRelative = 0 Then
                iRelative = tvwPrevious
            End If
                    
        Else
        
            Set objNode1 = TvwMenu.Nodes.Item(Left(sKey, Len(sKey) - 2))
            
            If objNode1.Children = 0 Then
                iRelative = tvwChild
            Else
                Set objNode1 = objNode1.Child
                
                Do While objNode1.Key < sKey
                    If objNode1.Next Is Nothing Then
                        iRelative = tvwNext
                        Exit Do
                    End If
                    Set objNode1 = objNode1.Next
                Loop
                
                If iRelative = 0 Then
                    iRelative = tvwPrevious
                End If
                        
            End If
            
        End If
        
        Set objNode = objNode1
                
    End If

End Sub

Private Sub TvwMenu_NodeClick(ByVal Node As MSComctlLib.Node)

Dim iIndice As Integer
Dim objTecladoProdutoItens As ClassTecladoProdutoItem
Dim sTecla As String

    'limpa os botoes
    For iIndice = 0 To Produto.Count - 1
        Produto(iIndice).Caption = ""
        Produto(iIndice).Tag = ""
        Produto(iIndice).BackColor = COR_DEFAULT
    Next

    BotaoModelo.Caption = ""
    BotaoModelo.Tag = ""
    BotaoModelo.BackColor = COR_DEFAULT

    DescProduto.Caption = ""
    Fundo.SelectedColor = COR_DEFAULT

    CodProduto.PromptInclude = False
    CodProduto.Text = ""
    CodProduto.PromptInclude = True
    
    Tecla.Text = ""
    Titulo.Text = ""

    'preenche os botoes de acordo com o que esta armazendo na treeview no nivel selecionado
    For Each objTecladoProdutoItens In gColTecladoProdutoItens
    
        If Len(Node.Key) = 3 Then
            If Len(objTecladoProdutoItens.sArvoreKey) = 3 Then
                
                iIndice = CInt(Right(objTecladoProdutoItens.sArvoreKey, 2)) - 1
                
                If Len(objTecladoProdutoItens.sTitulo) > 0 Then
                    If objTecladoProdutoItens.iTecla <> 0 Then
                
                        If objTecladoProdutoItens.iTecla >= vbKeyF2 Then
                            Call Acha_Tecla(objTecladoProdutoItens.iTecla, sTecla)
                        Else
                            sTecla = Chr(objTecladoProdutoItens.iTecla)
                        End If
            
                        Produto(iIndice).Caption = objTecladoProdutoItens.sTitulo & "(" & sTecla & ")"
                        
                    Else
                        Produto(iIndice).Caption = objTecladoProdutoItens.sTitulo
                    End If
                Else
                    Produto(iIndice).Caption = ""
                End If
                
                Produto(iIndice).Tag = objTecladoProdutoItens.sArvoreKey
                Produto(iIndice).BackColor = objTecladoProdutoItens.lColor
            End If
        Else
            If Len(objTecladoProdutoItens.sArvoreKey) = Len(Node.Key) Then
                If Left(objTecladoProdutoItens.sArvoreKey, Len(objTecladoProdutoItens.sArvoreKey) - 2) = Left(Node.Key, Len(Node.Key) - 2) Then
                
                    iIndice = CInt(Right(objTecladoProdutoItens.sArvoreKey, 2)) - 1
                
                    If Len(objTecladoProdutoItens.sTitulo) > 0 Then
                
                        If objTecladoProdutoItens.iTecla <> 0 Then
                    
                            If objTecladoProdutoItens.iTecla >= vbKeyF2 Then
                                Call Acha_Tecla(objTecladoProdutoItens.iTecla, sTecla)
                            Else
                                sTecla = Chr(objTecladoProdutoItens.iTecla)
                            End If
                
                            Produto(iIndice).Caption = objTecladoProdutoItens.sTitulo & "(" & sTecla & ")"
                            
                        Else
                            Produto(iIndice).Caption = objTecladoProdutoItens.sTitulo
                        End If
                            
                    Else
                        Produto(iIndice).Caption = ""
                    End If
                    
                    
                    Produto(iIndice).Tag = objTecladoProdutoItens.sArvoreKey
                    Produto(iIndice).BackColor = objTecladoProdutoItens.lColor
                End If
            End If
        End If
    Next
                
    Call Produto_Click(CInt(Right(Node.Key, 2)) - 1)
    
End Sub

Function Valida_Elemento_Duplicado(ByVal sKey As String, ByVal ColTecladoProdutoItens As Collection) As Long

Dim objTecladoProdutoItens As ClassTecladoProdutoItem

On Error GoTo Erro_Valida_Elemento_Duplicado

    For Each objTecladoProdutoItens In ColTecladoProdutoItens
        
        'se usar a mesma tecla aceleradora de outro botao ==> erro
        If StrParaInt(Tecla.Text) <> 0 And objTecladoProdutoItens.sArvoreKey <> sKey Then
            If objTecladoProdutoItens.iTecla = StrParaInt(Tecla.Text) Then gError 214908
        End If
        
    Next

    Valida_Elemento_Duplicado = SUCESSO
    
    Exit Function
        
Erro_Valida_Elemento_Duplicado:
    
    Valida_Elemento_Duplicado = gErr
    
    Select Case gErr
        
        Case 214908
            Call Rotina_Erro(vbOKOnly, "ERRO_TECLA_JA_UTILIZADA", gErr, objTecladoProdutoItens.iTecla, objTecladoProdutoItens.sTitulo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 214909)

    End Select

End Function

Function Carrega_Arvore(ByVal ColTecladoProdutoItens As Collection) As Long

Dim objTecladoProdutoItens As ClassTecladoProdutoItem
Dim objNode As Node
Dim iRelative As Integer

On Error GoTo Erro_Carrega_Arvore

    For Each objTecladoProdutoItens In ColTecladoProdutoItens
        
        Call Procura_Irmao(objTecladoProdutoItens.sArvoreKey, objNode, iRelative)
        
        If objNode Is Nothing Then
            
            TvwMenu.Nodes.Add , , objTecladoProdutoItens.sArvoreKey, objTecladoProdutoItens.sTitulo
            
        Else
                    
            Set objNode = TvwMenu.Nodes.Add(objNode, iRelative, objTecladoProdutoItens.sArvoreKey, objTecladoProdutoItens.sTitulo)
        
        End If
        
    Next
        

    Carrega_Arvore = SUCESSO
    
    Exit Function
        
Erro_Carrega_Arvore:
    
    Carrega_Arvore = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 214919)

    End Select

End Function

