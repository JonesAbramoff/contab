VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpClientesSemRelac 
   ClientHeight    =   6255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7620
   ScaleHeight     =   6255
   ScaleWidth      =   7620
   Begin VB.CheckBox AnalisarFiliaisClientes 
      Caption         =   "Analisar por Filiais"
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
      Left            =   240
      TabIndex        =   29
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Frame FrameCategoriaCliente 
      Caption         =   "Categoria"
      Height          =   1470
      Left            =   240
      TabIndex        =   20
      Top             =   4320
      Width           =   4980
      Begin VB.ComboBox CategoriaCliente 
         Height          =   315
         Left            =   1395
         TabIndex        =   24
         Top             =   540
         Width           =   2745
      End
      Begin VB.ComboBox CategoriaClienteDe 
         Height          =   315
         Left            =   705
         TabIndex        =   23
         Top             =   1020
         Width           =   1740
      End
      Begin VB.CheckBox CategoriaClienteTodas 
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
         Height          =   252
         Left            =   300
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox CategoriaClienteAte 
         Height          =   315
         Left            =   3030
         TabIndex        =   21
         Top             =   1005
         Width           =   1740
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   15
         Left            =   360
         TabIndex        =   28
         Top             =   720
         Width           =   30
      End
      Begin VB.Label LabelCategoriaClienteAte 
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
         Left            =   2595
         TabIndex        =   27
         Top             =   1065
         Width           =   360
      End
      Begin VB.Label LabelCategoriaClienteDe 
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
         Left            =   330
         TabIndex        =   26
         Top             =   1065
         Width           =   315
      End
      Begin VB.Label LabelCategoriaCliente 
         Caption         =   "Categoria:"
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
         Left            =   480
         TabIndex        =   25
         Top             =   585
         Width           =   855
      End
   End
   Begin VB.Frame FrameTipoRelacionamento 
      Caption         =   "Tipo de Relacionamento"
      Height          =   1095
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   4965
      Begin VB.ComboBox TipoRelacionamento 
         Height          =   315
         Left            =   1245
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   600
         Width           =   2550
      End
      Begin VB.OptionButton TipoRelacApenas 
         Caption         =   "Apenas"
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
         Left            =   180
         TabIndex        =   10
         Top             =   615
         Width           =   1050
      End
      Begin VB.OptionButton TipoRelacTodos 
         Caption         =   "Todos os tipos"
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
         Left            =   195
         TabIndex        =   9
         Top             =   285
         Width           =   1620
      End
   End
   Begin VB.Frame FrameTipoCliente 
      Caption         =   "Tipo de Cliente"
      Height          =   1095
      Left            =   240
      TabIndex        =   16
      Top             =   3120
      Width           =   4965
      Begin VB.OptionButton TipoClienteTodos 
         Caption         =   "Todos os tipos"
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
         Left            =   195
         TabIndex        =   11
         Top             =   285
         Width           =   1620
      End
      Begin VB.OptionButton TipoClienteApenas 
         Caption         =   "Apenas"
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
         Left            =   180
         TabIndex        =   18
         Top             =   615
         Width           =   1050
      End
      Begin VB.ComboBox TipoCliente 
         Height          =   315
         Left            =   1245
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   585
         Width           =   2550
      End
   End
   Begin VB.Frame FrameDias 
      Caption         =   "Clientes sem relacionamentos há"
      Height          =   735
      Left            =   240
      TabIndex        =   12
      Top             =   1080
      Width           =   4935
      Begin MSMask.MaskEdBox NumDias 
         Height          =   315
         Left            =   3120
         TabIndex        =   13
         Top             =   300
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   "_"
      End
      Begin VB.Label LabelNumDias 
         AutoSize        =   -1  'True
         Caption         =   "Clientes sem relacionamentos há:"
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
         TabIndex        =   15
         Top             =   360
         Width           =   2850
      End
      Begin VB.Label LabelDias 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3720
         TabIndex        =   14
         Top             =   360
         Width           =   360
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5250
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   165
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpClientesSemRelac.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpClientesSemRelac.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpClientesSemRelac.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpClientesSemRelac.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoExecutar 
      Caption         =   "Executar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5445
      Picture         =   "RelOpClientesSemRelac.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   795
      Width           =   1815
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpClientesSemRelac.ctx":0A96
      Left            =   1050
      List            =   "RelOpClientesSemRelac.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   285
      Width           =   2730
   End
   Begin VB.Label Label1 
      Caption         =   "Opção:"
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
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   330
      Width           =   615
   End
End
Attribute VB_Name = "RelOpClientesSemRelac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'Se a opção AnalisarFiliaisClientes estiver MARCADA e houver uma categoria selecionada, chamar o relatório FCLSRLCT
'Se a opção AnalisarFiliaisClientes estiver MARCADA e NÃO houver uma categoria selecionada, chamar o relatório FCLSRL
'Se a opção AnalisarFiliaisClientes estiver DESMARCADA, chamar o relatório CLSRL'Verificar a tela RelacClientesOcx em RelatoriosFAT2 para fazer o tratamento do frame TipoRelacionamento
'Se a opção AnalisarFiliaisClientes estiver DESMARCADA, os controles do frame CategoriaCliente devem estar desabilitados. Ao marcar essa opção, os mesmos devem ser habilitados
'Verificar a tela RelOpTitRecOcx em RelatoriosCPR2 para fazer o tratamento do frame TipoCliente
'Verificar a tela RelOpEstoqueVendasOcx em RelatoriosFAT2 para fazer o tratamento do frame Categoria
'O campo NumDias, deve ser passado na expressão de seleção como DataBase<=
'Para encontrar essa data base, deve ser utilizada a data atual - o número de dias informado pelo usuário
