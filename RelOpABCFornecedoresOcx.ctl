VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpABCFornecedoresOcx 
   ClientHeight    =   6465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8160
   KeyPreview      =   -1  'True
   ScaleHeight     =   6465
   ScaleWidth      =   8160
   Begin VB.Frame FrameTab 
      BorderStyle     =   0  'None
      Height          =   5055
      Index           =   1
      Left            =   240
      TabIndex        =   24
      Top             =   1095
      Width           =   5415
      Begin VB.Frame FrameFornecedoresTop 
         Caption         =   "Fornecedores Top"
         Height          =   735
         Left            =   120
         TabIndex        =   55
         Top             =   4320
         Width           =   2715
         Begin MSMask.MaskEdBox FornTop 
            Height          =   315
            Left            =   1680
            TabIndex        =   12
            Top             =   300
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin VB.Label LabelAnalisarFornecedoresTop 
            AutoSize        =   -1  'True
            Caption         =   "Fornecedores top:"
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
            Left            =   120
            TabIndex        =   56
            Top             =   345
            Width           =   1560
         End
      End
      Begin VB.Frame FrameFornecedores 
         Caption         =   "Fornecedores"
         Height          =   1395
         Left            =   120
         TabIndex        =   52
         Top             =   2825
         Width           =   2355
         Begin MSMask.MaskEdBox FornecedorDe 
            Height          =   300
            Left            =   675
            TabIndex        =   7
            Top             =   345
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   8
            Mask            =   "########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FornecedorAte 
            Height          =   300
            Left            =   690
            TabIndex        =   8
            Top             =   915
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   8
            Mask            =   "########"
            PromptChar      =   " "
         End
         Begin VB.Label LabelFornecedorAte 
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
            Height          =   192
            Left            =   240
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   54
            Top             =   954
            Width           =   360
         End
         Begin VB.Label LabelFornecedorDe 
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
            Height          =   192
            Left            =   276
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   53
            Top             =   384
            Width           =   312
         End
      End
      Begin VB.Frame FrameData 
         Caption         =   "Data "
         Height          =   1335
         Left            =   120
         TabIndex        =   47
         Top             =   0
         Width           =   2355
         Begin MSComCtl2.UpDown UpDownDataDe 
            Height          =   315
            Left            =   1785
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   240
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataDe 
            Height          =   315
            Left            =   615
            TabIndex        =   1
            Top             =   255
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataAte 
            Height          =   315
            Left            =   1830
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   840
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataAte 
            Height          =   315
            Left            =   645
            TabIndex        =   2
            Top             =   855
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   270
            TabIndex        =   51
            Top             =   915
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   285
            TabIndex        =   50
            Top             =   315
            Width           =   315
         End
      End
      Begin VB.Frame FrameProdutos 
         Caption         =   "Produtos"
         Height          =   1290
         Left            =   120
         TabIndex        =   42
         Top             =   1435
         Width           =   5160
         Begin MSMask.MaskEdBox ProdutoDe 
            Height          =   315
            Left            =   495
            TabIndex        =   5
            Top             =   360
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdutoAte 
            Height          =   315
            Left            =   510
            TabIndex        =   6
            Top             =   825
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
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
            Left            =   120
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   46
            Top             =   870
            Width           =   435
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
            Left            =   150
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   45
            Top             =   390
            Width           =   360
         End
         Begin VB.Label ProdutoDescDe 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2040
            TabIndex        =   44
            Top             =   360
            Width           =   3000
         End
         Begin VB.Label ProdutoDescAte 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2040
            TabIndex        =   43
            Top             =   825
            Width           =   3000
         End
      End
      Begin VB.CheckBox DetalharFilial 
         Caption         =   "Detalhar por Filial"
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
         Left            =   3120
         TabIndex        =   13
         Top             =   4590
         Width           =   1935
      End
      Begin VB.Frame FrameFilialEmpresa 
         Caption         =   "Filial Empresa"
         Height          =   1335
         Left            =   2520
         TabIndex        =   26
         Top             =   0
         Width           =   2760
         Begin VB.ComboBox FilialEmpresaAte 
            Height          =   315
            Left            =   585
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   870
            Width           =   1860
         End
         Begin VB.ComboBox FilialEmpresaDe 
            Height          =   315
            Left            =   585
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   285
            Width           =   1860
         End
         Begin VB.Label LabelFilialEmpresaDe 
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
            Left            =   265
            TabIndex        =   28
            Top             =   330
            Width           =   315
         End
         Begin VB.Label LabelFilialEmpresaAte 
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
            Left            =   220
            TabIndex        =   27
            Top             =   930
            Width           =   360
         End
      End
      Begin VB.Frame FrameTipoProdutos 
         Caption         =   "Tipo de Produtos"
         Height          =   1395
         Left            =   2520
         TabIndex        =   25
         Top             =   2820
         Width           =   2760
         Begin VB.ComboBox TipoProdutos 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   900
            Width           =   1365
         End
         Begin VB.OptionButton TipoProdutosTodos 
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
            Height          =   225
            Left            =   240
            TabIndex        =   9
            Top             =   345
            Width           =   1050
         End
         Begin VB.OptionButton TipoProdutosApenas 
            Caption         =   "Apenas "
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
            Left            =   225
            TabIndex        =   10
            Top             =   900
            Width           =   1035
         End
      End
   End
   Begin VB.Frame FrameTab 
      BorderStyle     =   0  'None
      Height          =   4455
      Index           =   2
      Left            =   240
      TabIndex        =   29
      Top             =   1080
      Width           =   5415
      Begin VB.Frame FrameCategoriaFornecedores 
         Caption         =   "Categoria de Fornecedores"
         Height          =   2175
         Left            =   120
         TabIndex        =   33
         Top             =   2280
         Width           =   5160
         Begin VB.CommandButton BotaoItensCatFornDesmarca 
            Height          =   360
            Left            =   4440
            Picture         =   "RelOpABCFornecedoresOcx.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Desmarca todos os itens da categoria selecionada."
            Top             =   1560
            Width           =   420
         End
         Begin VB.CommandButton BotaoItensCatFornMarca 
            Height          =   360
            Left            =   4440
            Picture         =   "RelOpABCFornecedoresOcx.ctx":067E
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Marca todos os itens da categoria selecionada."
            Top             =   1200
            Width           =   420
         End
         Begin VB.ListBox ItensCategoriaFornecedores 
            Height          =   960
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   34
            Top             =   1080
            Width           =   4215
         End
         Begin VB.ComboBox CategoriaFornecedores 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   360
            Width           =   2100
         End
         Begin VB.Label LabelItensCategoriaFornecedores 
            AutoSize        =   -1  'True
            Caption         =   "Itens:"
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
            Left            =   120
            TabIndex        =   36
            Top             =   840
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label LabelCategoriaFornecedores 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   630
            TabIndex        =   35
            Top             =   390
            Width           =   870
         End
      End
      Begin VB.Frame FrameCategoriaProdutos 
         Caption         =   "Categoria de Produtos"
         Height          =   2175
         Left            =   120
         TabIndex        =   30
         Top             =   0
         Width           =   5160
         Begin VB.CommandButton BotaoItensCatProdDesmarca 
            Height          =   360
            Left            =   4440
            Picture         =   "RelOpABCFornecedoresOcx.ctx":0CB0
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Desmarca todos os itens da categoria selecionada."
            Top             =   1560
            Width           =   420
         End
         Begin VB.CommandButton BotaoItensCatProdMarca 
            Height          =   360
            Left            =   4440
            Picture         =   "RelOpABCFornecedoresOcx.ctx":132E
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Marca todos os itens da categoria selecionada."
            Top             =   1200
            Width           =   420
         End
         Begin VB.ListBox ItensCategoriaProdutos 
            Height          =   960
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   37
            Top             =   1080
            Width           =   4215
         End
         Begin VB.ComboBox CategoriaProdutos 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   360
            Width           =   2100
         End
         Begin VB.Label LabelCategoriaProdutos 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   630
            TabIndex        =   32
            Top             =   390
            Width           =   870
         End
         Begin VB.Label LabelItensCategoriaProdutos 
            AutoSize        =   -1  'True
            Caption         =   "Itens:"
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
            Left            =   120
            TabIndex        =   31
            Top             =   840
            Visible         =   0   'False
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5655
      Left            =   120
      TabIndex        =   23
      Top             =   720
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   9975
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Principal"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Categorias"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5880
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   135
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpABCFornecedoresOcx.ctx":1960
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpABCFornecedoresOcx.ctx":1ABA
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpABCFornecedoresOcx.ctx":1C44
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpABCFornecedoresOcx.ctx":2176
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpABCFornecedoresOcx.ctx":22F4
      Left            =   1800
      List            =   "RelOpABCFornecedoresOcx.ctx":22F6
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2490
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
      Left            =   6045
      Picture         =   "RelOpABCFornecedoresOcx.ctx":22F8
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label LabelOpcao 
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
      Left            =   1080
      TabIndex        =   22
      Top             =   285
      Width           =   615
   End
End
Attribute VB_Name = "RelOpABCFornecedoresOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Essa tela já existe em RelatoriosCOM... vc deverá renomear a tela atual para RelOpABCFornecedoresOldOcx
'e adicionar essa tela ao .vbp
'Se quiser, pode aproveitar boa parte do código já existente

'Caption: Relatório ABC de Fornecedores

'No Botao_Executar:
'   - chamar Move_Tela_Memoria: guardar os dados da tela em objRelABCFornecedoresTela (ClassRelABCFornecedoresTela)
'   - chamar a função SldDiaForn_Le_RelABCFornecedores( objRelABCFornecedoresTela, colItensRelABCFornecedores)
'   - chamar a função RelABCFornecedores_Grava (colItensRelABCFornecedores)
'   - Guardar em gobjRelOpcoes.sSelecao = "@NNUMINTREL =" & colItensRelABCFornecedores(1).lNumIntRel
'   - chamar a função gobjRelatorio.Executar_Prossegue2(Me)

'Ao selecionar uma categoria de fornecedor, a opção 'Detalhar por Filial' deverá ser marcada e desabilitada
'Ao limpar a categoria de fornecedor, a opção 'Detalhar por Filial' deverá ficar habilitada novamente

'Não é necessário fazer a função Monta_Expressao_Selecao


Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1
Private WithEvents objEventoFornDe As AdmEvento
Attribute objEventoFornDe.VB_VarHelpID = -1
Private WithEvents objEventoFornAte As AdmEvento
Attribute objEventoFornAte.VB_VarHelpID = -1

Dim iFrameAtual As Integer

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

'***** INICIALIZAÇÃO DA TELA - INÍCIO *****
Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    iFrameAtual = 1
    
    'Inicializa as variáveis do Browser
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    Set objEventoFornDe = New AdmEvento
    Set objEventoFornAte = New AdmEvento
        
    'Função que Carrega a Combo de FilialEmpresa
    lErro = Carrega_FilialEmpresa()
    If lErro <> SUCESSO Then gError 125949
    
    'Função que carrega a Combo de Categorias
    lErro = Carrega_CategoriasProduto()
    If lErro <> SUCESSO Then gError 125950
    
    'Carrega Combo Categoria Fornecedores
    lErro = Carrega_CategoriaFornecedores()
    If lErro <> SUCESSO Then gError 125951
    
    'Inicializa a máscara de ProdutoDe
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoDe)
    If lErro <> SUCESSO Then gError 125952

    'Inicializa a máscara de ProdutoAte
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoAte)
    If lErro <> SUCESSO Then gError 125953
    
    'Carrega a Combo TipoProdutos
    lErro = Carrega_TipoProdutos()
    If lErro <> SUCESSO Then gError 125954
    
    'Seta o TipoProdutos como Todos
    Call TipoProdutosTodos_Click
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 125949 To 125954
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166801)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    'Limpa Objetos da memoria
    Set objEventoFornDe = Nothing
    Set objEventoFornAte = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 125955

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 125956

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 125955
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case 125956
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166802)

    End Select

    Exit Function

End Function
'***** INICIALIZAÇÃO DA TELA - FIM *****

'***** EVENTO GOTFOCUS DOS CONTROLES - INÍCIO *****
Private Sub FornecedorDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(FornecedorDe)

End Sub

Private Sub FornecedorAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(FornecedorAte)

End Sub

Private Sub DataAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataAte)

End Sub

Private Sub DataDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataDe)

End Sub

Private Sub ProdutoAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ProdutoAte)

End Sub

Private Sub ProdutoDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ProdutoDe)

End Sub

Private Sub FornTop_GotFocus()

    Call MaskEdBox_TrataGotFocus(FornTop)

End Sub
'***** EVENTO GOTFOCUS DOS CONTROLES - FIM *****

'***** EVENTO VALIDATE DOS CONTROLES - INÍCIO *****
Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub FornecedorDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_FornecedorDe_Validate

    If Len(Trim(FornecedorDe.Text)) > 0 Then

        'Lê o código informado
        objFornecedor.lCodigo = LCodigo_Extrai(FornecedorDe.Text)
        
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12729 Then gError 125957
        
        'Se não encontrou o Fornecedor ==> erro
        If lErro = 12729 Then gError 125958
        
    End If

    Exit Sub

Erro_FornecedorDe_Validate:

    Cancel = True

    Select Case gErr

        Case 125957

        Case 125958
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166803)

    End Select

    Exit Sub

End Sub

Private Sub FornecedorAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_FornecedorAte_Validate

    If Len(Trim(FornecedorAte.Text)) > 0 Then

        'Lê o código informado
        objFornecedor.lCodigo = LCodigo_Extrai(FornecedorAte.Text)
        
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12729 Then gError 125959
        
        'Se não encontrou o Fornecedor ==> erro
        If lErro = 12729 Then gError 125960
        
    End If

    Exit Sub

Erro_FornecedorAte_Validate:

    Cancel = True

    Select Case gErr

        Case 125959

        Case 125960
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166804)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataDe.ClipText)) = 0 Then Exit Sub

    'Critica a DataPedDe informada
    lErro = Data_Critica(DataDe.Text)
    If lErro <> SUCESSO Then gError 125961

    Exit Sub
                   
Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 125961
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166805)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'Verifica se a DataAte está preenchida
    If Len(Trim(DataDe.ClipText)) = 0 Then Exit Sub

    'Critica a DataAte informada
    lErro = Data_Critica(DataAte.Text)
    If lErro <> SUCESSO Then gError 125962

    Exit Sub
                   
Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 125962
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166806)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_ProdutoDe_Validate

    If Len(Trim(ProdutoDe.ClipText)) > 0 Then
        
        'Verifica se o Produto é válido
        lErro = CF("Produto_Perde_Foco", ProdutoDe, ProdutoDescDe)
        If lErro <> SUCESSO And lErro <> 27095 Then gError 125963
    
        If lErro = 27095 Then gError 125964
    
    End If
    
    Exit Sub
    
Erro_ProdutoDe_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 125963
        
        Case 125964
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166807)
            
    End Select
    
End Sub

Private Sub ProdutoAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_ProdutoAte_Validate

    If Len(Trim(ProdutoAte.ClipText)) > 0 Then
        
        'Verifica se o produto é válido
        lErro = CF("Produto_Perde_Foco", ProdutoAte, ProdutoDescAte)
        If lErro <> SUCESSO And lErro <> 27095 Then gError 125965
    
        If lErro = 27095 Then gError 125966
    
    End If
    
    Exit Sub
    
Erro_ProdutoAte_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 125965
        
        Case 125966
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166808)
            
    End Select
    
End Sub
'***** EVENTO VALIDATE DOS CONTROLES - FIM *****

'***** EVENTO CLICK DOS CONTROLES - INÍCIO *****
Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Diminui um dia a Data
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 125967

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 125967
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 166809)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Aumenta um Dia a Data
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 125968

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 125968
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 166810)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Diminui um dia Data
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 125969

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 125969
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 166811)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Aumenta Um dia a Data
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 125970

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 125970
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 166812)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaProdutos_Click()
'Preenche os itens da categoria selecionada

Dim lErro As Long
Dim colItensCategoria As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As ClassCategoriaProdutoItem

On Error GoTo Erro_CategoriaProdutos_Click

    'Limpa a Combo de Itens
    ItensCategoriaProdutos.Clear
    
    If Len(Trim(CategoriaProdutos.Text)) > 0 Then

        'Preenche o Obj
        objCategoriaProduto.sCategoria = CategoriaProdutos.List(CategoriaProdutos.ListIndex)
        
        'Le as categorias do Produto
        lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colItensCategoria)
        If lErro <> SUCESSO And lErro <> 22541 Then gError 125971
                
        For Each objCategoriaProdutoItem In colItensCategoria
            ItensCategoriaProdutos.AddItem (objCategoriaProdutoItem.sItem)
        Next
        
    End If
    
    Exit Sub

Erro_CategoriaProdutos_Click:

    Select Case gErr

         Case 125971
         
         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166813)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaFornecedores_Click()
'Preenche os itens da categoria selecionada

Dim lErro As Long
Dim colItensCategoria As New Collection
Dim objCategoriaFornItem As New ClassCategoriaFornItem

On Error GoTo Erro_CategoriaFornecedores_Click

    'Limpa a Combo de Itens
    ItensCategoriaFornecedores.Clear
    
    If Len(Trim(CategoriaFornecedores.Text)) > 0 Then

        'Preenche o Obj
        objCategoriaFornItem.sCategoria = CategoriaFornecedores.List(CategoriaFornecedores.ListIndex)
        
        'Le as categorias do Fornecedor
        lErro = CF("CategoriaFornecedor_Le_Itens", objCategoriaFornItem, colItensCategoria)
        If lErro <> SUCESSO And lErro <> 91180 Then gError 125972
                
        For Each objCategoriaFornItem In colItensCategoria
            ItensCategoriaFornecedores.AddItem (objCategoriaFornItem.sItem)
        Next
        
        DetalharFilial.Value = 1
        DetalharFilial.Enabled = False
        
    Else
    
        DetalharFilial.Value = 0
        DetalharFilial.Enabled = True
        
    End If
    
    
    Exit Sub

Erro_CategoriaFornecedores_Click:

    Select Case gErr

         Case 125972
         
         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166814)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If Len(Trim(ComboOpcoes.Text)) = 0 Then gError 125973

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 125974

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 125975
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 125976
    
    Call Limpa_Tela_Rel
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 125973
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 125974 To 125976
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166815)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 125977

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 125978

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call Limpa_Tela_Rel

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 125977
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 125978

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166816)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela_Rel
    
End Sub

Private Sub BotaoFechar_Click()
    
    Unload Me
    
End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long
Dim objRelABCFornecedoresTela As New ClassRelABCFornecedoresTela
Dim colItensRelABCFornecedores As New Collection

On Error GoTo Erro_BotaoExecutar_Click

    'Seta o ponteiro do mouse como ampulheta
    MousePointer = vbHourglass
    
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 125979

    'Move para a memória as informações da tela
    lErro = Move_Tela_Memoria(objRelABCFornecedoresTela)
    If lErro <> SUCESSO Then gError 125980
    
    'Gera os dados do relatório
    lErro = CF("RelABCFornecedores_Gera", objRelABCFornecedoresTela, colItensRelABCFornecedores)
    If lErro <> SUCESSO Then gError 125981
    
    'Passa o critério de seleção dos registros que farão parte do relatório
    gobjRelOpcoes.sSelecao = "NumIntRel=" & colItensRelABCFornecedores(1).lNumIntRel
    
    'Se foi selecionada uma categoria de fornecedor
    If Len(Trim(objRelABCFornecedoresTela.sCategoriaFornecedores)) > 0 Then
    
        'chama o relatório preparado para imprimir o item de categoria de cada produto
        gobjRelatorio.sNomeTsk = "abcforct"
        
        'determina que o relatório deve ser impresso com layouttipo landscape
        gobjRelatorio.iLandscape = 1
    
    'Se foi selecionado para detalhar o relatório por filial de fornecedor
    ElseIf objRelABCFornecedoresTela.iDetalharFilial = MARCADO Then
    
        'chama o relatório preparado para imprimir detalhando por filial de fornecedor
        gobjRelatorio.sNomeTsk = "abcforfi"
        
        'determina que o relatório deve ser impresso com layouttipo landscape
        gobjRelatorio.iLandscape = 1
    
    'senão
    Else
        
        'chama o relatório preparado para imprimir sem o item de categoria de cada fornecedor e sem detalhar por filial
        gobjRelatorio.sNomeTsk = "abcfor"
        
    End If
    
    Call gobjRelatorio.Executar_Prossegue2(Me)

    'Seta o ponteiro padrão do mouse
    MousePointer = vbDefault

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 125979 To 125982

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166817)

    End Select

    'Seta o ponteiro padrão do mouse
    MousePointer = vbDefault

    Exit Sub

End Sub

Private Sub TipoProdutosTodos_Click()
    
    TipoProdutosTodos.Value = True
    TipoProdutosTodos.Enabled = True
    If Len(Trim(TipoProdutos.Text)) <> 0 Then TipoProdutos.ListIndex = -1
    TipoProdutos.Enabled = False
        
End Sub

Private Sub TipoProdutosApenas_Click()
    
    TipoProdutosApenas.Value = True
    TipoProdutos.Enabled = True
    
End Sub

Private Sub LabelProdutoDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String

On Error GoTo Erro_LabelProdutoDe_Click
    
    If Len(Trim(ProdutoDe.ClipText)) > 0 Then
        
        'Preenche com o Produto da tela
        lErro = CF("Produto_Formata", ProdutoDe.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 125983
        
        objProduto.sCodigo = sProdutoFormatado
    
    End If
    
    'Chama Tela ProdutoCompraLista
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 125983
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166818)

    End Select

    Exit Sub

End Sub

Private Sub LabelProdutoAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String

On Error GoTo Erro_LabelProdutoAte_Click
    
    If Len(Trim(ProdutoAte.ClipText)) > 0 Then
        
        'Preenche com o Produto da tela
        lErro = CF("Produto_Formata", ProdutoAte.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 125984
        
        objProduto.sCodigo = sProdutoFormatado
    
    End If
    
    'Chama Tela ProdutoCompraLista
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoProdutoAte)

   Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 125984
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166819)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index = iFrameAtual Then Exit Sub

    If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub
        
    'Torna Frame correspondente ao Tab selecionado visivel
    FrameTab(TabStrip1.SelectedItem.Index).Visible = True
    'Torna Frame atual invisivel
    FrameTab(iFrameAtual).Visible = False
    'Armazena novo valor de iFrameAtual
    iFrameAtual = TabStrip1.SelectedItem.Index
    
    Exit Sub
    
Erro_TabStrip1_Click:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166820)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub LabelFornecedorDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelFornecedorDe_Click

    If Len(Trim(FornecedorDe.Text)) <> 0 Then
    
        objFornecedor.lCodigo = StrParaLong(FornecedorDe.Text)
        
    End If
    
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornDe)
    
    Exit Sub

Erro_LabelFornecedorDe_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166821)

    End Select

    Exit Sub
    
End Sub

Private Sub LabelFornecedorAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelFornecedorAte_Click

    If Len(Trim(FornecedorAte.Text)) > 0 Then
    
        objFornecedor.lCodigo = StrParaLong(FornecedorAte.Text)
        
    End If
    
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornAte)
    
    Exit Sub

Erro_LabelFornecedorAte_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166822)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoItensCatProdMarca_Click()

Dim iIndice As Integer

On Error GoTo Erro_BotaoItensCatProdMarca_Click

    'Marca todos os itens
    For iIndice = 0 To ItensCategoriaProdutos.ListCount - 1

        ItensCategoriaProdutos.Selected(iIndice) = True
        
    Next
    
    Exit Sub
    
Erro_BotaoItensCatProdMarca_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166823)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoItensCatProdDesmarca_Click()

Dim iIndice As Integer

On Error GoTo Erro_BotaoItensCatProdDesmarca_Click

    'Desmarca todos os itens
    For iIndice = 0 To ItensCategoriaProdutos.ListCount - 1

        ItensCategoriaProdutos.Selected(iIndice) = False
        
    Next
    
    Exit Sub
    
Erro_BotaoItensCatProdDesmarca_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166824)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoItensCatFornMarca_Click()

Dim iIndice As Integer

On Error GoTo Erro_BotaoItensCatFornMarca_Click

    'marca todos os itens de categoria Fornecedor
    For iIndice = 0 To ItensCategoriaFornecedores.ListCount - 1

        ItensCategoriaFornecedores.Selected(iIndice) = True
        
    Next
    
    Exit Sub
    
Erro_BotaoItensCatFornMarca_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166825)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoItensCatFornDesmarca_Click()

Dim iIndice As Integer

On Error GoTo Erro_BotaoItensCatFornDesmarca_Click

    'Desmarca todos od itens de categoria Fornecedor
    For iIndice = 0 To ItensCategoriaFornecedores.ListCount - 1

        ItensCategoriaFornecedores.Selected(iIndice) = False
        
    Next
    
    Exit Sub
    
Erro_BotaoItensCatFornDesmarca_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166826)

    End Select

    Exit Sub
    
End Sub
'***** EVENTO CLICK DOS CONTROLES - FIM *****

'***** EVENTO BEFORECLICK DOS CONTROLES - FIM *****
Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub
'***** EVENTO BEFORECLICK DOS CONTROLES - FIM *****

'***** FUNÇÕES DE APOIO À TELA - INÍCIO *****
Private Function Carrega_FilialEmpresa() As Long
'Carrega a Combo FilialEmpresa com as informações do BD

Dim lErro As Long
Dim iIndice As Integer
Dim objFilialEmpresa As New AdmFiliais
Dim colFiliais As New Collection

On Error GoTo Erro_Carrega_FilialEmpresa

    'Faz a Leitura das Filiais
    lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, colFiliais)
    If lErro <> SUCESSO Then gError 125985
    
    FilialEmpresaDe.AddItem ("")
    FilialEmpresaAte.AddItem ("")
    
    'Carrega as combos
    For Each objFilialEmpresa In colFiliais
        
        'Se nao for a EMPRESA_TODA
        If objFilialEmpresa.iCodFilial <> EMPRESA_TODA Then
            
            FilialEmpresaDe.AddItem (objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome)
            FilialEmpresaAte.AddItem (objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome)
            
        End If
        
    Next

    Carrega_FilialEmpresa = SUCESSO
    
    Exit Function
    
Erro_Carrega_FilialEmpresa:

    Carrega_FilialEmpresa = gErr
    
    Select Case gErr
    
        Case 125985

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166827)
    
    End Select

    Exit Function

End Function

Private Function Carrega_TipoProdutos() As Long
'Carrega a coombo TipoProduto com as informações do BD

Dim lErro As Long
Dim colCod_DescReduzida As New AdmColCodigoNome
Dim objCod_DescReduzida As New AdmCodigoNome

On Error GoTo Erro_Carrega_TipoProdutos

    lErro = CF("TiposProduto_Le_Todos", colCod_DescReduzida)
    If lErro <> SUCESSO Then gError 125986

    'Carrega as combo TipoProdutos
    For Each objCod_DescReduzida In colCod_DescReduzida
    
        TipoProdutos.AddItem objCod_DescReduzida.iCodigo & SEPARADOR & objCod_DescReduzida.sNome
        TipoProdutos.ItemData(TipoProdutos.NewIndex) = objCod_DescReduzida.iCodigo
        
    Next
    
    Carrega_TipoProdutos = SUCESSO

    Exit Function

Erro_Carrega_TipoProdutos:

    Carrega_TipoProdutos = gErr
    
    Select Case gErr
    
        Case 125986
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166828)
    
    End Select

    Exit Function

End Function

Private Function Carrega_CategoriasProduto() As Long
'Carrega a Combo CategoriaProdutos com informações do BD

Dim lErro As Long
Dim objCategoria As New ClassCategoriaProduto
Dim colCategorias As New Collection

On Error GoTo Erro_Carrega_CategoriasProduto
    
    'Le a categoria
    lErro = CF("CategoriasProduto_Le_Todas", colCategorias)
    If lErro <> SUCESSO And lErro <> 22542 Then gError 125987
    
    'Se nao encontrou => Erro
    If lErro = 22542 Then gError 125988
    
    CategoriaProdutos.AddItem ("")
    
    'Carrega as combos de Categorias
    For Each objCategoria In colCategorias
    
        CategoriaProdutos.AddItem objCategoria.sCategoria
        
    Next
    
    Carrega_CategoriasProduto = SUCESSO
    
    Exit Function
    
Erro_Carrega_CategoriasProduto:

    Carrega_CategoriasProduto = gErr
    
    Select Case gErr
    
        Case 125987
        
        Case 125988
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_NAO_CADASTRADA", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166829)
    
    End Select

    Exit Function

End Function

Private Function Carrega_CategoriaFornecedores() As Long
'Carrega a Combo Categoria Fornecedor com as informações do BD

Dim lErro As Long
Dim objCategoriaFornecedor As New ClassCategoriaFornecedor
Dim colCategorias As New Collection

On Error GoTo Erro_Carrega_CategoriaFornecedores
    
    'Le a categoria
    lErro = CF("CategoriaFornecedor_Le_Todos", colCategorias)
    If lErro <> SUCESSO And lErro <> 68486 Then gError 125989
    
    CategoriaFornecedores.AddItem ("")
    
    'Carrega as combos de Categorias
    For Each objCategoriaFornecedor In colCategorias
    
        CategoriaFornecedores.AddItem objCategoriaFornecedor.sCategoria
        
    Next
    
    Carrega_CategoriaFornecedores = SUCESSO
    
    Exit Function
    
Erro_Carrega_CategoriaFornecedores:

    Carrega_CategoriaFornecedores = gErr
    
    Select Case gErr
    
        Case 125989
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166830)
    
    End Select

    Exit Function

End Function

Private Sub Limpa_Tela_Rel()
'Limpa a tela

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Rel

    'Limpa o Relatório
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 125991

    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    CategoriaProdutos.ListIndex = -1
    Call CategoriaProdutos_Click
    FilialEmpresaDe.ListIndex = -1
    FilialEmpresaAte.ListIndex = -1
    
    Call CategoriaFornecedores_Click
    
    ProdutoDescDe.Caption = ""
    ProdutoDescAte.Caption = ""
    
    Call TipoProdutosTodos_Click
    
    DetalharFilial.Value = 0
    DetalharFilial.Enabled = True
    
    Exit Sub

Erro_Limpa_Tela_Rel:

    Select Case gErr

        Case 125991

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166831)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sProd_I As String
Dim sProd_F As String
Dim iCodFilialDe As Integer
Dim iCodFilialAte As Integer
Dim colItens As New Collection
Dim iCont As Integer
Dim iIndice As Integer
Dim sTipoProdutos As String
Dim sCheckTipoProdutos As String

On Error GoTo Erro_PreenchgerrelOp

    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)

    iCodFilialDe = Codigo_Extrai(FilialEmpresaDe.Text)
    iCodFilialAte = Codigo_Extrai(FilialEmpresaAte.Text)
    
    'Critica os parametros
    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F, iCodFilialDe, iCodFilialAte, sCheckTipoProdutos, sTipoProdutos)
    If lErro <> SUCESSO Then gError 125992

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 125993
    
    'Preenche o Produto Inicial
    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then gError 125994

    'Preenche o Produto Final
    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then gError 125995

    'Preenche o Filial Inicial
    lErro = objRelOpcoes.IncluirParametro("NCODFILIALDE", CStr(iCodFilialDe))
    If lErro <> AD_BOOL_TRUE Then gError 125996

    'Preenche o Filial Final
    lErro = objRelOpcoes.IncluirParametro("TFILIALDE", CStr(FilialEmpresaDe.Text))
    If lErro <> AD_BOOL_TRUE Then gError 127090

    'Preenche o Filial Final
    lErro = objRelOpcoes.IncluirParametro("NCODFILIALATE", CStr(iCodFilialAte))
    If lErro <> AD_BOOL_TRUE Then gError 125997
    
    'Preenche o Filial Final
    lErro = objRelOpcoes.IncluirParametro("TFILIALATE", CStr(FilialEmpresaAte.Text))
    If lErro <> AD_BOOL_TRUE Then gError 127091
    
    'Preenche a dataDe
    If Len(Trim(DataDe.ClipText)) = 0 Then
        lErro = objRelOpcoes.IncluirParametro("DDATADE", CStr(DATA_NULA))
        If lErro <> AD_BOOL_TRUE Then gError 125998
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATADE", DataDe.Text)
        If lErro <> AD_BOOL_TRUE Then gError 125999
    End If
    
    'Preenche a DataAte
    If Len(Trim(DataAte.ClipText)) = 0 Then
        lErro = objRelOpcoes.IncluirParametro("DDATAATE", CStr(DATA_NULA))
        If lErro <> AD_BOOL_TRUE Then gError 128000
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAATE", DataAte.Text)
        If lErro <> AD_BOOL_TRUE Then gError 128001
    End If
    
    'Inicia o Contador
    iCont = 0
    
    'Monta o Filtro
    For iIndice = 0 To ItensCategoriaProdutos.ListCount - 1
        
        'Verifica se o Item da Categoria foi selecionado
        If ItensCategoriaProdutos.Selected(iIndice) = True Then
            
            'Incrementa o Contador
            iCont = iCont + 1
            
            lErro = objRelOpcoes.IncluirParametro("TITEMDE" & iCont, CStr(ItensCategoriaProdutos.List(iIndice)))
            If lErro <> AD_BOOL_TRUE Then gError 128002
                            
            colItens.Add CStr(ItensCategoriaProdutos.List(iIndice))
                             
        End If
            
    Next
        
    lErro = objRelOpcoes.IncluirParametro("TCATEGORIAPROD", CategoriaProdutos.Text)
    If lErro <> AD_BOOL_TRUE Then gError 128003
    
    'Inicia o Contador
    iCont = 0
    
    'Monta o Filtro
    For iIndice = 0 To ItensCategoriaFornecedores.ListCount - 1
        
        'Verifica se o Item da Categoria foi selecionado
        If ItensCategoriaFornecedores.Selected(iIndice) = True Then
            
            'Incrementa o Contador
            iCont = iCont + 1
            
            lErro = objRelOpcoes.IncluirParametro("TITEMFORNDE" & iCont, CStr(ItensCategoriaFornecedores.List(iIndice)))
            If lErro <> AD_BOOL_TRUE Then gError 128004
                            
            colItens.Add CStr(ItensCategoriaFornecedores.List(iIndice))
                             
        End If
            
    Next
        
    'Preenche a Categoria Fornecedor
    lErro = objRelOpcoes.IncluirParametro("TCATEGORIAFORN", CategoriaFornecedores.Text)
    If lErro <> AD_BOOL_TRUE Then gError 128005

    'Preenche o tipo do Produto
    lErro = objRelOpcoes.IncluirParametro("TTIPOPROD", sTipoProdutos)
    If lErro <> AD_BOOL_TRUE Then gError 128006
    
    'Preenche com a Opcao TipoProdutos(Todos Produtos ou um Produto)
    lErro = objRelOpcoes.IncluirParametro("NTIPOPRODUTOS", Codigo_Extrai(TipoProdutos.Text))
    If lErro <> AD_BOOL_TRUE Then gError 128007
    
    'Preenche o FornecedorDe
    lErro = objRelOpcoes.IncluirParametro("NFORNECEDORDE", CStr(FornecedorDe.Text))
    If lErro <> AD_BOOL_TRUE Then gError 128008

    'Preenche o FornecedorAte
    lErro = objRelOpcoes.IncluirParametro("NFORNECEDORATE", CStr(FornecedorAte.Text))
    If lErro <> AD_BOOL_TRUE Then gError 128009

    'Preenche o FornecedorTop
    lErro = objRelOpcoes.IncluirParametro("NFORNECEDORESTOP", CStr(Trim(FornTop.Text)))
    If lErro <> AD_BOOL_TRUE Then gError 128010
    
    'Preenche a CheckBox DetalharFilial
    lErro = objRelOpcoes.IncluirParametro("NDETALHARFILIAL", CStr(DetalharFilial.Value))
    If lErro <> AD_BOOL_TRUE Then gError 128052
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreenchgerrelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 125992 To 125999, 128000 To 128010, 128052

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166832)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim iCont As Integer
Dim iIndice As Integer

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 128011

    'Traz o Parâmetro Referênte ao Produto Inicial
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro <> SUCESSO Then gError 128012
    
    ProdutoDe.PromptInclude = False
    ProdutoDe.Text = sParam
    ProdutoDe.PromptInclude = True
    Call ProdutoDe_Validate(bSGECancelDummy)
    
    'Traz o Parâmetro Referênte ao Produto Final
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 128013
    
    ProdutoAte.PromptInclude = False
    ProdutoAte.Text = sParam
    ProdutoAte.PromptInclude = True
    Call ProdutoAte_Validate(bSGECancelDummy)
    
    'Traz o Codigo da Filial Inicial
    lErro = objRelOpcoes.ObterParametro("NCODFILIALDE", sParam)
    If lErro <> SUCESSO Then gError 128014
    
    For iIndice = 0 To FilialEmpresaDe.ListCount - 1
        If Codigo_Extrai(FilialEmpresaDe.List(iIndice)) = StrParaInt(sParam) Then
            FilialEmpresaDe.ListIndex = iIndice
            Exit For
        End If
    Next

    'Traz o Codigo da Filial Final
    lErro = objRelOpcoes.ObterParametro("NCODFILIALATE", sParam)
    If lErro <> SUCESSO Then gError 128015
    
    For iIndice = 0 To FilialEmpresaAte.ListCount - 1
        If Codigo_Extrai(FilialEmpresaAte.List(iIndice)) = StrParaInt(sParam) Then
            FilialEmpresaAte.ListIndex = iIndice
            Exit For
        End If
    Next
    
    'Traz a Datade Para a Tela
    lErro = objRelOpcoes.ObterParametro("DDATADE", sParam)
    If lErro <> SUCESSO Then gError 128016
    
    If sParam <> DATA_NULA Then
        
        DataDe.PromptInclude = False
        DataDe.Text = sParam
        DataDe.PromptInclude = True
        Call DataDe_Validate(bSGECancelDummy)
    
    Else
        DataDe.PromptInclude = False
        DataDe.Text = ""
        DataDe.PromptInclude = True
        
    End If
    
    'Traz a Datade Para a Tela
    lErro = objRelOpcoes.ObterParametro("DDATAATE", sParam)
    If lErro <> SUCESSO Then gError 128017
    
    If sParam <> DATA_NULA Then
        
        DataAte.PromptInclude = False
        DataAte.Text = sParam
        DataAte.PromptInclude = True
        Call DataAte_Validate(bSGECancelDummy)
    
    Else
        
        DataAte.PromptInclude = False
        DataAte.Text = ""
        DataAte.PromptInclude = True
        Call DataAte_Validate(bSGECancelDummy)
    
    End If
    
    'Traz a Categoria para a Tela
    lErro = objRelOpcoes.ObterParametro("TCATEGORIAPROD", sParam)
    If lErro <> SUCESSO Then gError 128018

    For iIndice = 0 To CategoriaProdutos.ListCount - 1
        If Trim(CategoriaProdutos.List(iIndice)) = Trim(sParam) Then
            CategoriaProdutos.ListIndex = iIndice
            Exit For
        End If
    Next
    
    'Para Habilitar os Itens
    Call CategoriaProdutos_Click

    iCont = 1
    sParam = ""
    
    'Traz o Itemde da Categoria
    lErro = objRelOpcoes.ObterParametro("TITEMDE1", sParam)
    If lErro <> SUCESSO Then gError 128019
    
    Do While sParam <> ""
        
       For iIndice = 0 To ItensCategoriaProdutos.ListCount - 1
            If Trim(sParam) = Trim(ItensCategoriaProdutos.List(iIndice)) Then
                ItensCategoriaProdutos.Selected(iIndice) = True
                Exit For
            End If
        Next
        
        iCont = iCont + 1
        
        lErro = objRelOpcoes.ObterParametro("TITEMDE" & iCont, sParam)
        If lErro <> SUCESSO Then gError 128020

    Loop
    
    'Traz a CategoriaFornecedor para a Tela
    lErro = objRelOpcoes.ObterParametro("TCATEGORIAFORN", sParam)
    If lErro <> SUCESSO Then gError 128021

    For iIndice = 0 To CategoriaFornecedores.ListCount - 1
        If Trim(CategoriaFornecedores.List(iIndice)) = Trim(sParam) Then
            CategoriaFornecedores.ListIndex = iIndice
            Exit For
        End If
    Next
    
    'Para Habilitar os Itens
    Call CategoriaFornecedores_Click

    iCont = 1
    sParam = ""
    
    'Traz o Itemde da Categoria
    lErro = objRelOpcoes.ObterParametro("TITEMFORNDE1", sParam)
    If lErro <> SUCESSO Then gError 128022
    
    Do While sParam <> ""
        
       For iIndice = 0 To ItensCategoriaFornecedores.ListCount - 1
            If Trim(sParam) = Trim(ItensCategoriaFornecedores.List(iIndice)) Then
                ItensCategoriaFornecedores.Selected(iIndice) = True
                Exit For
            End If
        Next
        
        iCont = iCont + 1
        
        lErro = objRelOpcoes.ObterParametro("TITEMFORNDE" & iCont, sParam)
        If lErro <> SUCESSO Then gError 128023

    Loop
            
    'pega  Tipo Produto e Exibe
    lErro = objRelOpcoes.ObterParametro("NTIPOPRODUTOS", sParam)
    If lErro <> SUCESSO Then gError 128024
                   
    'Se não foi passado um tipo de produto
    If StrParaInt(sParam) = 0 Then
    
        'Seleciona a opção TipoProdutosTodos
        Call TipoProdutosTodos_Click
    
    'Senão, ou seja, se foi passado um tipo de produto
    Else
        
        'Seleciona a opção TipoProdutosApenas
        Call TipoProdutosApenas_Click
        
        'Percorre a combo, para selecionar o tipo
        For iIndice = 0 To TipoProdutos.ListCount - 1
            
            'Se o código do tipo for o mesmo código no itemdata => significa que encontrou o tipo
            If StrParaInt(sParam) = TipoProdutos.ItemData(iIndice) Then
                
                TipoProdutos.ListIndex = iIndice
                
                Exit For
            End If
        
        Next
        
    End If
    
    'Traz o Forncedorde para a tela
    lErro = objRelOpcoes.ObterParametro("NFORNECEDORDE", sParam)
    If lErro <> SUCESSO Then gError 128026
    
    FornecedorDe.PromptInclude = False
    FornecedorDe.Text = sParam
    FornecedorDe.PromptInclude = True
    Call FornecedorDe_Validate(bSGECancelDummy)
    
    'Traz o ForncedorAte para a tela
    lErro = objRelOpcoes.ObterParametro("NFORNECEDORATE", sParam)
    If lErro <> SUCESSO Then gError 128027
    
    FornecedorAte.PromptInclude = False
    FornecedorAte.Text = sParam
    FornecedorAte.PromptInclude = True
    Call FornecedorAte_Validate(bSGECancelDummy)
    
    'Traz o ForncedorTop para a tela
    lErro = objRelOpcoes.ObterParametro("NFORNECEDORESTOP", sParam)
    If lErro <> SUCESSO Then gError 128028
    
    FornTop.PromptInclude = False
    FornTop.Text = sParam
    FornTop.PromptInclude = True
    
    'Traz a CheckBox DetalharFilial
    lErro = objRelOpcoes.ObterParametro("NDETALHARFILIAL", sParam)
    If lErro <> SUCESSO Then gError 128053
    
    DetalharFilial.Value = StrParaInt(sParam)
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 128011 To 128028, 128053

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166833)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String, iCodFilialDe As Integer, iCodFilialAte As Integer, sCheckTipoProdutos As String, sTipoProdutos As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long
Dim sCclFormata As String
Dim iCclPreenchida As Integer
Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim iIndice As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros
       
    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoDe.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 128029

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoAte.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 128030

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then gError 128031

    End If
   
   If iCodFilialAte <> 0 Then
        
        'critica Codigo da Filial Inicial e Final
        If iCodFilialDe <> 0 And iCodFilialAte <> 0 Then
        
            If iCodFilialDe > iCodFilialAte Then gError 128032
        
        End If
   
   End If
   
    'data inicial não pode ser maior que a data final
    If Len(Trim(DataDe.ClipText)) <> 0 And Len(Trim(DataAte.ClipText)) <> 0 Then

         If StrParaDate(DataDe.Text) > StrParaDate(DataAte.Text) Then gError 128033

    End If
    
    'Se a opção para todos os Clientes estiver selecionada
    If TipoProdutosTodos.Value = True Then
        sCheckTipoProdutos = "Todos"
        sTipoProdutos = ""
    
    'Se a opção para apenas um Cliente estiver selecionada
    Else
        'TEm que indicar o tipo do Cliente
        If TipoProdutos.Text = "" Then gError 128034
        sCheckTipoProdutos = "Um"
        sTipoProdutos = TipoProdutos.Text
    
    End If
    
    'Verifica se o Fornecedor está Preenchido se Estiver
    'Fornecedor de  não pode ser maior que o Fornecedor até
    If Len(Trim(FornecedorDe.Text)) <> 0 And Len(Trim(FornecedorAte.Text)) <> 0 Then

        If StrParaLong(LCodigo_Extrai(FornecedorDe.Text)) > StrParaLong(LCodigo_Extrai(FornecedorAte.Text)) Then gError 128035

    End If
    If Len(Trim(CategoriaProdutos.Text)) <> 0 Then
    
        For iIndice = 0 To ItensCategoriaProdutos.ListCount - 1
            If ItensCategoriaProdutos.Selected(iIndice) = True Then
                Exit For
            End If
        Next
    
        If iIndice = ItensCategoriaProdutos.ListCount Then gError 128036
    
    End If
           
    'Verifica se a CatgoriaFornecedores foi preenchida
    If Len(Trim(CategoriaFornecedores.Text)) <> 0 Then
        
        '-> se foi algum item tem que estar selecionado
        For iIndice = 0 To ItensCategoriaFornecedores.ListCount - 1
            If ItensCategoriaFornecedores.Selected(iIndice) = True Then
                Exit For
            End If
        Next
    
        If iIndice = ItensCategoriaFornecedores.ListCount Then gError 128037
    
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                
        Case 128029, 128030
        
        Case 128031
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            
        Case 128032
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            
        Case 128033
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            
        Case 128034
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TIPOPRODUTO_NAO_PREENCHIDO", gErr)
            
        Case 128035
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_FINAL_MENOR", gErr)
            
        Case 128036
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_PRODUTO_ITEM_NAO_SELECIONADO", gErr)
            
        Case 128037
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAFORNECEDORITEM_ITEM_NAO_SELECIONADO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166834)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(objRelABCFornecedoresTela As ClassRelABCFornecedoresTela) As Long
'Move os elementos dsa tela para a memória

Dim iIndice As Integer
Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Move_Tela_Memoria

    'Passa o codigo do produto para o formato do BD
    lErro = CF("Produto_Formata", ProdutoDe.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 128038
    
    'Guarda o código do ProdutoDe
    objRelABCFornecedoresTela.sProdutoDe = sProdutoFormatado
    
    'Passa o codigo do produto para o formato do BD
    lErro = CF("Produto_Formata", ProdutoAte.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 128039
    
    'Guarda o código do ProdutoAte
    objRelABCFornecedoresTela.sProdutoAte = sProdutoFormatado
    
    'Preenche o Obj
    objRelABCFornecedoresTela.dtDataDe = StrParaDate(DataDe.Text)
    objRelABCFornecedoresTela.dtDataAte = StrParaDate(DataAte.Text)
    objRelABCFornecedoresTela.iFilialEmpresaDe = Codigo_Extrai(FilialEmpresaDe.Text)
    objRelABCFornecedoresTela.iFilialEmpresaAte = Codigo_Extrai(FilialEmpresaAte.Text)
    objRelABCFornecedoresTela.iTipoProduto = Codigo_Extrai(TipoProdutos.Text)
    objRelABCFornecedoresTela.lFornecedorDe = StrParaLong(FornecedorDe.Text)
    objRelABCFornecedoresTela.lFornecedorAte = StrParaLong(FornecedorAte.Text)
    objRelABCFornecedoresTela.sCategoriaProdutos = Trim(CategoriaProdutos.Text)
    objRelABCFornecedoresTela.sCategoriaFornecedores = Trim(CategoriaFornecedores.Text)
    objRelABCFornecedoresTela.iDetalharFilial = DetalharFilial.Value
    objRelABCFornecedoresTela.iFornecedorTop = StrParaInt(FornTop.ClipText)

    'Para cada item da categoria de produto selecionada
    For iIndice = 0 To ItensCategoriaProdutos.ListCount - 1
    
        'Se o item estiver marcado
        If ItensCategoriaProdutos.Selected(iIndice) = True Then
            
            'Guarda-o no obj
            objRelABCFornecedoresTela.colItensCategoriaProdutos.Add ItensCategoriaProdutos.List(iIndice)
        End If
        
    Next
    
    'Para cada item da categoria de fornecedor selecionada
    For iIndice = 0 To ItensCategoriaFornecedores.ListCount - 1
    
        'Se o item estiver marcado
        If ItensCategoriaFornecedores.Selected(iIndice) = True Then
            
            'Guarda-o no obj
            objRelABCFornecedoresTela.colItensCategoriaFornecedores.Add ItensCategoriaFornecedores.List(iIndice)
        End If
        
    Next
    
    Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr
    
        Case 128038, 128039
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166835)

    End Select

    Exit Function

End Function
'***** FUNÇÕES DE APOIO À TELA - FIM *****

'***** EVENTOS DO BROWSER - INÍCIO *****
Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim sProdutoMascarado As String
Dim lErro As Long

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 128040
    
    ProdutoAte.PromptInclude = False
    ProdutoAte.Text = sProdutoMascarado
    ProdutoAte.PromptInclude = True
    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr
    
        Case 128040
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166836)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoProdutoDe_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim sProdutoMascarado As String
Dim lErro As Long

On Error GoTo Erro_objEventoProdutoDe_evSelecao

    Set objProduto = obj1

    lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 128041
    
    ProdutoDe.PromptInclude = False
    ProdutoDe.Text = sProdutoMascarado
    ProdutoDe.PromptInclude = True
    
    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr
    
        Case 128041
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166837)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoFornDe_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor

    Set objFornecedor = obj1

    FornecedorDe.PromptInclude = False
    FornecedorDe.Text = objFornecedor.lCodigo
    FornecedorDe.PromptInclude = True

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoFornAte_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor

    Set objFornecedor = obj1

    FornecedorAte.PromptInclude = False
    FornecedorAte.Text = objFornecedor.lCodigo
    FornecedorAte.PromptInclude = True
    
    Me.Show

    Exit Sub

End Sub
'***** EVENTOS DO BROWSER - FIM *****

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_NF
    Set Form_Load_Ocx = Me
    Caption = "ABC de Fornecedores"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpABCFornecedores"

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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
    
        If Me.ActiveControl Is ProdutoDe Then
            Call LabelProdutoDe_Click
        ElseIf Me.ActiveControl Is ProdutoAte Then
            Call LabelProdutoAte_Click
        ElseIf Me.ActiveControl Is FornecedorDe Then
            Call LabelFornecedorDe_Click
        ElseIf Me.ActiveControl Is FornecedorAte Then
            Call LabelFornecedorAte_Click
        End If
    
    End If
        
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Sub Unload(objme As Object)
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
