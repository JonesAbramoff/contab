VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl VendaNFe 
   ClientHeight    =   9900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12030
   ScaleHeight     =   9900
   ScaleWidth      =   12030
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Height          =   2910
      Left            =   90
      TabIndex        =   16
      Top             =   915
      Width           =   7170
      Begin VB.CommandButton BotaoProxNum 
         Height          =   300
         Left            =   5880
         Picture         =   "VendaNFe.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Numeração Automática"
         Top             =   2550
         Width           =   300
      End
      Begin VB.OptionButton OptionPreVenda 
         Caption         =   "&Pré Venda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1260
         TabIndex        =   23
         Top             =   2565
         Width           =   1485
      End
      Begin VB.OptionButton OptionCF 
         Caption         =   "Venda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   105
         TabIndex        =   22
         Top             =   2535
         Width           =   1110
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   300
         Left            =   6360
         Picture         =   "VendaNFe.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Limpar"
         Top             =   2520
         Width           =   300
      End
      Begin VB.TextBox NomeCliente 
         Height          =   315
         Left            =   1440
         TabIndex        =   20
         Top             =   1050
         Width           =   3720
      End
      Begin VB.OptionButton OptionDAV 
         Caption         =   "&DAV"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2820
         TabIndex        =   19
         Top             =   2565
         Width           =   945
      End
      Begin VB.CommandButton BotaoMesclar 
         Caption         =   "M"
         Height          =   300
         Left            =   6765
         Picture         =   "VendaNFe.ctx":061C
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Mesclar PréVenda"
         Top             =   2535
         Width           =   300
      End
      Begin VB.CommandButton BotaoEndEntrega 
         Caption         =   "Entrega"
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
         Left            =   5520
         TabIndex        =   17
         Top             =   1305
         Width           =   1335
      End
      Begin MSMask.MaskEdBox CodVendedor 
         Height          =   315
         Left            =   1440
         TabIndex        =   25
         Top             =   255
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CodCliente 
         Height          =   300
         Left            =   1440
         TabIndex        =   26
         Top             =   660
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox Orcamento 
         Height          =   300
         Left            =   4800
         TabIndex        =   27
         Top             =   2550
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CGC 
         Height          =   315
         Left            =   1440
         TabIndex        =   28
         Top             =   1440
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##############"
         PromptChar      =   " "
      End
      Begin VB.Label LabelVendedor 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2160
         TabIndex        =   34
         Top             =   255
         Width           =   4155
      End
      Begin VB.Label LabelCodVendedor 
         AutoSize        =   -1  'True
         Caption         =   "&Vendedor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   33
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label LabelCliente 
         AutoSize        =   -1  'True
         Caption         =   "C&liente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   540
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   32
         Top             =   690
         Width           =   795
      End
      Begin VB.Label LabelOrcamento 
         AutoSize        =   -1  'True
         Caption         =   "&Número:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3795
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   31
         Top             =   2535
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   675
         TabIndex        =   30
         Top             =   1050
         Width           =   660
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Cnpj/Cpf:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   375
         TabIndex        =   29
         Top             =   1470
         Width           =   960
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11520
      Top             =   720
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   945
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   12030
      Begin VB.Label Exibe 
         BackStyle       =   0  'Transparent
         Caption         =   "ABC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   465
         Left            =   270
         TabIndex        =   15
         Top             =   -15
         Width           =   11625
      End
      Begin VB.Label Exibe1 
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   465
         Left            =   270
         TabIndex        =   14
         Top             =   465
         Width           =   11625
      End
   End
   Begin VB.CommandButton BotaoCancelaCupom 
      Caption         =   "(Esc)   Cancela Cupom"
      Height          =   375
      Left            =   405
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6105
      Width           =   1920
   End
   Begin VB.CommandButton BotaoProdutos 
      Caption         =   "(F9)   Produtos/Preço"
      Height          =   375
      Left            =   2670
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6105
      Width           =   1920
   End
   Begin VB.CommandButton BotaoPreco 
      Caption         =   "(F5)    Preço"
      Height          =   375
      Left            =   4980
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6105
      Width           =   1920
   End
   Begin VB.CommandButton BotaoCancelaItem 
      Caption         =   "(F6)    Cancela Item"
      Height          =   375
      Left            =   390
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6660
      Width           =   1920
   End
   Begin VB.CommandButton BotaoCancelaItemAtual 
      Caption         =   "(F4)   Cancela Item Atual"
      Height          =   375
      Left            =   2670
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6660
      Width           =   1920
   End
   Begin VB.CommandButton BotaoAbrirGaveta 
      Caption         =   "(F10)   Abrir Gaveta"
      Height          =   375
      Left            =   4965
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6660
      Width           =   1920
   End
   Begin VB.CommandButton BotaoPagamento 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   405
      Picture         =   "VendaNFe.ctx":0B4E
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   7170
      Width           =   6510
   End
   Begin VB.CommandButton BotaoSuspender 
      Caption         =   "(F7)  Suspender"
      Height          =   405
      Left            =   390
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Fechar"
      Top             =   7905
      Width           =   1920
   End
   Begin VB.CommandButton BotaoAtualizar 
      Caption         =   "(Ctrl+F1)  Atualizar Tabelas"
      Height          =   375
      Left            =   2610
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Fechar"
      Top             =   7920
      Width           =   2085
   End
   Begin VB.CommandButton BotaoFechar 
      Height          =   375
      Left            =   4995
      Picture         =   "VendaNFe.ctx":21B0
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Fechar"
      Top             =   7920
      Width           =   1905
   End
   Begin VB.PictureBox Picture1 
      Height          =   8340
      Left            =   6900
      ScaleHeight     =   8280
      ScaleWidth      =   4320
      TabIndex        =   2
      Top             =   5100
      Visible         =   0   'False
      Width           =   4380
   End
   Begin VB.ListBox ListCF 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   5295
      ItemData        =   "VendaNFe.ctx":27AE
      Left            =   7455
      List            =   "VendaNFe.ctx":27B0
      TabIndex        =   0
      Top             =   3825
      Width           =   4380
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5700
      Left            =   7440
      TabIndex        =   1
      Top             =   3855
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   10054
      _Version        =   393216
      Rows            =   1000
      Cols            =   9
      FixedRows       =   2
      FixedCols       =   0
      BackColor       =   12648447
      BackColorFixed  =   12648447
      GridColor       =   12648447
      AllowBigSelection=   -1  'True
      FocusRect       =   0
      HighLight       =   2
      GridLines       =   0
      GridLinesFixed  =   0
      ScrollBars      =   2
      SelectionMode   =   1
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSCommLib.MSComm LeitoraCodBarras 
      Left            =   4515
      Top             =   3810
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSMask.MaskEdBox ProdutoNomeRed 
      Height          =   585
      Left            =   2505
      TabIndex        =   35
      Top             =   4665
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   1032
      _Version        =   393216
      PromptInclude   =   0   'False
      HideSelection   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Quantidade 
      Height          =   585
      Left            =   2520
      TabIndex        =   36
      Top             =   3930
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1032
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   " "
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   330
      Top             =   4290
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RT1 
      Height          =   525
      Left            =   90
      TabIndex        =   37
      Top             =   1020
      Visible         =   0   'False
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   926
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"VendaNFe.ctx":27B2
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   5700
      Left            =   7470
      TabIndex        =   38
      Top             =   1035
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   10054
      _Version        =   393216
      Rows            =   20
      Cols            =   1
      FixedRows       =   19
      FixedCols       =   0
      BackColor       =   12648447
      BackColorFixed  =   12648447
      GridColor       =   12648447
      AllowBigSelection=   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   0
      GridLinesFixed  =   0
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Subtotal:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   825
      TabIndex        =   44
      Top             =   5445
      Width           =   1560
   End
   Begin VB.Label PrecoUnitario 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   8865
      TabIndex        =   43
      Top             =   4455
      Width           =   1800
   End
   Begin VB.Label PrecoItem 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   8685
      TabIndex        =   42
      Top             =   6120
      Width           =   1920
   End
   Begin VB.Label Subtotal 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2490
      TabIndex        =   41
      Top             =   5400
      Width           =   4425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Quantidade:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   270
      TabIndex        =   40
      Top             =   3975
      Width           =   2145
   End
   Begin VB.Label LabelProduto 
      AutoSize        =   -1  'True
      Caption         =   "&Produto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   900
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   39
      Top             =   4665
      Width           =   1500
   End
   Begin VB.Image Figura 
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   210
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   6750
   End
End
Attribute VB_Name = "VendaNFe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Global
Dim giItem As Integer
Dim gobjVenda As ClassVenda
Dim iAlterado As Integer
Dim gsMinutoAnt As String
Dim gsNomeOperador As String
Dim giLarguraOrig As Integer
Dim giAlturaOrig As Integer
Dim giLarguraListCF As Integer
Dim giAlturaListCF As Integer
Dim giAlturaFigura As Integer
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim giMaxCaracteres As Integer
Dim giLarguraGrid As Integer
Dim giAlturaGrid As Integer
Dim giUltimaLinhaGrid As Integer
Dim giLinhasVisiveisGrid As Integer
Const GRID_COL_ITEM = 0
Const GRID_COL_CODIGO = 1
Const GRID_COL_DESCRICAO = 2
Const GRID_COL_QUANTIDADE = 3
Const GRID_COL_X = 4
Const GRID_COL_VALOR_UNITARIO = 5
Const GRID_COL_ST = 6
Const GRID_COL_VALOR_TOTAL = 7




Public Sub Form_Load()
        
On Error GoTo Erro_Form_Load
Dim objOperador As New ClassOperador
Dim lErro As Long
Dim sTexto As String
Dim objOrcamento As Object
Dim objTela As Object

On Error GoTo Erro_Form_Load
        
    giLarguraOrig = Me.Width
    giAlturaOrig = Me.Height
    
    giLarguraListCF = ListCF.Width
    giAlturaListCF = ListCF.Height
    
    giLarguraGrid = Grid.Width
    giAlturaGrid = Grid.Height
    
    giAlturaFigura = Figura.Height
        
    gobjLojaECF.dtTime = CDate(Now)
        
'    Apresentacao.Caption = Formata_Campo(ALINHAMENTO_DIREITA, 50, " ", gsNomeEmpresa)
    Exibe.Caption = "PRÓXIMO CLIENTE...."
    Exibe1.Caption = ""
    
    UserControl.Parent.WindowState = 2
    giItem = 0
    
    Set gobjVenda = New ClassVenda
    
    
    If giStatusCaixa = STATUS_CAIXA_FECHADO And giOrcamentoECF <> CAIXA_SO_ORCAMENTO Then gError 133827
    
    'Se a Sessão Estiver Fechada então gera Error
    If giStatusSessao = SESSAO_ENCERRADA Then gError 99827

    'Se Sessão estiver Suspensa
    If giStatusSessao = SESSAO_SUSPENSA Then gError 99828
    
    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 210885
    
    'Inicia o Cupom
'    Call Inicia_Cupom_Tela
        
    Quantidade.Text = 1
        
    PrecoUnitario.Caption = Format(0, "Standard")
    Subtotal.Caption = Format(0, "Standard")
    PrecoItem.Caption = Format(0, "Standard")
    
    Orcamento.Text = ""
'    Orcamento.Enabled = True
'    LabelOrcamento.Enabled = True
    BotaoProxNum.Enabled = True
    
    If giPreVenda = 0 Then
        OptionPreVenda.Enabled = False
    End If
    
    If giDAV = 0 Then
        OptionDAV.Enabled = False
    End If
    
    If giPreVenda = 1 And giUsaImpressoraFiscal = 0 Then
        OptionPreVenda.Value = True
    ElseIf giDAV = 1 Then
        OptionDAV.Value = True
    End If
        
        
    If giOrcamentoECF = CAIXA_SO_ORCAMENTO Then
        
        OptionCF.Enabled = False
        If giDAV = 1 Then
            BotaoCancelaCupom.Caption = "(Esc)   Imprime Orçamento"
        ElseIf giPreVenda = 1 Then
            BotaoCancelaCupom.Caption = "(Esc)   Limpa Tela"
        End If
        BotaoAbrirGaveta.Caption = "(F10)   Grava Orçamento"
    Else
        OptionCF.Value = True
    End If
    
    
    'Seleiona o nome do operador
    For Each objOperador In gcolOperadores
        If objOperador.iCodigo = giCodOperador Then gsNomeOperador = objOperador.sNome
    Next
    
    'Função do AFRAC que informa o Operador
    lErro = AFRAC_InformarOperador(gsNomeOperador)
    If lErro <> SUCESSO Then
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Informa Operador")
        If lErro <> SUCESSO Then gError 99920
    End If
    
    Quantidade.SelStart = 0
    Quantidade.SelLength = Len(Quantidade.Text)
    
    If giDinheiroAtivo = ADMMEIOPAGTOCONDPAGTO_INATIVO Then
        BotaoAbrirGaveta.Enabled = False
    End If
    
    Set objTela = Me
    
    lErro = CF_ECF("Inicializa_LeitoraCodBarras", objTela)
    If lErro <> SUCESSO Then gError 117684
    
    Call Timer1_Timer
    
        
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 99827
            Call Rotina_ErroECF(vbOKOnly, ERRO_SESSAO_ABERTA_INEXISTENTE, gErr, giCodCaixa)

        Case 99828
            Call Rotina_ErroECF(vbOKOnly, ERRO_SESSAO_SUSPENSA, gErr, giCodCaixa)
        
        Case 99920, 117684, 199463, 210885
        
        Case 133827
            Call Rotina_ErroECF(vbOKOnly, ERRO_CAIXA_FECHADO, gErr, giCodCaixa)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175669)

    End Select

    Exit Sub

End Sub

Public Sub Inicia_Cupom_Tela()
'Joga o endereço, Nome da Empresa, CNPJ, IE no Cupom

Dim iIndice As Integer
Dim lErro As Long
Dim lLargura As Long
Dim tSize As typeSize
Dim tSize1 As typeSize
    
'    lErro = GetTextExtentPoint32(Parent.hdc, Formata_Campo(ALINHAMENTO_CENTRALIZADO, 300, "*", gsNomeEmpresa), 300, tSize)
        
'        Me.Width = .Width + (1440 * 2 * GetSystemMetrics(SM_CXFIXEDFRAME) / GetDeviceCaps(Me.hdc, LOGPIXELSX))
    
    If Grid.Width < 8000 Then
    
    
'        giMaxCaracteres = ControleMaxTamVisivel(Picture1, "*")
'
'        ListCF.AddItem Formata_Campo(ALINHAMENTO_CENTRALIZADO, giMaxCaracteres, "*", gsNomeEmpresa)
'
'        giMaxCaracteres = ControleMaxTamVisivel(Picture1, " ")
'
'        ListCF.AddItem Formata_Campo(ALINHAMENTO_CENTRALIZADO, giMaxCaracteres, " ", gsEndereco)
'        ListCF.AddItem Formata_Campo(ALINHAMENTO_CENTRALIZADO, giMaxCaracteres, " ", gsEnderecoComplemento)
'        ListCF.AddItem ""
'        ListCF.AddItem "CNPJ/CPF:" & Formata_Campo(ALINHAMENTO_DIREITA, 54, " ", gsCNPJ)
'        ListCF.AddItem "I.E.:" & Formata_Campo(ALINHAMENTO_DIREITA, 54, " ", gsInscricaoEstadual)
'    '    ListCF.AddItem TRACO_CAB
'        giMaxCaracteres = ControleMaxTamVisivel(Picture1, "-")
'        ListCF.AddItem Formata_Campo(ALINHAMENTO_CENTRALIZADO, giMaxCaracteres, "-*", "-")
'        ListCF.AddItem ""
'        giMaxCaracteres = ControleMaxTamVisivel(Picture1, " ")
'        ListCF.AddItem Formata_Campo(ALINHAMENTO_CENTRALIZADO, giMaxCaracteres, " ", "  ECF :" & FormataCpoNum(giCodECF, 4) & "           LJ :" & FormataCpoNum(giFilialEmpresa, 4) & "          OP :" & FormataCpoNum(giCodOperador, 4))
'        ListCF.AddItem ""
'        ListCF.AddItem Formata_Campo(ALINHAMENTO_CENTRALIZADO, giMaxCaracteres, " ", "CUPOM FISCAL")
'        If giMaxCaracteres > 100 Then
'            ListCF.AddItem Formata_Campo(ALINHAMENTO_CENTRALIZADO, giMaxCaracteres, " ", "Item    Codigo         Descrição                                Qtd  Un. x Unitário      ST            Valor(" & gobjLojaECF.sSimboloMoeda & ")")
'        Else
'           ListCF.AddItem Formata_Campo(ALINHAMENTO_DIREITA, 54, " ", "Item    Codigo         Descrição")
'            ListCF.AddItem Formata_Campo(ALINHAMENTO_DIREITA, 54, " ", "        Qtd  Un. x Unitário      ST            Valor(" & gobjLojaECF.sSimboloMoeda & ")")
'        End If
    '    ListCF.AddItem TRACO_CAB
'        ListCF.AddItem Formata_Campo(ALINHAMENTO_DIREITA, 54, " ", "Item    Codigo         Descrição")
'        ListCF.AddItem Formata_Campo(ALINHAMENTO_DIREITA, 54, " ", "        Qtd  Un. x Unitário      ST            Valor(" & gobjLojaECF.sSimboloMoeda & ")")
'        giMaxCaracteres = ControleMaxTamVisivel(Picture1, "-")
'        ListCF.AddItem Formata_Campo(ALINHAMENTO_CENTRALIZADO, giMaxCaracteres, "-*", "-")
    
    Else
    
        Grid.Row = Grid.FixedRows
        Grid.Col = Grid.FixedCols
        Grid.RowSel = Grid.Rows - 1
        Grid.ColSel = Grid.Cols - 1
        Grid.FillStyle = flexFillRepeat
        Grid.Text = ""
        Grid.FillStyle = flexFillSingle
        
        Grid.Row = 0
        Grid.Col = 0
        
        Grid.ScrollBars = flexScrollBarVertical
    
        Grid.TopRow = 2
        
        Grid.TextMatrix(0, GRID_COL_ITEM) = "Item"
        Grid.TextMatrix(0, GRID_COL_CODIGO) = "Codigo"
        Grid.TextMatrix(0, GRID_COL_DESCRICAO) = "Descrição"
        Grid.TextMatrix(0, GRID_COL_QUANTIDADE) = "Qtd. Un."
        Grid.TextMatrix(0, GRID_COL_X) = "x"
        Grid.TextMatrix(0, GRID_COL_VALOR_UNITARIO) = "Vl.Unit."
        Grid.TextMatrix(0, GRID_COL_ST) = "ST"
        Grid.TextMatrix(0, GRID_COL_VALOR_TOTAL) = "Valor(" & gobjLojaECF.sSimboloMoeda & "$)"
        
        Call GetTextExtentPoint32(Picture1.hdc, "00000", 5, tSize)
        
        Call GetTextExtentPoint32(Picture1.hdc, "xxx", 3, tSize1)
        
        lLargura = (Grid.Width - (tSize1.cx * Screen.TwipsPerPixelX)) / 7
        
        giMaxCaracteres = ControleMaxTamVisivel(Picture1, "-")
        
        For iIndice = 0 To 7
            Grid.TextMatrix(1, iIndice) = Formata_Campo(ALINHAMENTO_CENTRALIZADO, giMaxCaracteres + 100, "-*", "-")
            If iIndice <> 4 Then Grid.ColWidth(iIndice) = (Grid.Width - ((GetSystemMetrics(SM_CXVSCROLL) + tSize1.cx) * Screen.TwipsPerPixelX)) / 7
            Grid.ColAlignment(iIndice) = flexAlignCenterCenter
        Next
    
    
        
        Grid.ColWidth(GRID_COL_ITEM) = tSize.cx * (Screen.TwipsPerPixelX)
        Grid.ColWidth(GRID_COL_DESCRICAO) = lLargura + lLargura - (tSize.cx * (Screen.TwipsPerPixelX))
        Grid.ColWidth(GRID_COL_X) = tSize1.cx * (Screen.TwipsPerPixelX)
        Grid.ColWidth(8) = 0
        
        giLinhasVisiveisGrid = Grid.Height / Grid.RowHeight(2)
        
        giUltimaLinhaGrid = 1
    
        Grid.Rows = giLinhasVisiveisGrid + 1
        
    End If
    
    Grid1.ColWidth(0) = Grid1.Width
    Grid1.ColAlignment(0) = flexAlignCenterCenter
        
'        Grid.Row = 0
    
        
'    For iIndice = 0 To 10
'        Grid.Col = iIndice
''        Grid.CellWidth = Grid.Width
'        Grid.CellAlignment = flexAlignCenterCenter
'    Next
    
    Grid1.TextMatrix(0, 0) = Formata_Campo(ALINHAMENTO_CENTRALIZADO, giMaxCaracteres, "*", gsNomeEmpresa)
    Grid1.TextMatrix(1, 0) = gsEndereco
    Grid1.TextMatrix(2, 0) = gsEnderecoComplemento
    Grid1.Row = 4
    Grid1.CellAlignment = flexAlignLeftCenter
    Grid1.TextMatrix(4, 0) = "CNPJ/CPF: " & gsCNPJ
    Grid1.Row = 5
    Grid1.CellAlignment = flexAlignLeftCenter
    Grid1.TextMatrix(5, 0) = "I.E.: " & gsInscricaoEstadual
    Grid1.Row = 6
    Grid1.CellAlignment = flexAlignCenterCenter
    giMaxCaracteres = ControleMaxTamVisivel(Picture1, "-")
    Grid1.TextMatrix(6, 0) = Formata_Campo(ALINHAMENTO_CENTRALIZADO, giMaxCaracteres, "-*", "-")
    giMaxCaracteres = ControleMaxTamVisivel(Picture1, " ")
    Grid1.TextMatrix(7, 0) = "  ECF :" & FormataCpoNum(giCodECF, 4) & "           LJ :" & FormataCpoNum(giFilialEmpresa, 4) & "          OP :" & FormataCpoNum(giCodOperador, 4)
    Grid1.TextMatrix(9, 0) = "CUPOM FISCAL"
    
    If Grid1.Width < 8000 Then
        
        Grid1.Row = 11
        Grid1.CellAlignment = flexAlignLeftCenter
        Grid1.TextMatrix(11, 0) = "Item    Codigo         Descrição"
        Grid1.Row = 12
        Grid1.CellAlignment = flexAlignLeftCenter
        Grid1.TextMatrix(12, 0) = "        Qtd  Un. x Unitário      ST            Valor(" & gobjLojaECF.sSimboloMoeda & "$)"
'        ListCF.AddItem Formata_Campo(ALINHAMENTO_DIREITA, 54, " ", "Item    Codigo         Descrição")
'        ListCF.AddItem Formata_Campo(ALINHAMENTO_DIREITA, 54, " ", "        Qtd  Un. x Unitário      ST            Valor(" & gobjLojaECF.sSimboloMoeda & ")")
    
    End If
    
End Sub

Function Trata_Parametros() As Long
   
    Trata_Parametros = SUCESSO


End Function

Private Sub BotaoAtualizar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoAtualizar_Click

    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 210004

    Parent.MousePointer = vbHourglass
    
    lErro = CF_ECF("Carrega_Arquivo_FonteDados", 1)
    If lErro <> SUCESSO Then gError 133561

    Parent.MousePointer = vbDefault

    Call Rotina_AvisoECF(vbOKOnly, AVISO_TABELAS_ATUALIZADAS)
    
    Exit Sub

Erro_BotaoAtualizar_Click:

    Select Case gErr
                
        Case 133561, 133678
            Parent.MousePointer = vbDefault
                            
        Case 210004
                            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175670)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoEndEntrega_Click()
    Call Chama_TelaECF_Modal("EnderecoEntrega", gobjVenda)
End Sub

Private Sub BotaoFechar_Click()
    
    Unload Me
        
End Sub

Private Sub BotaoLimpar_Click()
    
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer
Dim lErro As Long
Dim lNum As Long
Dim objOperador As New ClassOperador
Dim iCodGerente As Integer

On Error GoTo Erro_Botaolimpar_Click
    
    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 210009
    
    lNum = Retorna_Count_ItensCupom
    
    'Se tiver uma venda acontecendo
    If lNum > 0 And OptionCF Then
        'Envia aviso perguntando se deseja cancelar a venda
        vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_CANCELA_VENDA)

        If vbMsgRes = vbYes Then
            'Se for Necessário a autorização do Gerente para abertura do Caixa
            If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then
        
                'Chama a Tela de Senha
                Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)
        
                'Sai de Função se a Tela de Login não Retornar ok
                If giRetornoTela <> vbOK Then gError 102506
                
                'Se Operador for Gerente
                iCodGerente = objOperador.iCodigo
        
            End If

            'Cancelar o Cupom de Venda
            lErro = AFRAC_CancelarCupom(Me, gobjVenda)
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Cancelar Cupom")
            If lErro <> SUCESSO Then gError 99617
            
            Call Move_Dados_Memoria_1
            
            'Realiza as operações necessárias para gravar
            lErro = CF_ECF("Grava_Venda_Arquivo", gobjVenda)
            If lErro <> SUCESSO Then gError 204527
            
            lErro = CF_ECF("Grava_CancelamentoCupom_Arquivo", gobjVenda)
            If lErro <> SUCESSO Then gError 204528
            
            
        Else
            Exit Sub
        End If
        
    End If
    
    'Se tiver um orçamento na tela
    If lNum > 0 And (OptionPreVenda.Value = True Or OptionDAV.Value = True) Then
        
        'Envia aviso que o orçamento será cancelado
        vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_CANCELA_ORCAMENTO)

        If vbMsgRes = vbNo Then Exit Sub
            
    End If
        
    Set gobjVenda = Nothing
    Set gobjVenda = New ClassVenda
        
    Call Limpa_Tela_Venda
    
    Exit Sub

Erro_Botaolimpar_Click:

    Select Case gErr
                
        Case 99617, 102506, 204527, 204528, 210009
                            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175671)

    End Select
    
    Exit Sub

End Sub

Private Sub codCliente_Change()
'Determina se Houve Mudança

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodCliente_GotFocus()
    Call MaskEdBox_TrataGotFocus(CodCliente, iAlterado)
End Sub

Private Sub CodCliente_Validate(Cancel As Boolean)

Dim objCliente As ClassCliente
Dim lIndice As Long
Dim lErro As Long

On Error GoTo Erro_Cliente_Validate
    
    If Len(Trim(CodCliente.ClipText)) > 0 Then
    
        If IsNumeric(CodCliente.ClipText) Then
    
            Set objCliente = gobjClienteCPF.Busca(CodCliente.ClipText, lIndice)
    
            If objCliente Is Nothing Then gError 126809
    
            Select Case Len(Trim(CodCliente.ClipText))
        
                Case STRING_CPF 'CPF
                    
                    'Critica Cpf
                    lErro = Cpf_Critica(CodCliente.ClipText)
                    If lErro <> SUCESSO Then gError 126806
                    
                    'Formata e coloca na Tela
                    CodCliente.Format = "000\.000\.000-00; ; ; "
        
                Case STRING_CGC 'CGC
                    
                    'Critica CGC
                    lErro = Cgc_Critica(CodCliente.ClipText)
                    If lErro <> SUCESSO Then gError 126807
                    
                    'Formata e Coloca na Tela
                    CodCliente.Format = "00\.000\.000\/0000-00; ; ; "
        
                Case Else
                    gError 126808
        
            End Select
    
        Else
    
            Set objCliente = gobjClienteNome.Busca(CodCliente.ClipText, lIndice)
            
            If objCliente Is Nothing Then gError 126804
            
        End If
    
        'joga o cliente no gobjvenda
        gobjVenda.objCupomFiscal.sCPFCGC = objCliente.sCgc
            
        CodCliente.Text = objCliente.sCgc

        If Len(Trim(CGC.ClipText)) = 0 Then
            CGC.Text = objCliente.sCgc
        End If
        
        If Len(Trim(NomeCliente.Text)) = 0 Then
            NomeCliente.Text = objCliente.sNomeReduzido
            gobjVenda.objCupomFiscal.sNomeCliente = objCliente.sNomeReduzido
        End If
    
    
    End If
    
    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr
        
        Case 126804
            Call Rotina_ErroECF(vbOKOnly, ERRO_CLIENTE_NAO_CADASTRADO3, gErr, CodCliente.Text)
        
        Case 126806, 126807
        
        Case 126808
            Call Rotina_Erro(vbOKOnly, ERRO_TAMANHO_CGC_CPF1, gErr)
        
        Case 126809
            Call Rotina_ErroECF(vbOKOnly, ERRO_CPFCNPJ_NAO_CADASTRADO, gErr, CodCliente.Text)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175672)
                        
    End Select
   
    Exit Sub

End Sub


Private Sub CodVendedor_GotFocus()
    Call MaskEdBox_TrataGotFocus(CodVendedor, iAlterado)
End Sub

Private Sub LabelCodVendedor_Click()

Dim objVendedor As New ClassVendedor
    
    'Chama tela de vendedorLista
    Call Chama_TelaECF_Modal("VendedorLista", objVendedor)
    
    If giRetornoTela = vbOK Then
        CodVendedor.Text = objVendedor.iCodigo
        Call CodVendedor_Validate(False)
    End If
    
    Exit Sub

End Sub

Private Sub LabelVendedor_Click()

Dim objVendedor As New ClassVendedor
    
    'Chama tela de vendedorLista
    Call Chama_TelaECF_Modal("VendedorLista", objVendedor)
    
    If giRetornoTela = vbOK Then
        CodVendedor.Text = objVendedor.iCodigo
        Call CodVendedor_Validate(False)
    End If
    
    Exit Sub

End Sub

Private Sub LabelOrcamento_Click()

Dim objOrcamento As New ClassOrcamentoLoja
Dim objVenda As New ClassVenda
Dim colOrcamento As New Collection
Dim iAchou As Integer
Dim lErro As Long
Dim objItens As ClassItemCupomFiscal
Dim iIndice As Integer
Dim objProduto As ClassProduto


On Error GoTo Erro_LabelOrcamento_Click

    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 210011


    'Chama tela de OrçamentoLista
    Call Chama_TelaECF_Modal("OrcamentoLista", objOrcamento)
    
    If giRetornoTela = vbOK Then
        
'        If giPreVenda = 1 And giUsaImpressoraFiscal = 0 Then
'            If Not OptionPreVenda.Value Then OptionPreVenda.Value = True
'        ElseIf giDAV = 1 Then
'            If Not OptionDAV.Value Then OptionDAV.Value = True
'        End If
        
        objVenda.objCupomFiscal.lNumOrcamento = objOrcamento.lNumOrcamento
        
        'Função Que le os orcamentos
        lErro = CF_ECF("OrcamentoECF_Le", objVenda)
        If lErro <> SUCESSO And lErro <> 204690 And lErro <> 210447 Then gError 105857
        
        If lErro = 210447 Then gError 210452
        
        'orcamento nao cadastrado
        If lErro <> SUCESSO Then gError 105858
        
        'descobre o nome reduzido do produto
        For Each objItens In objVenda.objCupomFiscal.colItens
            For iIndice = 1 To gaobjProdutosNome.Count
                Set objProduto = gaobjProdutosNome.Item(iIndice)
                If objItens.sProduto = objProduto.sCodigo Then
                    objItens.sProdutoNomeRed = objProduto.sNomeReduzido
                    Exit For
                End If
            Next
        Next
        
        'Traz ele para a tela
        Call Copia_Venda(gobjVenda, objVenda)
        Call Traz_Orcamento
        
        'se o cupom fiscal estiver ligado, cham OptionCF_Click para transformar o orcamento em cupom
        If OptionCF.Value Then Call OptionCF_Click
        
    End If
    
    Exit Sub

Erro_LabelOrcamento_Click:

    Select Case gErr

        Case 105857, 210011

        Case 105858
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_NAO_CADASTRADO1, gErr, objOrcamento.lNumOrcamento)

        Case 210452
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_BAIXADO, gErr, objVenda.objCupomFiscal.lNumOrcamento)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175673)

    End Select

    Exit Sub

End Sub

Public Sub LabelProduto_Click()

    Call BotaoProdutos_Click

End Sub

Public Sub BotaoProdutos_Click()
'Chama o browser do ProdutoLojaLista
'So traz produtos onde codigo de barras ou referencia está preenchida

Dim objProduto As New ClassProduto

On Error GoTo Erro_BotaoProdutos_Click
    
    objProduto.sNomeReduzido = ProdutoNomeRed.Text
    
    'Chama tela de ProdutosLista
    Call Chama_TelaECF_Modal("ProdutosLista", objProduto)
        
    UserControl.Refresh
    
    If giRetornoTela = vbOK Then
        If Len(Trim(objProduto.sReferencia)) > 0 Then
            ProdutoNomeRed.Text = objProduto.sReferencia
        Else
            ProdutoNomeRed.Text = objProduto.sCodigoBarras
        End If
        Call ProdutoNomeRed_Validate(False)
    End If
    
    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175674)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()
    
Dim lErro As Long
Dim lNumero As Long

On Error GoTo Erro_BotaoProxNum_Click

    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 210010

    If gobjVenda.iTipo <> OPTION_DAV And gobjVenda.iTipo <> OPTION_PREVENDA Then Exit Sub
    
    'Função que obtém o próximo número
    lErro = CF_ECF("Venda_Obtem_Num_Automatico", lNumero)
    If lErro <> SUCESSO Then gError 99901

    'joga na tela
    Orcamento.Text = lNumero
    gobjVenda.objCupomFiscal.lNumOrcamento = lNumero
    
    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr
        
        Case 99901, 210010
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175675)

    End Select

    Exit Sub

End Sub

Private Sub BotaoSuspender_Click()

Dim lErro As Long
Dim lNumItens As Long
Dim sCaption As String
Dim sCaption1 As String

On Error GoTo Erro_BotaoSuspender_Click
    
    'Retorna o count de gobjvenda.objcupomfiscal.colitens
    lNumItens = Retorna_Count_ItensCupom()
    
    'Se tiver uma venda acontecendo -> erro.
    If lNumItens > 0 And OptionCF Then gError 99903
          
    sCaption = Exibe.Caption
    sCaption1 = Exibe1.Caption
          
    Exibe.Caption = "AGUARDANDO RETORNO..."
    Exibe1.Caption = ""
    'Função que Executa a Suspenção da Sessão
    lErro = CF_ECF("Sessao_Executa_Suspensao")
    If lErro <> SUCESSO Then gError 99826
    
    'funcao que executa o termino da suspensao se a senha for digitada.
    lErro = CF_ECF("Sessao_Executa_Termino_Susp")
    If lErro <> SUCESSO Then gError 117546
    
    Exibe.Caption = sCaption
    Exibe1.Caption = sCaption1
    DoEvents
    
    Exit Sub

Erro_BotaoSuspender_Click:

    Select Case gErr
                
        Case 99826, 117546
        
        Case 99903
            Call Rotina_ErroECF(vbOKOnly, ERRO_VENDA_ANDAMENTO, gErr)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175676)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoCancelaCupom_Click()
    
Dim lErro As Long
Dim iCodigo As Integer
Dim lNumItens As Long
Dim iIndice As Integer
Dim objItens As New ClassItemCupomFiscal
Dim objAliquota As New ClassAliquotaICMS
Dim objVenda As New ClassVenda
Dim sRetorno As String
Dim vbMsgRes As VbMsgBoxResult
Dim lSequencial As Long
Dim colRegistro As New Collection
Dim sLog As String
Dim objCliente As ClassCliente
Dim sCPF As String
Dim lNumero As Long
Dim objOperador As New ClassOperador
Dim iCodGerente As Integer
Dim iFlag As Integer
Dim lRetorno As Long
Dim dtDataFinal As Date
Dim objTela As Object

On Error GoTo Erro_BotaoCancelaCupom_Click
    
    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 210000
    
      
    If gobjVenda.iTipo = OPTION_PREVENDA Then
        Call BotaoLimpar_Click
    Else
  
        'se se trata de um orcamento DAV
        If gobjVenda.iTipo = OPTION_DAV Then
        
            If gobjVenda.objCupomFiscal.lNumOrcamento = 0 Then gError 105888
        
            objVenda.objCupomFiscal.lNumOrcamento = gobjVenda.objCupomFiscal.lNumOrcamento
    
            'le o orcamento em questao
            lErro = CF_ECF("OrcamentoECF_Le", objVenda)
            If lErro <> SUCESSO And lErro <> 204690 And lErro <> 210447 Then gError 105867
            
            If lErro = 210447 Then gError 210453
            
            'se o orcamento ja esta cadastrado ==> vai imprimir o DAV se nao tiver ja sido impresso
            If lErro = SUCESSO Then
        
            
                    'se for dav é ja tiver sido impresso ==> nao imprime nem altera o DAV
                    If gobjVenda.iTipo = OPTION_DAV And objVenda.objCupomFiscal.iDAVImpresso <> 0 Then gError 210991
            

                    Set objTela = Me
    
                    gobjVenda.objCupomFiscal.dtDataEmissao = Date
                    gobjVenda.objCupomFiscal.dHoraEmissao = CDbl(Time)
                    gobjVenda.objCupomFiscal.dtDataReducao = gdtDataAnterior
    
                    'le os registros do orcamento e loca o arquivo
                    lErro = CF_ECF("Imprime_OrcamentoECF", dtDataFinal, objVenda.objCupomFiscal.lNumOrcamento, objTela, gobjVenda)
                    If lErro <> SUCESSO Then gError 105886
            
                    Set gobjVenda = Nothing
                    Set gobjVenda = New ClassVenda
            
            End If
                
'                Set gobjVenda = New ClassVenda
                
'            Else
'
'                'permite limpar a tela caso nao esteja cadastrado
'                vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_EXCLUSAO_ORCAMENTO_ECF)
'
'                If vbMsgRes = vbNo Then Exit Sub
'
'                Set gobjVenda = New ClassVenda
'
'            End If
        
        
        Else
        
            'se for um cupom e o
            'cupom a ser cancelado é um anterior(naum está na tela)
            If gobjVenda.objCupomFiscal.lNumero = 0 Then
                If gcolVendas.Count = 0 Then gError 112075
                For iIndice = gcolVendas.Count To 1 Step -1
                    Set gobjVenda = gcolVendas.Item(iIndice)
                    If gobjVenda.iTipo = OPTION_CF Then
                    
                        If AFRAC_ImpressoraCFe(giCodModeloECF) Then
                            '???????
                            lRetorno = gobjVenda.objCupomFiscal.lNumero
                        
                        Else
                    
                            'COO atual
                            lErro = AFRAC_LerInformacaoImpressora("023", sRetorno)
                            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Informação Impressora")
                            If lErro <> SUCESSO Then gError 112061
                            
                            lRetorno = StrParaLong(sRetorno)
                            
                            If giCodModeloECF <> 7 And giCodModeloECF <> 6 Then
                            'este teste tem que existir pois daruma e bematech quando fechado o cupom
                            'retorna o proximo numero de COO enquanto elgin retorna o atual mesmo q estiver
                            'fechado o cupom
                                lErro = AFRAC_Le_FlagsFiscais(iFlag)
                                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Informação Impressora")
                                If lErro <> SUCESSO Then gError 199581
                                
                                
                                If Not (iFlag And 1) Then
                                    lRetorno = lRetorno - 1
                                End If
                            
                            End If
                        
                        End If
                        
                        'se o último númeor de cupom é o da última venda executada--> pode cancelar esta venda
                        If lRetorno = gobjVenda.objCupomFiscal.lNumero Then
                        
                            'Envia aviso perguntando se deseja cancelar o cupom
                            vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_CANCELA_CUPOM_ANTERIOR)
                    
                            If vbMsgRes = vbNo Then Exit Sub
                            
                            'Se for Necessário a autorização do Gerente para abertura do Caixa
                            If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then
                        
                                'Chama a Tela de Senha
                                Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)
                        
                                'Sai de Função se a Tela de Login não Retornar ok
                                If giRetornoTela <> vbOK Then gError 102501
                                
                                'Se Operador for Gerente
                                iCodGerente = objOperador.iCodigo
                            
                            End If
                            
                            Exibe.Caption = "CANCELADO CUPOM " & gobjVenda.objCupomFiscal.lNumero
                            Exibe1.Caption = ""
                            
                            
                            'cancelar o Cupom de Venda
                            lErro = AFRAC_CancelarCupom(Me, gobjVenda)
                            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Cancelar Cupom")
                            If lErro <> SUCESSO Then gError 99610
                            
                            'Fecha a Transação
                            lErro = CF_ECF("Grava_CancelamentoCupom_Arquivo", gobjVenda)
                            If lErro <> SUCESSO Then gError 112421
                            
                            
                            lErro = Alteracoes_CancelamentoCupom(gobjVenda)
                            If lErro <> SUCESSO Then gError 112078
                            
                            gcolVendas.Remove (iIndice)
                        Else
                            gError 112075
                        End If
                        Exit For
                    Else
                        If iIndice = 1 Then gError 112075
                    End If
                Next
            'se vai ser cancelado o cupom que esta aberto
            Else
                'Envia aviso perguntando se deseja cancelar o cupom
                vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_CANCELA_CUPOM_TELA)
        
                If vbMsgRes = vbNo Then Exit Sub
                
                'Se for Necessário a autorização do Gerente para abertura do Caixa
                If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then
            
                    'Chama a Tela de Senha
                    Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)
            
                    'Sai de Função se a Tela de Login não Retornar ok
                    If giRetornoTela <> vbOK Then gError 102502
                    
                    'Se Operador for Gerente
                    iCodGerente = objOperador.iCodigo
            
                End If
                
                Exibe.Caption = "CANCELADO CUPOM " & gobjVenda.objCupomFiscal.lNumero
                Exibe1.Caption = ""
                
                'cancelar o Cupom de Venda
                lErro = AFRAC_CancelarCupom(Me, gobjVenda)
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Cancelar Cupom")
                If lErro <> SUCESSO Then gError 99610
                
                Call Move_Dados_Memoria_1
                
                'Realiza as operações necessárias para gravar
                lErro = CF_ECF("Grava_Venda_Arquivo", gobjVenda)
                If lErro <> SUCESSO Then gError 204536
                
                lErro = CF_ECF("Grava_CancelamentoCupom_Arquivo", gobjVenda)
                If lErro <> SUCESSO Then gError 204537
                
                
                Set gobjVenda = Nothing
                Set gobjVenda = New ClassVenda
                
                gobjVenda.iTipo = OPTION_CF
                
                If gobjLojaECF.iAbreAposFechamento = MARCADO Then
                    sCPF = gobjVenda.objCupomFiscal.sCPFCGC1
                    lErro = CF_ECF("Abre_Cupom", gobjVenda)
                    If lErro <> SUCESSO Then gError 99818
                End If
                
            End If
            
        End If
            
        Call Limpa_Tela_Venda
        
    End If
        
    Exit Sub

Erro_BotaoCancelaCupom_Click:

    Select Case gErr
                
        Case 99610, 112078, 112061, 99818, 102501, 102502, 105789, 105867, 199581, 204529, 204536, 204537, 204713, 210000, 210461, 210468
                            
        Case 105888
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_NAO_PREENCHIDO, gErr)
        
        Case 105790
            Call Rotina_ErroECF(vbOKOnly, AVISO_ORCAMENTO_INEXISTENTE, gErr)
        
        Case 112075
            Call Rotina_ErroECF(vbOKOnly, ERRO_CUPOM_NAO_CANCELADO, gErr)
                    
        Case 210453
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_BAIXADO, gErr, objVenda.objCupomFiscal.lNumOrcamento)

        Case 210991
            Call Rotina_ErroECF(vbOKOnly, ERRO_DAV_NAO_PODE_SER_REIMPRESSO, gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175677)

    End Select
    
    Exit Sub
        
End Sub

Private Function Alteracoes_CancelamentoCupom(objVenda As ClassVenda) As Long

Dim objMovCaixa As ClassMovimentoCaixa
Dim objCheque As ClassChequePre
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
Dim iIndice As Integer
Dim objCarne As ClassCarne
Dim lErro As Long
Dim objAdmMeioPagto As New ClassAdmMeioPagto
Dim lSequencial As Long
Dim objAliquota As New ClassAliquotaICMS
Dim objItens As ClassItemCupomFiscal
Dim iIndice1 As Integer
Dim sLog As String
Dim colRegistro As New Collection

On Error GoTo Erro_Alteracoes_CancelamentoCupom
    
    For iIndice = gobjVenda.colMovimentosCaixa.Count To 1 Step -1
        'Pega o Movimento
        Set objMovCaixa = gobjVenda.colMovimentosCaixa.Item(iIndice)
        'se for um recebimento em cartão de crédito/Debito de TEF
        If (objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAOCREDITO Or objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAODEBITO) And objMovCaixa.iTipoCartao = TIPO_TEF Then
            '''?????efetua caneclamento de TEF
        End If
    Next
    
    For Each objItens In objVenda.objCupomFiscal.colItens
        For Each objAliquota In gcolAliquotasTotal
            If objItens.dAliquotaICMS = objAliquota.dAliquota Then
                objAliquota.dValorTotalizadoLoja = objAliquota.dValorTotalizadoLoja - ((objItens.dPrecoUnitario * objItens.dQuantidade) * objAliquota.dAliquota)
                Exit For
            End If
        Next
    Next
    
    For iIndice = gcolMovimentosCaixa.Count To 1 Step -1
        Set objMovCaixa = gcolMovimentosCaixa.Item(iIndice)
        If objMovCaixa.lCupomFiscal = objVenda.objCupomFiscal.lNumero Then gcolMovimentosCaixa.Remove (iIndice)
    Next
    
    'Para cada movimento da venda
    For Each objMovCaixa In objVenda.colMovimentosCaixa
        If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_DINHEIRO Then gdSaldoDinheiro = gdSaldoDinheiro - objMovCaixa.dValor
        'Se for de cartao de crédito ou débito especificado
        If (objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAOCREDITO Or objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAODEBITO) And objMovCaixa.iAdmMeioPagto <> 0 Then
            'Busca em gcolCartão a ocorrencia de Cartão nao especificado
            For iIndice = gcolCartao.Count To 1 Step -1
                Set objAdmMeioPagtoCondPagto = gcolCartao.Item(iIndice)
                'Se encontrou
                If objAdmMeioPagtoCondPagto.iAdmMeioPagto = objMovCaixa.iAdmMeioPagto And objAdmMeioPagtoCondPagto.iParcelamento = objMovCaixa.iParcelamento And objAdmMeioPagtoCondPagto.iTipoCartao = objMovCaixa.iTipoCartao Then
                    'Atualiza o saldo do cartão
                    objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                    If objAdmMeioPagtoCondPagto.dSaldo = 0 Then gcolCartao.Remove (iIndice)
                    Exit For
                End If
            Next
        End If
        'Se o omvimento for de cartão de crédito não especificado
        If (objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAOCREDITO) And objMovCaixa.iAdmMeioPagto = 0 Then
            'inclui na col como não especificado
            For Each objAdmMeioPagtoCondPagto In gcolCartao
                If objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = STRING_NAO_DETALHADO_CDEBITO Then
                    'Atualiza o saldo de não especificado
                    objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                    Exit For
                End If
            Next
        End If
        'Se o omvimento for de cartão de débito não especificado
        If (objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAODEBITO) And objMovCaixa.iAdmMeioPagto = 0 Then
            'inclui na col como não especificado
            For Each objAdmMeioPagtoCondPagto In gcolCartao
                If objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = STRING_NAO_DETALHADO_CCREDITO Then
                    'Atualiza o saldo de não especificado
                    objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                    Exit For
                End If
            Next
        End If
    Next
    
    'Para cada movimento
    For iIndice = objVenda.colMovimentosCaixa.Count To 1 Step -1
        'Pega o movimento
        Set objMovCaixa = objVenda.colMovimentosCaixa.Item(iIndice)
        'Se for um recebimento em ticket
        If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_VALETICKET Then
            'Se for não especificado
            If objMovCaixa.iAdmMeioPagto = 0 Then
                'Para cada obj de ticket da coleção global de tickets
                For Each objAdmMeioPagtoCondPagto In gcolTicket
                    'Se for o não especificado
                    If objAdmMeioPagtoCondPagto.iAdmMeioPagto = 0 Then
                        'Atualiza o saldo de não especificado
                        objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                    End If
                Next
            'Se for especificado
            Else
                'Para cada Ticket da coleção global
                For iIndice1 = gcolTicket.Count To 1 Step -1
                    Set objAdmMeioPagtoCondPagto = gcolTicket.Item(iIndice1)
                    'Se encontrou o ticket/parcelamento
                    If objAdmMeioPagtoCondPagto.iAdmMeioPagto = objMovCaixa.iAdmMeioPagto Then
                        'Atualiza o saldo
                        objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                        If objAdmMeioPagtoCondPagto.dSaldo = 0 Then gcolTicket.Remove (iIndice1)
                        'Sinaliza que encontrou
                        Exit For
                    End If
                Next
            End If
        End If
    Next
    
    Set objAdmMeioPagtoCondPagto = New ClassAdmMeioPagtoCondPagto
    
    'Verifica se já existe movimentos de Outros\
    'Para cada MOvimento de Outros
    For iIndice = objVenda.colMovimentosCaixa.Count To 1 Step -1
        'Pega o MOvimento
        Set objMovCaixa = objVenda.colMovimentosCaixa.Item(iIndice)
        'Se for do tipo outros
        If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_OUTROS Then
            'Se for não especificado
            If objMovCaixa.iAdmMeioPagto = 0 Then
                'Para cada pagamento em outros na coleção global
                For Each objAdmMeioPagtoCondPagto In gcolOutros
                    'Se for o não especificado
                    If objAdmMeioPagtoCondPagto.iAdmMeioPagto = 0 Then
                        'Atualiza o saldo não especificado
                        objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                    End If
                Next
            'Se for especificado
            Else
                'Para cada Pagamento em outros na col global
                For iIndice1 = gcolOutros.Count To 1 Step -1
                    Set objAdmMeioPagtoCondPagto = gcolOutros.Item(iIndice1)
                    'Se for do mesmo tipo que o atual
                    If objAdmMeioPagtoCondPagto.iAdmMeioPagto = objMovCaixa.iAdmMeioPagto Then
                        'Atualiza o saldo
                        objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                        If objAdmMeioPagtoCondPagto.dSaldo = 0 Then gcolOutros.Remove (iIndice1)
                        Exit For
                    End If
                Next
            End If
        End If
    Next
        
    'remove o Carne na col global
    If objVenda.objCarne.colParcelas.Count > 0 Then
        For iIndice = 1 To gcolCarne.Count
            Set objCarne = gcolCarne.Item(iIndice)
            If objCarne.lCupomFiscal = objVenda.objCupomFiscal.lNumero Then gcolCarne.Remove (iIndice)
        Next
    End If
    
    'remove o Cheque na col global
    If objVenda.colCheques.Count > 0 Then
        For iIndice = gcolCheque.Count To 1 Step -1
            Set objCheque = gcolCheque.Item(iIndice)
            If objCheque.lCupomFiscal = objVenda.objCupomFiscal.lNumero Then gcolCheque.Remove (iIndice)
        Next
    End If
    
    Alteracoes_CancelamentoCupom = SUCESSO
    
    Exit Function
    
Erro_Alteracoes_CancelamentoCupom:
    
    Alteracoes_CancelamentoCupom = gErr
    
    Select Case gErr
    
        Case 99901, 99953, 99952
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175678)

    End Select
        
    lSequencial = glSeqTransacaoAberta
            
    'Rollback na transação
    Call CF_ECF("Caixa_Transacao_Rollback", lSequencial)
        
    Exit Function
    
End Function

Private Sub BotaoCancelaItemAtual_Click()
    
Dim iItem As Integer
Dim lErro As Long
Dim iIndice As Integer
Dim lNum As Long
Dim objItens As New ClassItemCupomFiscal
Dim objVenda As ClassVenda
Dim objVendaParam As New ClassVenda
Dim iAchou As Integer

On Error GoTo Erro_BotaoCancelaItemAtual_Click

    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 210002

    'Retorna o count de gobjvenda.objcupomfiscal.colitens
    lNum = Retorna_Count_ItensCupom
    
    If lNum > 0 Then
        
        Set objVenda = gobjVenda
        
        If (objVenda.iTipo = OPTION_DAV Or objVenda.iTipo = OPTION_PREVENDA) And objVenda.objCupomFiscal.lNumOrcamento = 0 Then gError 210498
    
        objVendaParam.objCupomFiscal.lNumOrcamento = gobjVenda.objCupomFiscal.lNumOrcamento
        
        'le o orcamento em questao
        lErro = CF_ECF("OrcamentoECF_Le", objVendaParam)
        If lErro <> SUCESSO And lErro <> 204690 And lErro <> 210447 Then gError 210499
        
        'se o orcamento é um DAV e ja foi impresso ==> nao pode alterar
        If objVendaParam.iTipo = OPTION_DAV And objVendaParam.objCupomFiscal.iDAVImpresso <> 0 Then gError 210500
        
        iAchou = 0
        
        For iIndice = gobjVenda.objCupomFiscal.colItens.Count To 1 Step -1
            Set objItens = gobjVenda.objCupomFiscal.colItens.Item(iIndice)
            If objItens.iStatus = STATUS_ATIVO Then
                iAchou = 1
                Exit For
            End If
        Next
        
        If iAchou = 1 Then
        
            If gobjVenda.iTipo = OPTION_CF Then
                'cancelar o Item anterior
                lErro = AFRAC_CancelarItem(CInt(objItens.iItem))
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Cancelar Item")
                If lErro <> SUCESSO Then gError 99611
                
    '            Call CF_ECF("Requisito_XXII_AtualizaGT")
                
            End If
            
            'Recolhe o Código que deve ser excluído da col de itens
    '        iItem = ListCF.ItemData(ListCF.ListCount - 1)
    '
    '        If iItem = 0 Then Exit Sub
            
            Exibe.Caption = "CANCELADO ITEM " & objItens.iItem
            Exibe1.Caption = ""
            Call Exclui_Item_ColItens(objItens.iItem)
        
        Else
            gError 210518
        
        End If
        
    Else
        'Senão erro-->deve existir um item
        gError 99926
    End If
    
    
    
    Exit Sub

Erro_BotaoCancelaItemAtual_Click:

    Select Case gErr
                
        Case 99611, 210002, 210499
                            
        Case 99926
            Call Rotina_ErroECF(vbOKOnly, ERRO_NAO_EXISTE_ITEM, gErr, Error$)
            
        Case 210498
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_NAO_PREENCHIDO, gErr)
        
        Case 210500
            Call Rotina_ErroECF(vbOKOnly, ERRO_DAV_NAO_ALTERADO_DEPOIS_DE_IMPRESSO, gErr)
            
        Case 210518
            Call Rotina_ErroECF(vbOKOnly, ERRO_ITEM_NAO_ENCONTRADO_CANCELAR, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175679)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoCancelaItem_Click()
    
Dim lErro As Long
Dim iIndice As Integer
Dim lNum As Long
Dim objVenda As New ClassVenda
Dim objItem As New ClassItemCupomFiscal
Dim colItem As New Collection
Dim objVendaParam As New ClassVenda

On Error GoTo Erro_BotaoCancelaItem_Click
    
    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 210001
    
    'Retorna o count de gobjvenda.objcupomfiscal.colitens
    lNum = Retorna_Count_ItensCupom
    
    If lNum > 0 Then
        Set objVenda = gobjVenda
        
        If (objVenda.iTipo = OPTION_DAV Or objVenda.iTipo = OPTION_PREVENDA) And objVenda.objCupomFiscal.lNumOrcamento = 0 Then gError 210492
    
        objVendaParam.objCupomFiscal.lNumOrcamento = gobjVenda.objCupomFiscal.lNumOrcamento
        
        'le o orcamento em questao
        lErro = CF_ECF("OrcamentoECF_Le", objVendaParam)
        If lErro <> SUCESSO And lErro <> 204690 And lErro <> 210447 Then gError 210493
        
        'se o orcamento é um DAV e ja foi impresso ==> nao pode alterar
        If objVendaParam.iTipo = OPTION_DAV And objVendaParam.objCupomFiscal.iDAVImpresso <> 0 Then gError 210494
        
        If Grid.Width < 8000 Then
        
            'Se tiver algum item selecionado e ele não for do cabeçalho-->continua
'            If ListCF.ListIndex > 13 Then
                objVenda.objCupomFiscal.iItem = ListCF.ItemData(ListCF.ListIndex)
'            Else
'                objVenda.objCupomFiscal.iItem = 0
'            End If
        
        Else
            If Grid.RowSel > 2 And Grid.RowSel <= giUltimaLinhaGrid Then
                objVenda.objCupomFiscal.iItem = StrParaInt(Grid.TextMatrix(Grid.RowSel, GRID_COL_ITEM))
            Else
                objVenda.objCupomFiscal.iItem = 0
            End If
            
        End If
            
        
        
        Call Chama_TelaECF_Modal("CancelaItem", objVenda)
        
        If giRetornoTela = vbOK Then
            If gobjVenda.iTipo = OPTION_CF Then
                'cancelar o item de Venda
                lErro = AFRAC_CancelarItem(objVenda.objCupomFiscal.iItem)
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Cancelar Item")
                If lErro <> SUCESSO Then gError 99612
                
'                Call CF_ECF("Requisito_XXII_AtualizaGT")
                
            End If
                
            Call Exclui_Item_ColItens(objVenda.objCupomFiscal.iItem)
            Exibe.Caption = "CANCELADO ITEM " & objVenda.objCupomFiscal.iItem
            
        End If
    Else
        'Senão erro-->deve existir um item
        gError 99923
    End If
    
    Exit Sub

Erro_BotaoCancelaItem_Click:

    Select Case gErr
                
        Case 99612, 210001, 210493
                            
        Case 99883
            Call Rotina_ErroECF(vbOKOnly, ITEM_CUPOM_NAO_SELECIONADO, gErr, Error$)
            
        Case 99923
            Call Rotina_ErroECF(vbOKOnly, ERRO_NAO_EXISTE_ITEM, gErr, Error$)
        
        Case 210492
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_NAO_PREENCHIDO, gErr)
        
        Case 210494
            Call Rotina_ErroECF(vbOKOnly, ERRO_DAV_NAO_ALTERADO_DEPOIS_DE_IMPRESSO, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175680)

    End Select
    
    Exit Sub
    
End Sub

Private Sub Exclui_Item_ColItens(ByVal iItem As Integer)

Dim objItens As ClassItemCupomFiscal
Dim iIndice As Integer
Dim iLinha As Integer

    'Percorre toda a lista
'    For iIndice = (ListCF.ListCount - 1) To 13 Step -1
'        'Se tiver o itemdata do código passado
'        If ListCF.ItemData(iIndice) = iItem Then
'            'Exclui este item
'            ListCF.RemoveItem (iIndice)
'        End If
'    Next
        
'    ListCF.AddItem Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "           ***** ITEM " & iItem & " CANCELADO *****")
        
        
    For Each objItens In gobjVenda.objCupomFiscal.colItens
        If objItens.iItem = iItem And objItens.iStatus = STATUS_ATIVO Then
'            objItens.icancel = ITEM_CANCELADO
            objItens.iStatus = STATUS_CANCELADO
            'Atualiza o subtotal
            Subtotal.Caption = Format(Subtotal.Caption - ((objItens.dPrecoUnitario * objItens.dQuantidade) - objItens.dValorDesconto), "standard")
                
            If objItens.dValorDesconto = 0 Then
                Exibe1.Caption = Formata_Campo(ALINHAMENTO_ESQUERDA, 8, " ", Format(objItens.dQuantidade, "0.000")) & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 4, " ", "x") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(objItens.dPrecoUnitario, "standard")) & Formata_Campo(ALINHAMENTO_ESQUERDA, 20, " ", Format(objItens.dPrecoUnitario * objItens.dQuantidade, "standard"))
            Else
                Exibe1.Caption = Formata_Campo(ALINHAMENTO_ESQUERDA, 8, " ", Format(objItens.dQuantidade, "0.000")) & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 4, " ", "x") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(objItens.dPrecoUnitario, "standard")) & "-" & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(objItens.dValorDesconto, "standard")) & Formata_Campo(ALINHAMENTO_ESQUERDA, 10, " ", Format(objItens.dPrecoUnitario - objItens.dValorDesconto, "standard"))
            End If
                        
            If Grid.Width < 8000 Then
                ListCF.AddItem Formata_Campo(ALINHAMENTO_DIREITA, 40, " ", "           ***** ITEM " & iItem & " CANCELADO *****") & Formata_Campo(ALINHAMENTO_ESQUERDA, 14, " ", "-" & Format(objItens.dPrecoUnitario * objItens.dQuantidade, "standard"))
                ListCF.ItemData(ListCF.NewIndex) = iItem
                
                'Para rolar automaticamente a barra de rolagem
                ListCF.ListIndex = ListCF.NewIndex
                
            Else
                Call Proxima_Linha_Grid
            
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_DESCRICAO) = "***** ITEM CANCELADO *****"
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_VALOR_TOTAL) = "-" & Format(objItens.dPrecoUnitario * objItens.dQuantidade, "standard")
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ITEM) = Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, iItem)
                
            End If
                                    
            Exit For
            
        End If
    Next
    
    
    
End Sub
        
Private Sub CodVendedor_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub


Private Sub CodVendedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As ClassVendedor
Dim bAchou As Boolean

On Error GoTo Erro_Codvendedor_Validate

    'Verifica se o vendedor foi preenchidO
    If Len(Trim(CodVendedor.Text)) = 0 Then
        'joga o vendedor no gobjvenda
        gobjVenda.objCupomFiscal.iVendedor = StrParaLong(CodVendedor.Text)
        LabelVendedor.Caption = ""
        Exit Sub
    End If
    
    bAchou = False
    
    For Each objVendedor In gcolVendedores
        'verifica se existe o vendedor na col
        If objVendedor.iCodigo = StrParaInt(CodVendedor.Text) Then
            LabelVendedor.Caption = objVendedor.sNomeReduzido
            bAchou = True
            Exit For
        End If
    Next
            
    'Não encontrou o vendedor
    If bAchou = False Then gError 99604
    
    'joga o vendedor no gobjvenda
    gobjVenda.objCupomFiscal.iVendedor = StrParaInt(CodVendedor.Text)
    
    'Função do AFRAC que informa vendedor
    lErro = AFRAC_InformarVendedor(CodVendedor.Text)
    If lErro <> SUCESSO Then
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Informar Vendedor")
        If lErro <> SUCESSO Then gError 99921
    End If

    Exit Sub

Erro_Codvendedor_Validate:

    Cancel = True

    Select Case gErr
    
        Case 99604
            Call Rotina_ErroECF(vbOKOnly, ERRO_VENDEDOR_NAO_CADASTRADO2, gErr)
            
        Case 99921
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175681)

    End Select
    
    Exit Sub

End Sub



Private Sub Orcamento_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Orcamento_GotFocus()
    Call MaskEdBox_TrataGotFocus(Orcamento, iAlterado)
End Sub

Private Sub Orcamento_Validate(Cancel As Boolean)

Dim objVenda As ClassVenda
Dim colOrcamento As New Collection
Dim objItens As ClassItemCupomFiscal
Dim iIndice As Integer
Dim lErro As Long
Dim objProduto As ClassProduto
Dim objVenda1 As New ClassVenda

On Error GoTo Erro_Orcamento_Validate

    If Len(Trim(Orcamento.Text)) > 0 Then
        lErro = CF_ECF("Requisito_XXII")
        If lErro <> SUCESSO Then gError 210008
    End If

    'Se existe um número
    If Len(Trim(Orcamento.Text)) > 0 And gobjVenda.objCupomFiscal.lNumOrcamento <> StrParaLong(Orcamento.Text) Then
    
        
        objVenda1.objCupomFiscal.lNumOrcamento = StrParaLong(Orcamento.Text)
        gobjVenda.objCupomFiscal.lNumOrcamento = StrParaLong(Orcamento.Text)
            
        
        
        'Função Que le os orcamentos
        lErro = CF_ECF("OrcamentoECF_Le", objVenda1)
        If lErro <> SUCESSO And lErro <> 204690 And lErro <> 210447 Then gError 105913
        
        If lErro = 204690 And OptionCF.Value = True Then gError 210456
        
        If lErro = 210447 Then gError 210454
        
        If lErro = SUCESSO Then
            
            'descobre o nome reduzido do produto
            For Each objItens In objVenda1.objCupomFiscal.colItens
                For iIndice = 1 To gaobjProdutosNome.Count
                    Set objProduto = gaobjProdutosNome.Item(iIndice)
                    If objItens.sProduto = objProduto.sCodigo Then
                        objItens.sProdutoNomeRed = objProduto.sNomeReduzido
                        Exit For
                    End If
                Next
            Next
            
            Call Copia_Venda(gobjVenda, objVenda1)
            Call Traz_Orcamento
            
            If OptionCF.Value Then Call OptionCF_Click
            
        End If
        
    End If
    
    Exit Sub
    
Erro_Orcamento_Validate:
    
    Cancel = True
    
    Select Case gErr
                
        Case 105913, 210008
                
        Case 210454
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_BAIXADO, gErr, objVenda1.objCupomFiscal.lNumOrcamento)
                
        Case 210456
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_NAO_CADASTRADO1, gErr, objVenda1.objCupomFiscal.lNumOrcamento)
                
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175682)

    End Select
    
    Exit Sub
    
End Sub

Private Sub Traz_Orcamento()
    
Dim objItens As New ClassItemCupomFiscal
Dim objProduto As ClassProduto
Dim sProduto As String
Dim lErro As Long
Dim iIndice As Integer
Dim sProduto1 As String

On Error GoTo Erro_Traz_Orcamento
    
    Exibe.Caption = "CONSULTA DE ORÇAMENTO."
    Exibe1.Caption = ""
    DoEvents
    
    ListCF.Clear
    Subtotal.Caption = ""
    
    Call Inicia_Cupom_Tela
        
    Orcamento.Text = CStr(gobjVenda.objCupomFiscal.lNumOrcamento)

    If gobjVenda.objCupomFiscal.iVendedor > 0 Then
        CodVendedor.Text = gobjVenda.objCupomFiscal.iVendedor
    Else
        CodVendedor.Text = ""
    End If
    Call CodVendedor_Validate(False)
    
    NomeCliente.Text = gobjVenda.objCupomFiscal.sNomeCliente
    CodCliente.Text = gobjVenda.objCupomFiscal.sCPFCGC
    CGC.Text = gobjVenda.objCupomFiscal.sCPFCGC1
    'Endereco.Text = gobjVenda.objCupomFiscal.sEndereco
    
    Call CodCliente_Validate(False)
    
    'Para cada Item --> inclui no Cupom
    For Each objItens In gobjVenda.objCupomFiscal.colItens
                       
        ProdutoNomeRed.Text = objItens.sProdutoNomeRed
        Quantidade.Text = objItens.dQuantidade
        sProduto1 = objItens.sProdutoNomeRed
                
        Call TP_Produto_Le_Col(gaobjProdutosReferencia, gaobjProdutosCodBarras, gaobjProdutosNome, sProduto1, objProduto)
        
        'caso o produto não seja encontrado
        If objProduto Is Nothing Then gError 99884
        
        ProdutoNomeRed.Text = objProduto.sNomeReduzido
        PrecoUnitario.Caption = Format(objProduto.dPrecoLoja, "standard")
        
        'Prenche a col de itens do cupom com os dados do mesmo
        PrecoItem.Caption = Format(StrParaDbl(Quantidade.Text) * StrParaDbl(PrecoUnitario.Caption), "Standard")
        
        If objItens.iStatus = STATUS_ATIVO Then
            Subtotal.Caption = Format(StrParaDbl(Subtotal.Caption) + (StrParaDbl(PrecoItem.Caption) - objItens.dValorDesconto), "standard")
        End If
            
        If Len(Trim(objProduto.sReferencia)) > 0 Then
            sProduto = objProduto.sReferencia
        Else
            sProduto = objProduto.sCodigoBarras
        End If
        
        If Grid.Width < 8000 Then
            ListCF.AddItem Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, objItens.iItem) & "     " & Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", objProduto.sCodigo) & Formata_Campo(ALINHAMENTO_DIREITA, 30, " ", objProduto.sDescricao)
'            ListCF.AddItem Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, objItens.iItem) & "     " & Formata_Campo(ALINHAMENTO_DIREITA, 50, " ", objProduto.sDescricao)
            ListCF.ItemData(ListCF.NewIndex) = objItens.iItem
            ListCF.AddItem Formata_Campo(ALINHAMENTO_ESQUERDA, 11, " ", Format(Quantidade.Text, "#0.000")) & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 4, " ", "x") & Formata_Campo(ALINHAMENTO_DIREITA, 12, " ", Format(PrecoUnitario.Caption, "standard")) & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(objItens.dAliquotaICMS * 100, "fixed") & "%") & Formata_Campo(ALINHAMENTO_ESQUERDA, 14, " ", Format(PrecoItem.Caption, "standard"))
            ListCF.ItemData(ListCF.NewIndex) = objItens.iItem
        Else
            Call Proxima_Linha_Grid
    
            Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ITEM) = Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, objItens.iItem)
            Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_CODIGO) = objProduto.sCodigo
            Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_DESCRICAO) = objProduto.sDescricao
            Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_QUANTIDADE) = Format(Quantidade.Text, "#0.000") & objProduto.sSiglaUMVenda
            Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_X) = "x"
            Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_VALOR_UNITARIO) = Format(PrecoUnitario.Caption, "standard")
            Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ST) = objProduto.sSituacaoTribECF & Format(objItens.dAliquotaICMS * 100, "fixed") & "%"
            Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_VALOR_TOTAL) = Format(PrecoItem.Caption, "standard")
            
        End If
        
        'se existir desconto sobre o item...
        If objItens.dValorDesconto > 0 Then
            If Grid.Width < 8000 Then
                ListCF.AddItem Formata_Campo(ALINHAMENTO_DIREITA, 9, " ", "DESCONTO:") & Formata_Campo(ALINHAMENTO_ESQUERDA, 10, " ", Format(objProduto.dPercentMenosReceb, "fixed") & "%") & Formata_Campo(ALINHAMENTO_ESQUERDA, 11, " ", "-" & Format(StrParaDbl(PrecoItem.Caption) * (objProduto.dPercentMenosReceb / 100), "standard")) & Formata_Campo(ALINHAMENTO_ESQUERDA, 20, " ", Format(StrParaDbl(PrecoItem.Caption) - objItens.dValorDesconto, "standard"))
                ListCF.ItemData(ListCF.NewIndex) = objItens.iItem
            Else
                
                Call Proxima_Linha_Grid
                
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_DESCRICAO) = "DESCONTO:"
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_VALOR_UNITARIO) = Format(objProduto.dPercentMenosReceb, "fixed") & "%"
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ST) = "-" & Format(StrParaDbl(PrecoItem.Caption) * (objProduto.dPercentMenosReceb / 100), "standard")
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_VALOR_TOTAL) = Format(StrParaDbl(PrecoItem.Caption) - objItens.dValorDesconto, "standard")
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ITEM) = Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, objItens.iItem)
            End If
        End If
        
        If objItens.iStatus = STATUS_CANCELADO Then
            If Grid.Width < 8000 Then
'                ListCF.AddItem Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "           ***** ITEM " & iItem & " CANCELADO *****")
                ListCF.AddItem Formata_Campo(ALINHAMENTO_DIREITA, 40, " ", "           ***** ITEM " & objItens.iItem & " CANCELADO *****") & Formata_Campo(ALINHAMENTO_ESQUERDA, 14, " ", "-" & Format(objItens.dPrecoUnitario * objItens.dQuantidade, "standard"))
                ListCF.ItemData(ListCF.NewIndex) = objItens.iItem
            Else
                Call Proxima_Linha_Grid
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_DESCRICAO) = "***** ITEM CANCELADO *****"
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_VALOR_TOTAL) = "-" & Format(objItens.dPrecoUnitario * objItens.dQuantidade, "standard")
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ITEM) = Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, objItens.iItem)
                
            End If
        End If
        
        
        'Para rolar automaticamente a barra de rolagem
        If Grid.Width < 8000 Then
            ListCF.ListIndex = ListCF.NewIndex
        End If
        
        Call Limpa_Cupom_Tela
    Next
        
'    If gobjVenda.iTipo = OPTION_PREVENDA Then
'        OptionPreVenda.Value = True
'    ElseIf gobjVenda.iTipo = OPTION_DAV Then
'        OptionDAV.Value = True
'    End If
        
    Exit Sub
    
Erro_Traz_Orcamento:

    Select Case gErr
                
        Case 99884
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175683)

    End Select
    
    Exit Sub
    
End Sub

Private Sub ProdutoNomeRed_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProdutoNomeRed_GotFocus()
    ProdutoNomeRed.SelStart = 0
    ProdutoNomeRed.SelLength = Len(ProdutoNomeRed.Text)
End Sub

Private Sub ProdutoNomeRed_Validate(Cancel As Boolean)

Dim lErro As Long
Dim bAchou As Boolean
Dim objProduto As ClassProduto
Dim objItens As New ClassItemCupomFiscal
Dim sProduto As String

On Error GoTo Erro_ProdutoNomeRed_Validate
    
    If Len(Trim(ProdutoNomeRed.Text)) > 0 Then
    
        lErro = CF_ECF("Requisito_XXII")
        If lErro <> SUCESSO Then gError 210012

    End If

    Parent.MousePointer = vbHourglass
    
    'Se o produto não está preenchido
    If Len(Trim(ProdutoNomeRed.Text)) = 0 Then
        PrecoUnitario.Caption = Format(0, "standard")
        PrecoItem.Caption = Format(0, "standard")
    'caso contrário
    Else
        'Verifica a quantidade
        If Len(Trim(Quantidade.Text)) <> 0 Then
            'Rotina de cupom
            
            lErro = Adiciona_Cupom(0)
            If lErro <> SUCESSO Then gError 99885

        End If
    End If
    
    Parent.MousePointer = vbDefault
    
    Exit Sub

Erro_ProdutoNomeRed_Validate:

    Cancel = True

    Parent.MousePointer = vbDefault

    Select Case gErr
                
        Case 99885, 210012
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175684)

    End Select
    
    Exit Sub

End Sub

Private Sub Quantidade_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Quantidade_GotFocus()

    Quantidade.SelStart = 0
    Quantidade.SelLength = Len(Quantidade.Text)

End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Quantidade_Validate
    
    'Se a quantidade e o produto estão prenchidos
    If Len(Trim(Quantidade.Text)) > 0 Then
        
        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 99886
        If Len(Trim(Fix(Quantidade.Text))) > 4 Then gError 112423
        Quantidade.Text = Format(Quantidade.Text, "0.000")
        If right(Quantidade.Text, 4) = ",000" Then Quantidade.Text = Format(Quantidade.Text, "#,#")
        
        If Len(Trim(ProdutoNomeRed.Text)) <> 0 Then
            'Rotina de cupom
            lErro = Adiciona_Cupom(0)
            If lErro <> SUCESSO Then gError 210883

        End If
        
        
    End If
    
    Exit Sub

Erro_Quantidade_Validate:

    Cancel = True

    Select Case gErr
            
        Case 99886, 210883
        
        Case 112423
            Call Rotina_ErroECF(vbOKOnly, ERRO_QUANTIDADE_INVALIDA, gErr)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175685)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoPagamento_Click()

Dim iCodigo As Integer
Dim lErro As Long
Dim objGenerico As New AdmGenerico
Dim objTela As Object

On Error GoTo Erro_BotaoPagamento_Click
    
    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 210004

    'Se não há valor para pagar --> erro.
    If StrParaDbl(Subtotal.Caption) = 0 Then gError 99717
    
    
    'sevé obrigatório o preenchimento do vendedor
    If gobjLojaECF.iVendedorObrigatorio = 1 Then
        If Len(Trim(CodVendedor.Text)) = 0 Then gError 112072
    End If
    
    gobjVenda.objCupomFiscal.dValorTotal = StrParaDbl(Subtotal.Caption)
    gobjVenda.objCupomFiscal.dValorProdutos = StrParaDbl(Subtotal.Caption)
    
    
    If gobjVenda.iTipo = OPTION_CF Then
'        Exibe.Caption = ""
'        Exibe1.Caption = ""
'        DoEvents
        
        Call Chama_TelaECF_Modal("Pagamento", gobjVenda, objGenerico)
        'Se não deu tudo certo
        If objGenerico.vVariavel = vbCancel Then
            Exibe.Caption = "CANCELADO CUPOM " & gobjVenda.objCupomFiscal.lNumero
            Exibe1.Caption = ""
            
            'cancelar o Cupom de Venda
            lErro = AFRAC_CancelarCupom(Me, gobjVenda)
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Cancela Cupom")
            If lErro <> SUCESSO Then gError 99614
            
            Call Move_Dados_Memoria_1
            
            'Realiza as operações necessárias para gravar
            lErro = CF_ECF("Grava_Venda_Arquivo", gobjVenda)
            If lErro <> SUCESSO Then gError 204530
            
            lErro = CF_ECF("Grava_CancelamentoCupom_Arquivo", gobjVenda)
            If lErro <> SUCESSO Then gError 204531
            
            
        End If
    Else
    
        Set objTela = Me
    
        lErro = CF_ECF("Valida_Orcamento", objTela)
        If lErro <> SUCESSO Then gError 204330
        
        Call Chama_TelaECF_Modal("FechaOrcamento", gobjVenda, objGenerico)
    End If
    
    If objGenerico.vVariavel <> vbAbort Then
        Set gobjVenda = Nothing
        Set gobjVenda = New ClassVenda
    
        Call Limpa_Tela_Venda
        
    End If
    
    ProdutoNomeRed.SetFocus
    
    Exit Sub

Erro_BotaoPagamento_Click:

    Select Case gErr
                
        Case 99614
            Set gobjVenda = Nothing
            Set gobjVenda = New ClassVenda
                
            Call Limpa_Tela_Venda
                
        Case 99717
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_NAO_EXISTENTE, gErr)
                    
        Case 99896
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_NAO_PREENCHIDO, gErr, Error$)
                                
        Case 112072
            Call Rotina_ErroECF(vbOKOnly, ERRO_VENDEDOR_NAO_PREENCHIDO, gErr, Error$)
        
        Case 204330, 204530, 204531, 210004
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175686)

    End Select
    
    Exit Sub
    
End Sub

Private Sub OptionCF_Click()
    
Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long
Dim lNum As Long
Dim lNumero As Long
Dim objVendaAux As New ClassVenda
Dim objCliente As ClassCliente
Dim sCPF As String
Dim iIndice As Integer
Dim colItens As New Collection
Dim colItens1 As New Collection
Dim objVendaOrc As New ClassVenda

On Error GoTo Erro_OptionCF_Click

    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 210005
       
    lNum = Retorna_Count_ItensCupom()
    
    'Se tiver um orçamento selecionado
    If lNum <> 0 And (gobjVenda.iTipo = OPTION_PREVENDA Or gobjVenda.iTipo = OPTION_DAV) Then
        
        
        'Envia aviso perguntando se deseja transforma o orçamento em venda
        vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_ORCAMENTO_VENDA)

        If vbMsgRes = vbNo Then
            Call Limpa_Tela_Venda
            Set gobjVenda = Nothing
            Set gobjVenda = New ClassVenda
            
        Else
            Exibe.Caption = "AGUARDE..."
            Exibe1.Caption = ""
            DoEvents
            
'            If giRemoveOrc = REMOVER_ORC Then

                'exclui o orcamento que está sendo transformado em cupom
'                lErro = CF_ECF("Caixa_Exclui_Orcamento", gobjVenda)
'                If lErro <> SUCESSO And lErro <> 105761 Then gError 105766

 '           End If
            
            'se o orcamento nao estava gravado ==> zera o numero do orcamento
'            If lErro <> SUCESSO Then gobjVenda.objCupomFiscal.lNumOrcamento = 0
            
            
            'se o numero do orcamento esta preenchido verifica se esta gravado. Se nao estiver gravado ==> limpa o numero do orcamento
            If gobjVenda.objCupomFiscal.lNumOrcamento <> 0 Then
                
                objVendaOrc.objCupomFiscal.lNumOrcamento = gobjVenda.objCupomFiscal.lNumOrcamento
            
                'Função Que le os orcamentos
                lErro = CF_ECF("OrcamentoECF_Le", objVendaOrc)
                If lErro <> SUCESSO And lErro <> 204690 And lErro <> 210447 Then gError 210511
            
                If lErro = 210447 Then gError 210512
                
                'se nao esta cadatrado o orcamento, limpa o numero
                If lErro = 204690 Then gobjVenda.objCupomFiscal.lNumOrcamento = 0
            
            End If
            
            
            Call Copia_Venda(objVendaAux, gobjVenda)
            
            'guardo o número do orçamento que agora está relacionado com o cupom
            Set gobjVenda = objVendaAux
            gobjVenda.objCupomFiscal.dtDataOrcamento = gobjVenda.objCupomFiscal.dtDataEmissao
            
            For iIndice = gobjVenda.objCupomFiscal.colItens.Count To 1 Step -1
                colItens1.Add gobjVenda.objCupomFiscal.colItens.Item(iIndice)
                gobjVenda.objCupomFiscal.colItens.Remove (iIndice)
            Next
            
            For iIndice = colItens1.Count To 1 Step -1
                colItens.Add colItens1.Item(iIndice)
            Next
            
            If Len(Trim(Orcamento.Text)) > 0 Then gobjVenda.objCupomFiscal.lNumOrcamento = StrParaLong(Orcamento.Text)
            
            gobjVenda.iTipo = OPTION_CF
'            gobjVenda.objCupomFiscal.dtDataEmissao = Date
'            gobjVenda.objCupomFiscal.dHoraEmissao = CDbl(Time)
            gobjVenda.objCupomFiscal.iCodCaixa = giCodCaixa
            gobjVenda.objCupomFiscal.iECF = giCodECF
'            gobjVenda.objCupomFiscal.dtDataReducao = gdtDataAnterior

            'Abre o cupom
'            sCPF = CodCliente.FormattedText
'            lErro = CF_ECF("Abre_Cupom", lNumero, sCPF)
'            If lErro <> SUCESSO Then gError 99818
'            gobjVenda.objCupomFiscal.lNumero = lNumero
    
            lErro = Transforma_Cupom(colItens)
            If lErro <> SUCESSO Then gError 109396
            Exibe.Caption = "TRANSF.: ORÇAMENTO EM VENDA."
            Exibe1.Caption = ""
            
        End If
        
    End If
    
    'Quando eu clico no cupom fiscal desativa o número do orcamento(campo)
    Orcamento.Text = ""
'    Orcamento.Enabled = False
'    LabelOrcamento.Enabled = False
    BotaoProxNum.Enabled = False
    gobjVenda.iTipo = OPTION_CF
    
    BotaoAbrirGaveta.Caption = "(F10)   Abrir Gaveta"
    BotaoCancelaCupom.Caption = "(Esc)   Cancela Cupom"
    
    Exit Sub

Erro_OptionCF_Click:

    Select Case gErr
                
        Case 99615, 99818, 99847, 105766, 109396, 210005, 210511
                            
        Case 105767
            Call Rotina_ErroECF(vbOKOnly, AVISO_ORCAMENTO_INEXISTENTE, gErr)
                            
        Case 210512
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_BAIXADO, gErr, gobjVenda.objCupomFiscal.lNumOrcamento)
                            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175687)

    End Select
    
    Exit Sub
        
End Sub

Private Sub Copia_Venda(objVendaAux As ClassVenda, objVenda As ClassVenda)
    
Dim iIndice As Integer
Dim objCheque As New ClassChequePre
Dim objChequeAux As New ClassChequePre
Dim objMVCX As New ClassMovimentoCaixa
Dim objMVCXAux As New ClassMovimentoCaixa
Dim objTroca As New ClassTroca
Dim objTrocaAux As New ClassTroca
Dim objCarne As New ClassCarne
Dim objCarneAux As New ClassCarne
Dim objCarneParc As New ClassCarneParcelas
Dim objCarneParcAux As New ClassCarneParcelas
Dim objCupom As New ClassCupomFiscal
Dim objCupomAux As New ClassCupomFiscal
Dim objCupomItem As New ClassItemCupomFiscal
Dim objCupomItemAux As New ClassItemCupomFiscal
    
    Set objVendaAux = New ClassVenda
    
    For iIndice = 1 To objVenda.colCheques.Count
        Set objCheque = objVenda.colCheques.Item(iIndice)
        Set objChequeAux = New ClassChequePre
        
        objChequeAux.dtDataDeposito = objCheque.dtDataDeposito
        objChequeAux.dValor = objCheque.dValor
        objChequeAux.iAprovado = objCheque.iAprovado
        objChequeAux.iBanco = objCheque.iBanco
        objChequeAux.iCaixa = objCheque.iCaixa
        objChequeAux.iChequeSel = objCheque.iChequeSel
        objChequeAux.iECF = objCheque.iECF
        objChequeAux.iFilial = objCheque.iFilial
        objChequeAux.iFilialEmpresa = objCheque.iFilialEmpresa
        objChequeAux.iFilialEmpresaLoja = objCheque.iFilialEmpresaLoja
        objChequeAux.iNaoEspecificado = objCheque.iNaoEspecificado
        objChequeAux.iStatus = objCheque.iStatus
        objChequeAux.iTipoBordero = objCheque.iTipoBordero
        objChequeAux.lCliente = objCheque.lCliente
        objChequeAux.lCupomFiscal = objCheque.lCupomFiscal
        objChequeAux.lNumBordero = objCheque.lNumBordero
        objChequeAux.lNumBorderoLoja = objCheque.lNumBorderoLoja
        objChequeAux.lNumero = objCheque.lNumero
        objChequeAux.lNumIntCheque = objCheque.lNumIntCheque
        objChequeAux.lNumIntChequeBord = objCheque.lNumIntChequeBord
        objChequeAux.lNumIntExt = objCheque.lNumIntExt
        objChequeAux.lNumMovtoCaixa = objCheque.lNumMovtoCaixa
        objChequeAux.lNumMovtoSangria = objCheque.lNumMovtoSangria
        objChequeAux.lSequencial = objCheque.lSequencial
        objChequeAux.lSequencialBack = objCheque.lSequencialBack
        objChequeAux.lSequencialCaixa = objCheque.lSequencialCaixa
        objChequeAux.lSequencialLoja = objCheque.lSequencialLoja
        objChequeAux.sAgencia = objCheque.sAgencia
        objChequeAux.sCarne = objCheque.sCarne
        objChequeAux.sContaCorrente = objCheque.sContaCorrente
        objChequeAux.sCPFCGC = objCheque.sCPFCGC
        
        objVendaAux.colCheques.Add objChequeAux
    Next
    
    For iIndice = 1 To objVenda.colMovimentosCaixa.Count
        Set objMVCX = objVenda.colMovimentosCaixa.Item(iIndice)
        Set objMVCXAux = New ClassMovimentoCaixa
        
        objMVCXAux.dHora = objMVCX.dHora
        objMVCXAux.dtDataMovimento = objMVCX.dtDataMovimento
        objMVCXAux.dValor = objMVCX.dValor
        objMVCXAux.iAdmMeioPagto = objMVCX.iAdmMeioPagto
        objMVCXAux.iCaixa = objMVCX.iCaixa
        objMVCXAux.iCodConta = objMVCX.iCodConta
        objMVCXAux.iCodOperador = objMVCX.iCodOperador
        objMVCXAux.iExcluiu = objMVCX.iExcluiu
        objMVCXAux.iFilialEmpresa = objMVCX.iFilialEmpresa
        objMVCXAux.iGerente = objMVCX.iGerente
        objMVCXAux.iParcelamento = objMVCX.iParcelamento
        objMVCXAux.iQuantLog = objMVCX.iQuantLog
        objMVCXAux.iTipo = objMVCX.iTipo
        objMVCXAux.iTipoCartao = objMVCX.iTipoCartao
        objMVCXAux.lCupomFiscal = objMVCX.lCupomFiscal
        objMVCXAux.lMovtoEstorno = objMVCX.lMovtoEstorno
        objMVCXAux.lMovtoTransf = objMVCX.lMovtoTransf
        objMVCXAux.lNumero = objMVCX.lNumero
        objMVCXAux.lNumIntDocLog = objMVCX.lNumIntDocLog
        objMVCXAux.lNumIntExt = objMVCX.lNumIntExt
        objMVCXAux.lNumMovto = objMVCX.lNumMovto
        objMVCXAux.lNumRefInterna = objMVCX.lNumRefInterna
        objMVCXAux.lSequencial = objMVCX.lSequencial
        objMVCXAux.lSequencialConta = objMVCX.lSequencialConta
        objMVCXAux.lTransferencia = objMVCX.lTransferencia
        objMVCXAux.sFavorecido = objMVCX.sFavorecido
        objMVCXAux.sHistorico = objMVCX.sHistorico
        
        objVendaAux.colMovimentosCaixa.Add objMVCXAux
    Next
    
    For iIndice = 1 To objVenda.colTroca.Count
        Set objTroca = objVenda.colTroca.Item(iIndice)
        Set objTrocaAux = New ClassTroca
        
        objTrocaAux.dQuantidade = objTroca.dQuantidade
        objTrocaAux.dValor = objTroca.dValor
        objTrocaAux.iFilialEmpresa = objTroca.iFilialEmpresa
        objTrocaAux.lNumIntDoc = objTroca.lNumIntDoc
        objTrocaAux.lNumMovtoCaixa = objTroca.lNumMovtoCaixa
        objTrocaAux.sCodProduto = objTroca.sCodProduto
        objTrocaAux.sProduto = objTroca.sProduto
        objTrocaAux.sUnidadeMed = objTroca.sUnidadeMed
        
        objVendaAux.colTroca.Add objTrocaAux
    Next
    
    objVendaAux.iTipo = objVenda.iTipo
    
    Set objCarne = objVenda.objCarne
    
    objCarneAux.dtDataReferencia = objCarne.dtDataReferencia
    objCarneAux.iFilialEmpresa = objCarne.iFilialEmpresa
    objCarneAux.iStatus = objCarne.iStatus
    objCarneAux.lCliente = objCarne.lCliente
    objCarneAux.lCupomFiscal = objCarne.lCupomFiscal
    objCarneAux.lNumIntDoc = objCarne.lNumIntDoc
    objCarneAux.lNumIntExt = objCarne.lNumIntExt
    objCarneAux.sAutorizacao = objCarne.sAutorizacao
    objCarneAux.sCodBarrasCarne = objCarne.sCodBarrasCarne
        
    Set objVendaAux.objCarne = objCarneAux
    
    For iIndice = 1 To objVenda.objCarne.colParcelas.Count
        Set objCarneParc = objVenda.objCarne.colParcelas.Item(iIndice)
        Set objCarneParcAux = New ClassCarneParcelas
        
        objCarneParcAux.dtDataVencimento = objCarneParc.dtDataVencimento
        objCarneParcAux.dValor = objCarneParc.dValor
        objCarneParcAux.iFilialEmpresa = objCarneParc.iFilialEmpresa
        objCarneParcAux.iParcela = objCarneParc.iParcela
        objCarneParcAux.iStatus = objCarneParc.iStatus
        objCarneParcAux.lNumIntCarne = objCarneParc.lNumIntCarne
        objCarneParcAux.lNumIntDoc = objCarneParc.lNumIntDoc
        
        objVendaAux.objCarne.colParcelas.Add objCarneParcAux
    Next
    
    Set objCupom = objVenda.objCupomFiscal
    
    objCupomAux.dHoraEmissao = objCupom.dHoraEmissao
    objCupomAux.dtDataEmissao = objCupom.dtDataEmissao
    objCupomAux.dValorAcrescimo = objCupom.dValorAcrescimo
    objCupomAux.dValorDesconto = objCupom.dValorDesconto
    objCupomAux.dValorProdutos = objCupom.dValorProdutos
    objCupomAux.dValorTotal = objCupom.dValorTotal
    objCupomAux.dValorTroco = objCupom.dValorTroco
    objCupomAux.iCodCaixa = objCupom.iCodCaixa
    objCupomAux.iECF = objCupom.iECF
    objCupomAux.iFilialEmpresa = objCupom.iFilialEmpresa
    objCupomAux.iStatus = objCupom.iStatus
    objCupomAux.iTabelaPreco = objCupom.iTabelaPreco
    objCupomAux.iTipo = objCupom.iTipo
    objCupomAux.iVendedor = objCupom.iVendedor
    objCupomAux.lCliente = objCupom.lCliente
    objCupomAux.lDuracao = objCupom.lDuracao
    objCupomAux.lGerenteCancel = objCupom.lGerenteCancel
    objCupomAux.lNumero = objCupom.lNumero
    objCupomAux.lNumIntDoc = objCupom.lNumIntDoc
    objCupomAux.lNumOrcamento = objCupom.lNumOrcamento
    objCupomAux.sCPFCGC = objCupom.sCPFCGC
    objCupomAux.sMotivoCancel = objCupom.sMotivoCancel
    objCupomAux.sNaturezaOp = objCupom.sNaturezaOp
    objCupomAux.sCPFCGC1 = objCupom.sCPFCGC1
    objCupomAux.sNomeCliente = objCupom.sNomeCliente
    objCupomAux.lNumeroDAV = objCupom.lNumeroDAV
    objCupomAux.sEndereco = objCupom.sEndereco
    objCupomAux.sEndEntLogradouro = objCupom.sEndEntLogradouro
    objCupomAux.sEndEntNúmero = objCupom.sEndEntNúmero
    objCupomAux.sEndEntComplemento = objCupom.sEndEntComplemento
    objCupomAux.sEndEntBairro = objCupom.sEndEntBairro
    objCupomAux.sEndEntCidade = objCupom.sEndEntCidade
    objCupomAux.sEndEntUF = objCupom.sEndEntUF
    objCupomAux.iDAVImpresso = objCupom.iDAVImpresso
    objCupomAux.lCOOCupomOrigDAV = objCupom.lCOOCupomOrigDAV
        
    Set objVendaAux.objCupomFiscal = objCupomAux
    
    For iIndice = 1 To objVenda.objCupomFiscal.colItens.Count
        Set objCupomItem = objVenda.objCupomFiscal.colItens.Item(iIndice)
        Set objCupomItemAux = New ClassItemCupomFiscal
        
        objCupomItemAux.dAliquotaICMS = objCupomItem.dAliquotaICMS
        objCupomItemAux.dPercDesc = objCupomItem.dPercDesc
        objCupomItemAux.dPrecoUnitario = objCupomItem.dPrecoUnitario
        objCupomItemAux.dQuantidade = objCupomItem.dQuantidade
        objCupomItemAux.dValorDesconto = objCupomItem.dValorDesconto
'        objCupomItemAux.icancel = objCupomItem.icancel
        objCupomItemAux.iCodCaixa = objCupomItem.iCodCaixa
        objCupomItemAux.iFilialEmpresa = objCupomItem.iFilialEmpresa
        objCupomItemAux.iItem = objCupomItem.iItem
        objCupomItemAux.iStatus = objCupomItem.iStatus
        objCupomItemAux.lNumIntCupom = objCupomItem.lNumIntCupom
        objCupomItemAux.lNumIntDoc = objCupomItem.lNumIntDoc
        objCupomItemAux.sProduto = objCupomItem.sProduto
        objCupomItemAux.sSituacaoTrib = objCupomItem.sSituacaoTrib
        objCupomItemAux.sUnidadeMed = objCupomItem.sUnidadeMed
        objCupomItemAux.sProdutoNomeRed = objCupomItem.sProdutoNomeRed
        
        objVendaAux.objCupomFiscal.colItens.Add objCupomItemAux
    Next
    
End Sub

Private Function Transforma_Cupom(colItens As Collection) As Long
    
Dim objItens As ClassItemCupomFiscal
Dim lErro As Long
Dim bAchou As Boolean
Dim objProduto As ClassProduto
Dim sProduto As String
Dim lNum As Long
Dim lNumero As Long

On Error GoTo Erro_Transforma_Cupom
        
'    Call Limpa_Tela_Venda_1
    Call Limpa_Tela_Venda
    
    For Each objItens In colItens
        
'        If objItens.iStatus = STATUS_ATIVO Or gobjVenda.objCupomFiscal.lNumeroDAV <> 0 Then
            
            ProdutoNomeRed.Text = objItens.sProdutoNomeRed
            Quantidade.Text = objItens.dQuantidade
            
            lErro = Adiciona_Cupom(1)
            If lErro <> SUCESSO Then gError 210510
            
'        End If
                
                
'        If objItens.iStatus = STATUS_CANCELADO And gobjVenda.objCupomFiscal.lNumeroDAV <> 0 Then
        If objItens.iStatus = STATUS_CANCELADO Then
                
            'cancelar o Item anterior
            lErro = AFRAC_CancelarItem(CInt(objItens.iItem))
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Cancelar Item")
            If lErro <> SUCESSO Then gError 210516
            
            'Atualiza o subtotal
            Subtotal.Caption = Format(Subtotal.Caption - ((objItens.dPrecoUnitario * objItens.dQuantidade) - objItens.dValorDesconto), "standard")
            
            If Grid.Width < 8000 Then
'                ListCF.AddItem Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", " ***** ITEM " & objItens.iItem & " CANCELADO *****")
                ListCF.AddItem Formata_Campo(ALINHAMENTO_DIREITA, 40, " ", "           ***** ITEM " & objItens.iItem & " CANCELADO *****") & Formata_Campo(ALINHAMENTO_ESQUERDA, 14, " ", "-" & Format(objItens.dPrecoUnitario * objItens.dQuantidade, "standard"))
                ListCF.ItemData(ListCF.NewIndex) = objItens.iItem
            Else
                Call Proxima_Linha_Grid
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_DESCRICAO) = "***** ITEM CANCELADO *****"
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_VALOR_TOTAL) = "-" & Format(objItens.dPrecoUnitario * objItens.dQuantidade, "standard")
                Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ITEM) = Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, objItens.iItem)
            
            End If
            
        End If
                
    Next
    
    If gobjVenda.objCupomFiscal.iVendedor <> 0 Then
        CodVendedor.Text = gobjVenda.objCupomFiscal.iVendedor
    Else
        CodVendedor.Text = ""
    End If
    
    Call CodVendedor_Validate(False)
    
    CodCliente.Text = gobjVenda.objCupomFiscal.sCPFCGC
    NomeCliente.Text = gobjVenda.objCupomFiscal.sNomeCliente
    CGC.Text = gobjVenda.objCupomFiscal.sCPFCGC1
    'Endereco.Text = gobjVenda.objCupomFiscal.sEndereco
    
    Call Limpa_Cupom_Tela
    
    Transforma_Cupom = SUCESSO
    
    Exit Function

Erro_Transforma_Cupom:
    
    Transforma_Cupom = gErr
    
    Select Case gErr
                  
        Case 99818, 99884, 99912, 204210, 210510, 210516
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175688)

    End Select
    
    Exit Function
        
End Function

Private Sub OptionDAV_Click()
    
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer
Dim lErro As Long
Dim iCodGerente As Integer
Dim objOperador As New ClassOperador

On Error GoTo Erro_OptionDAV_Click
    
    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 210007
    
    'Se tiver um Cupom em andamento
    If gobjVenda.objCupomFiscal.lNumero <> 0 And gobjVenda.iTipo <> OPTION_DAV And gobjVenda.iTipo <> OPTION_PREVENDA Then
        'Envia aviso perguntando se deseja cancelar o cupom em andamento
        vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_CANCELA_CUPOM)
    
        If vbMsgRes = vbYes Then
            'Se for Necessário a autorização do Gerente para abertura do Caixa
            If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then
        
                'Chama a Tela de Senha
                Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)
        
                'Sai de Função se a Tela de Login não Retornar ok
                If giRetornoTela <> vbOK Then gError 102501
                
                'Se Operador for Gerente
                iCodGerente = objOperador.iCodigo
            
            End If
            
            Exibe.Caption = "CANCELADO CUPOM " & gobjVenda.objCupomFiscal.lNumero
            Exibe1.Caption = ""
            DoEvents
            'Cancelar o Cupom de Venda
            lErro = AFRAC_CancelarCupom(Me, gobjVenda)
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Cancela Cupom")
            If lErro <> SUCESSO Then gError 99616
            
            Call Move_Dados_Memoria_1
            
            'Realiza as operações necessárias para gravar
            lErro = CF_ECF("Grava_Venda_Arquivo", gobjVenda)
            If lErro <> SUCESSO Then gError 204532
            
            lErro = CF_ECF("Grava_CancelamentoCupom_Arquivo", gobjVenda)
            If lErro <> SUCESSO Then gError 204533
            
            Exibe1.Caption = ""
            Exibe.Caption = "TRANSF.: VENDA EM ORÇAMENTO."
        Else
            Exibe.Caption = "PRÓXIMO CLIENTE..."
            Exibe1.Caption = ""
            'Seleciona o cupom
            OptionCF.Value = True
            Exit Sub
        End If
    End If
    
    'Quando eu clico no Orçamento reativa o número do orcamento(campo)
'    Orcamento.Enabled = True
'    LabelOrcamento.Enabled = True
    BotaoProxNum.Enabled = True
    gobjVenda.objCupomFiscal.lNumero = 0
    gobjVenda.iTipo = OPTION_DAV
    Orcamento.Text = ""
    
    BotaoAbrirGaveta.Caption = "(F10)   Grava Orçamento"
    BotaoCancelaCupom.Caption = "(Esc)   Imprime Orçamento"
    
    Exit Sub

Erro_OptionDAV_Click:

    Select Case gErr
                
        Case 99616, 102501, 204532, 204533, 210007
                            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175689)

    End Select
    
    Exit Sub
        
End Sub

Private Sub OptionPreVenda_Click()
    
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer
Dim lErro As Long
Dim iCodGerente As Integer
Dim objOperador As New ClassOperador

On Error GoTo Erro_OptionPreVenda_Click
    
    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 210006
    
    'Se tiver um Cupom em andamento
    If gobjVenda.objCupomFiscal.lNumero <> 0 And gobjVenda.iTipo <> OPTION_DAV And gobjVenda.iTipo <> OPTION_PREVENDA Then
        'Envia aviso perguntando se deseja cancelar o cupom em andamento
        vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_CANCELA_CUPOM)
    
        If vbMsgRes = vbYes Then
            'Se for Necessário a autorização do Gerente para abertura do Caixa
            If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then
        
                'Chama a Tela de Senha
                Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)
        
                'Sai de Função se a Tela de Login não Retornar ok
                If giRetornoTela <> vbOK Then gError 102501
                
                'Se Operador for Gerente
                iCodGerente = objOperador.iCodigo
            
            End If
            
            Exibe.Caption = "CANCELADO CUPOM " & gobjVenda.objCupomFiscal.lNumero
            Exibe1.Caption = ""
            DoEvents
            'Cancelar o Cupom de Venda
            lErro = AFRAC_CancelarCupom(Me, gobjVenda)
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Cancela Cupom")
            If lErro <> SUCESSO Then gError 99616
            
            Call Move_Dados_Memoria_1
            
            'Realiza as operações necessárias para gravar
            lErro = CF_ECF("Grava_Venda_Arquivo", gobjVenda)
            If lErro <> SUCESSO Then gError 204532
            
            lErro = CF_ECF("Grava_CancelamentoCupom_Arquivo", gobjVenda)
            If lErro <> SUCESSO Then gError 204533
            
            Exibe1.Caption = ""
            Exibe.Caption = "TRANSF.: VENDA EM ORÇAMENTO."
        Else
            Exibe.Caption = "PRÓXIMO CLIENTE..."
            Exibe1.Caption = ""
            'Seleciona o cupom
            OptionCF.Value = True
            Exit Sub
        End If
    End If
    
    'Quando eu clico no Orçamento reativa o número do orcamento(campo)
'    Orcamento.Enabled = True
'    LabelOrcamento.Enabled = True
    BotaoProxNum.Enabled = True
    gobjVenda.objCupomFiscal.lNumero = 0
    gobjVenda.iTipo = OPTION_PREVENDA
    Orcamento.Text = ""
    
    BotaoAbrirGaveta.Caption = "(F10)   Grava Orçamento"
    BotaoCancelaCupom.Caption = "(Esc)   Limpa Tela"
    
    Exit Sub

Erro_OptionPreVenda_Click:

    Select Case gErr
                
        Case 99616, 102501, 204532, 204533, 210006
                            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175689)

    End Select
    
    Exit Sub
        
End Sub

Private Sub BotaoPreco_Click()

Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iTipo As Integer
        
    'Chama tela de Preco
    Call Chama_TelaECF_Modal("Preco", objProduto)
    
    If giRetornoTela = vbOK Then
        ProdutoNomeRed.Text = objProduto.sCodigo
        Call ProdutoNomeRed_Validate(False)
    End If
     
End Sub

Private Function Adiciona_Cupom(ByVal iTransfOrcCF As Integer) As Long
'Pega o produto que deve estar em ProdutoNomeRed

Dim lErro As Long
Dim bAchou As Boolean
Dim objProduto As ClassProduto
Dim objItens As New ClassItemCupomFiscal
Dim sProduto As String
Dim lNum As Long
Dim lNumero As Long
Dim objAliquota As New ClassAliquotaICMS
Dim objCliente As ClassCliente
Dim sCPF As String
Dim sRet As String
Dim lErro1 As Long
Dim sPrecoItem As String
Dim sAliquota As String
Dim sPeso As String
Dim sPrecoKilo As String
Dim sPrecoTotal As String
Dim sProduto1 As String
Dim objAliquota1 As ClassAliquotaICMS
Dim iTotalizador As Integer
Dim objVenda As New ClassVenda

On Error GoTo Erro_Adiciona_Cupom
    
    If gobjVenda.iTipo = 0 Then gError 126815
    
    
    'se se tratar de DAV e ja tiver sido impresso ==> nao pode alterar
'    If gobjVenda.iTipo = OPTION_DAV And gobjVenda.objCupomFiscal.lNumero <> 0 Then gError 210507
    
        
    If gobjVenda.iTipo = OPTION_DAV Then
    
        objVenda.objCupomFiscal.lNumOrcamento = gobjVenda.objCupomFiscal.lNumOrcamento

        lErro = CF_ECF("OrcamentoECF_Le", objVenda)
        If lErro <> SUCESSO And lErro <> 204690 And lErro <> 210447 Then gError 210895

        'se o orcamento é um DAV e ja foi impresso ==> nao pode alterar
        If objVenda.iTipo = OPTION_DAV And objVenda.objCupomFiscal.iDAVImpresso <> 0 Then gError 210507
        
    End If
    
    giItem = gobjVenda.objCupomFiscal.colItens.Count + 1
    
    sProduto1 = ProdutoNomeRed.Text
    
    Call TP_Produto_Le_Col(gaobjProdutosReferencia, gaobjProdutosCodBarras, gaobjProdutosNome, sProduto1, objProduto)
            
    'caso o produto não seja encontrado
    If objProduto Is Nothing Then gError 99884
        
    'verifica se a figura foi preenchida
    If objProduto.sFigura <> "" Then
        'verifica se o arquivo é do tipo imagem
        sRet = Dir(objProduto.sFigura, vbNormal)
        If sRet <> "" Then
            If GetAttr(objProduto.sFigura) = vbArchive Or GetAttr(objProduto.sFigura) = vbArchive + vbReadOnly Then
                'coloca a figura na tela
                Figura.Picture = LoadPicture(objProduto.sFigura)
            End If
        Else
            gError 99607
        End If
    Else
        Figura.Picture = LoadPicture
    End If
    
    If iTransfOrcCF = 0 Then
        If objProduto.iUsaBalanca = USA_BALANCA And Len(Trim(gsBalancaPorta)) > 0 And giBalancaModelo > 0 Then
            lErro = AFRAC_Le_Balanca(gsBalancaPorta, giBalancaModelo, sPeso, sPrecoKilo, sPrecoTotal)
            lErro1 = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Leitura Balança")
            If lErro1 <> SUCESSO Or lErro <> SUCESSO Then gError 133754
            
            Quantidade.Text = sPeso
        End If
    End If
    
    For Each objAliquota In gcolAliquotasTotal
        If objAliquota.sSigla = objProduto.sICMSAliquota Then objItens.dAliquotaICMS = objAliquota.dAliquota
    Next
    
    objItens.dQuantidade = StrParaDbl(Format(Quantidade.Text, "0.000"))
    objItens.dPrecoUnitario = objProduto.dPrecoLoja
    objItens.sUnidadeMed = objProduto.sSiglaUMVenda
    objItens.sSituacaoTrib = objProduto.sSituacaoTribECF
    objItens.sProduto = objProduto.sCodigo
    objItens.sProdutoNomeRed = objProduto.sNomeReduzido
    
    objItens.iItem = giItem
    objItens.iStatus = STATUS_ATIVO
        
    lNum = Retorna_Count_ItensCupom
    
    'Abre o cupom caso seja o primeiro item deste cupom
    If gobjVenda.iTipo = OPTION_CF And lNum = 0 Then
        sCPF = gobjVenda.objCupomFiscal.sCPFCGC1
        lErro = CF_ECF("Abre_Cupom", gobjVenda)
        If lErro <> SUCESSO Then gError 99818
    End If
       
           
       
    sPrecoItem = Format(StrParaDbl(Format(Quantidade.Text, "0.000")) * StrParaDbl(Format(objProduto.dPrecoLoja, "standard")), "Standard")
    
    If objProduto.dDescontoValor > 0 Then
        objItens.dValorDesconto = objProduto.dDescontoValor
    ElseIf objProduto.dPercentMenosReceb > 0 Then
        objItens.dValorDesconto = Fix(StrParaDbl(sPrecoItem) * objProduto.dPercentMenosReceb) / 100
    End If

    If objItens.dAliquotaICMS > 0 Then
        If objProduto.sSituacaoTribECF = TIPOTRIBISS_SITUACAOTRIBECF_INTEGRAL Then
            
            sAliquota = TIPOTRIBISS_SITUACAOTRIBECF_INTEGRAL & Format(objItens.dAliquotaICMS * 10000, "0000")
            
'                For Each objAliquota1 In gobjLojaECF.colAliquotaICMS
'                    If objAliquota1.iISS = 1 And objItens.dAliquotaICMS = objAliquota1.dAliquota Then
'                        Exit For
'                    End If
'                    iTotalizador = iTotalizador + 1
'                Next

'                objItens.sSituacaoTrib = Format(iTotalizador, "00") & sAliquota
            objItens.sSituacaoTrib = sAliquota
            
        Else
            
'                For Each objAliquota1 In gobjLojaECF.colAliquotaICMS
'                    If objAliquota1.iISS = 0 And objItens.dAliquotaICMS = objAliquota1.dAliquota Then
'                        Exit For
'                    End If
'                    iTotalizador = iTotalizador + 1
'                Next
            
            sAliquota = Format(objItens.dAliquotaICMS * 10000, "0000")
'                objItens.sSituacaoTrib = Format(iTotalizador, "00") & TIPOTRIBICMS_SITUACAOTRIBECF_INTEGRAL & sAliquota
            objItens.sSituacaoTrib = TIPOTRIBICMS_SITUACAOTRIBECF_INTEGRAL & sAliquota
            
        End If
    Else
       'colocando o 1 para ficar o codigo F1, I1, N1
       sAliquota = left(objProduto.sSituacaoTribECF, 1)
       objItens.sSituacaoTrib = sAliquota & "1"
    End If
        
    If gobjVenda.iTipo = OPTION_CF Then
        'Vende o item
        
        
        lErro = AFRAC_VenderItem(objProduto.sCodigo, objProduto.sDescricao, StrParaDbl(Format(Quantidade.Text, "0.000")), Format(objProduto.dPrecoLoja, "standard"), 1, 1, objItens.dValorDesconto, StrParaDbl(Format(StrParaDbl(Format(Quantidade.Text, "0.000")) * StrParaDbl(Format(objProduto.dPrecoLoja, "standard")), "Standard")), sAliquota, objProduto.sSiglaUMVenda, False)
        lErro1 = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Vender Item")
        If lErro1 <> SUCESSO Or lErro <> SUCESSO Then gError 99912
        
        lErro = CF_ECF("Requisito_XXII_AtualizaGT")
        If lErro <> SUCESSO Then gError 210424
        
        
    End If

    
    'Prenche a col de itens do cupom com os dados do mesmo
    PrecoUnitario.Caption = Format(objProduto.dPrecoLoja, "standard")
    ProdutoNomeRed.Text = objProduto.sNomeReduzido
    PrecoItem.Caption = sPrecoItem
    Subtotal.Caption = Format(StrParaDbl(Subtotal.Caption) + (StrParaDbl(PrecoItem.Caption)), "standard")
       
    'Joga na col
    gobjVenda.objCupomFiscal.colItens.Add objItens
    
    'Joga no cupom o item
    
    If Grid.Width < 8000 Then

        ListCF.AddItem Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, objItens.iItem) & "     " & Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", objProduto.sCodigo) & Formata_Campo(ALINHAMENTO_DIREITA, 30, " ", objProduto.sDescricao)
        ListCF.ItemData(ListCF.NewIndex) = objItens.iItem
        ListCF.AddItem Formata_Campo(ALINHAMENTO_ESQUERDA, 11, " ", Format(Quantidade.Text, "#0.000")) & "  " & Formata_Campo(ALINHAMENTO_DIREITA, 4, " ", objProduto.sSiglaUMVenda) & " x " & Formata_Campo(ALINHAMENTO_DIREITA, 12, " ", Format(PrecoUnitario.Caption, "standard")) & objProduto.sSituacaoTribECF & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(objItens.dAliquotaICMS * 100, "fixed") & "%") & Formata_Campo(ALINHAMENTO_ESQUERDA, 14, " ", Format(PrecoItem.Caption, "standard"))
        ListCF.ItemData(ListCF.NewIndex) = objItens.iItem
        
    Else
    
        Call Proxima_Linha_Grid

        Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ITEM) = Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, objItens.iItem)
        Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_CODIGO) = objProduto.sCodigo
        Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_DESCRICAO) = objProduto.sDescricao
        Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_QUANTIDADE) = Format(Quantidade.Text, "#0.000") & objProduto.sSiglaUMVenda
        Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_X) = "x"
        Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_VALOR_UNITARIO) = Format(PrecoUnitario.Caption, "standard")
        Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ST) = objProduto.sSituacaoTribECF & Format(objItens.dAliquotaICMS * 100, "fixed") & "%"
        Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_VALOR_TOTAL) = Format(PrecoItem.Caption, "standard")
    
    End If
    
    
    'se existir desconto sobre o item...
    If objProduto.dPercentMenosReceb > 0 Then
        
        If Grid.Width < 8000 Then
        
'        ListCF.AddItem Formata_Campo(ALINHAMENTO_DIREITA, 9, " ", "DESCONTO:") & Formata_Campo(ALINHAMENTO_ESQUERDA, 10, " ", Format(objProduto.dPercentMenosReceb, "fixed") & "%") & Formata_Campo(ALINHAMENTO_ESQUERDA, 11, " ", "-" & Format(Fix(StrParaDbl(PrecoItem.Caption) * (objProduto.dPercentMenosReceb)) / 100, "standard")) & Formata_Campo(ALINHAMENTO_ESQUERDA, 20, " ", Format(StrParaDbl(PrecoItem.Caption) - objItens.dValorDesconto, "standard"))
            ListCF.AddItem Formata_Campo(ALINHAMENTO_DIREITA, 9, " ", "DESCONTO:") & Formata_Campo(ALINHAMENTO_ESQUERDA, 10, " ", Format(objProduto.dPercentMenosReceb, "fixed") & "%") & Formata_Campo(ALINHAMENTO_ESQUERDA, 11, " ", "-" & Format(objItens.dValorDesconto, "standard")) & Formata_Campo(ALINHAMENTO_ESQUERDA, 20, " ", Format(StrParaDbl(PrecoItem.Caption) - objItens.dValorDesconto, "standard"))
            ListCF.ItemData(ListCF.NewIndex) = objItens.iItem
'        Subtotal.Caption = Format(StrParaDbl(Subtotal.Caption) - (StrParaDbl(PrecoItem.Caption)), "standard")
'        Subtotal.Caption = Format(StrParaDbl(Subtotal.Caption) + (StrParaDbl(PrecoItem.Caption) - (StrParaDbl(PrecoItem.Caption) * objProduto.dPercentMenosReceb / 100)), "standard")

        Else
        
            Call Proxima_Linha_Grid
            
            Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_DESCRICAO) = "DESCONTO:"
            Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_VALOR_UNITARIO) = Format(objProduto.dPercentMenosReceb, "fixed") & "%"
            Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ST) = "-" & Format(objItens.dValorDesconto, "standard")
            Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_VALOR_TOTAL) = Format(StrParaDbl(PrecoItem.Caption) - objItens.dValorDesconto, "standard")
            Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ITEM) = Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, objItens.iItem)

        End If
        
        Subtotal.Caption = Format(StrParaDbl(Subtotal.Caption) - objItens.dValorDesconto, "standard")
        
        
    ElseIf objProduto.dDescontoValor > 0 Then
    
        If Grid.Width < 8000 Then
    
            ListCF.AddItem Formata_Campo(ALINHAMENTO_DIREITA, 9, " ", "DESCONTO:") & Formata_Campo(ALINHAMENTO_ESQUERDA, 21, " ", "-" & Format(objItens.dValorDesconto, "standard")) & Formata_Campo(ALINHAMENTO_ESQUERDA, 20, " ", Format(StrParaDbl(PrecoItem.Caption) - objItens.dValorDesconto, "standard"))
            ListCF.ItemData(ListCF.NewIndex) = objItens.iItem
    '        Subtotal.Caption = Format(StrParaDbl(Subtotal.Caption) - (StrParaDbl(PrecoItem.Caption)), "standard")
    '        Subtotal.Caption = Format(StrParaDbl(Subtotal.Caption) + (StrParaDbl(PrecoItem.Caption) - objProduto.dDescontoValor), "standard")
    
        Else
        
            Call Proxima_Linha_Grid
            
            Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_DESCRICAO) = "DESCONTO:"
            Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ST) = "-" & Format(objItens.dValorDesconto, "standard")
            Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_VALOR_TOTAL) = Format(StrParaDbl(PrecoItem.Caption) - objItens.dValorDesconto, "standard")
            Grid.TextMatrix(giUltimaLinhaGrid, GRID_COL_ITEM) = Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, objItens.iItem)
        
        End If
        
        Subtotal.Caption = Format(StrParaDbl(Subtotal.Caption) - objItens.dValorDesconto, "standard")
        
    
    End If
    
    
    'Para rolar automaticamente a barra de rolagem
    If Grid.Width < 8000 Then
        ListCF.ListIndex = ListCF.NewIndex
    End If
    
    Exibe.Caption = Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, objItens.iItem) & "   " & Formata_Campo(ALINHAMENTO_DIREITA, 20, " ", objProduto.sNomeReduzido)
    
'    If objProduto.dPercentMenosReceb > 0 Then
'        Exibe1.Caption = Formata_Campo(ALINHAMENTO_ESQUERDA, 8, " ", Format(Quantidade.Text, "0.000")) & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 4, " ", "x") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(PrecoUnitario.Caption, "standard")) & "-" & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(objItens.dValorDesconto, "standard")) & Formata_Campo(ALINHAMENTO_ESQUERDA, 10, " ", Format(StrParaDbl(PrecoItem.Caption) - objItens.dValorDesconto, "standard"))
'    ElseIf objProduto.dDescontoValor > 0 Then
'        Exibe1.Caption = Formata_Campo(ALINHAMENTO_ESQUERDA, 8, " ", Format(Quantidade.Text, "0.000")) & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 4, " ", "x") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(PrecoUnitario.Caption, "standard")) & "-" & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(objItens.dValorDesconto, "standard")) & Formata_Campo(ALINHAMENTO_ESQUERDA, 10, " ", Format(StrParaDbl(PrecoItem.Caption) - objItens.dValorDesconto, "standard"))
'    Else
'        Exibe1.Caption = Formata_Campo(ALINHAMENTO_ESQUERDA, 8, " ", Format(Quantidade.Text, "0.000")) & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 4, " ", "x") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(PrecoUnitario.Caption, "standard")) & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 4, " ", "=") & Formata_Campo(ALINHAMENTO_ESQUERDA, 10, " ", Format(PrecoItem.Caption, "standard"))
'    End If
    
    If objProduto.dPercentMenosReceb > 0 Or objProduto.dDescontoValor > 0 Then
        Exibe1.Caption = Formata_Campo(ALINHAMENTO_ESQUERDA, 8, " ", Format(Quantidade.Text, "0.000")) & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 4, " ", "x") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(PrecoUnitario.Caption, "standard")) & "-" & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(objItens.dValorDesconto, "standard")) & Formata_Campo(ALINHAMENTO_ESQUERDA, 10, " ", Format(StrParaDbl(PrecoItem.Caption) - objItens.dValorDesconto, "standard"))
    Else
        Exibe1.Caption = Formata_Campo(ALINHAMENTO_ESQUERDA, 8, " ", Format(Quantidade.Text, "0.000")) & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 4, " ", "x") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(PrecoUnitario.Caption, "standard")) & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 4, " ", "=") & Formata_Campo(ALINHAMENTO_ESQUERDA, 10, " ", Format(PrecoItem.Caption, "standard"))
    End If
    
    Set objItens.objTributacaoDocItem = New ClassTributacaoDocItem
    Call objItens.objTributacaoDocItem.Copia(objProduto.objTributacaoDocItem)
    'ajusta de acordo com o que efetivamente foi vendido
    Call ItemCupom_AjustaTrib(objItens)
    
    Call Limpa_Cupom_Tela
    
    ProdutoNomeRed.SetFocus
    
    Adiciona_Cupom = SUCESSO
    
    Exit Function

Erro_Adiciona_Cupom:
    
    Adiciona_Cupom = gErr
    
    Select Case gErr
                
        Case 99607
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FIGURA_INVALIDO, gErr, objProduto.sFigura)
                    
        Case 99818, 99884, 210895
        
        Case 99912, 133754
            ProdutoNomeRed.Text = ""
        
        Case 126815
            Call Rotina_ErroECF(vbOKOnly, ERRO_TIPOCF_NAO_ESCOLHIDO, gErr)
        
        Case 210507
            Call Rotina_ErroECF(vbOKOnly, ERRO_DAV_NAO_ALTERADO_DEPOIS_DE_IMPRESSO, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175690)

    End Select
    
    Exit Function
    
End Function

Private Function Retorna_Count_ItensCupom() As Long

Dim objItens As ClassItemCupomFiscal
Dim lNum As Long
    
    
'    lNum = 0
'
'    For Each objItens In gobjVenda.objCupomFiscal.colItens
'        'Caso não seja um item cancelado
'        If objItens.iStatus = STATUS_ATIVO Then lNum = lNum + 1
'    Next
    
'    Retorna_Count_ItensCupom = lNum
    Retorna_Count_ItensCupom = gobjVenda.objCupomFiscal.colItens.Count
    
End Function

Private Sub Limpa_Tela_Venda()

Dim lErro As Long
Dim objOrcamento As Object

On Error GoTo Erro_Limpa_Tela_Venda

    Call Limpa_Cupom_Tela
        
    CGC.Text = ""
    CodCliente.Text = ""
    NomeCliente.Text = ""
'    Endereco.Text = ""
        
    Call Limpa_Tela_Venda_1
    
    ProdutoNomeRed.SetFocus
    
    If giOrcamentoECF = CAIXA_SO_ORCAMENTO Then
    
'        If giPrevenda = 1 And giUsaImpressoraFiscal = 0 Then
'            OptionPreVenda.Value = False
'            OptionPreVenda.Value = True
'        ElseIf giDAV = 1 Then
'            OptionDAV.Value = False
'            OptionDAV.Value = True
'        End If
            
        
        Set objOrcamento = Orcamento
        
        lErro = CF_ECF("Retorna_Prox_Orcamento", objOrcamento, gobjVenda)
        If lErro <> SUCESSO Then gError 199466
    Else
'        OptionCF.Value = False
'        OptionCF.Value = True
    End If
    
    
    If OptionPreVenda.Value Then gobjVenda.iTipo = OPTION_PREVENDA
    
    If OptionDAV.Value Then gobjVenda.iTipo = OPTION_DAV

    If OptionCF.Value Then gobjVenda.iTipo = OPTION_CF
    
    
    Exit Sub

Erro_Limpa_Tela_Venda:
    
    Select Case gErr
                
        Case 199466
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 199467)

    End Select
    
    Exit Sub
    
End Sub

Private Sub Limpa_Tela_Venda_1()

    Exibe.Caption = "PRÓXIMO CLIENTE..."
    Exibe1.Caption = ""
    
    giItem = 0
    Figura.Picture = LoadPicture
    
    Subtotal.Caption = Format(0, "standard")
    CodVendedor.Text = ""
    CodCliente.Text = ""
    LabelVendedor.Caption = ""
    NomeCliente.Text = ""
    
    Orcamento.Text = ""
'    Orcamento.Enabled = False
'    LabelOrcamento.Enabled = False
    
    'Limpa toda a list
    ListCF.Clear
    
    'Inicia o Cupom
    Call Inicia_Cupom_Tela

End Sub

Private Sub Limpa_Cupom_Tela()

    PrecoUnitario.Caption = Format(0, "standard")
    PrecoItem.Caption = Format(0, "standard")
    ProdutoNomeRed.Text = ""
    Quantidade.Text = 1
    
End Sub


Private Sub Timer1_Timer()
    
'Dim sHora As String
'Dim iPosHora As Integer
'Dim sMinuto As String
'Dim iPosMinuto As Integer
'Dim ssegundo As String
'Dim iPossegundo As Integer
Dim vbMsgBox As VbMsgBoxResult
Dim bAchou As Boolean
Dim lErro As Long
Dim dTimerTemp As Double
Dim dtData As Date
Dim dtime As Double
Dim dtTime As Date
Dim lSequencial As Long
Dim sHora As String
Dim dtUltimaReducaoECF As Date

On Error GoTo Erro_Timer1_Timer

    dtTime = Time


'    dtime = Timer
'    If dtime > 3600 Then
'        'Coloca a hora atual do Sistema
'        sHora = CStr(dtime / (60 * 60))
'        iPosHora = InStr(1, sHora, ",")
'        If iPosHora > 0 Then sHora = Mid(sHora, 1, iPosHora - 1)
'    Else
'        sHora = 0
'    End If
'
'    If sHora <> 0 Then
'        dTimerTemp = dtime - (CLng(sHora * 3600))
'    Else
'        dTimerTemp = dtime
'    End If
'
'    If dTimerTemp > 60 Then
'        sMinuto = CStr(dtime / 60) - (CInt(sHora * 60))
'        iPosMinuto = InStr(1, sMinuto, ",")
'        If iPosMinuto > 0 Then sMinuto = Mid(sMinuto, 1, iPosMinuto - 1)
'    Else
'        sMinuto = 0
'    End If
'
'    ssegundo = CStr(dtime) - ((CLng(sMinuto * 60)) + (CLng(sHora * 3600)))
'    iPossegundo = InStr(1, ssegundo, ",")
'    If iPossegundo > 0 Then ssegundo = Mid(ssegundo, 1, iPossegundo - 1)
'
''    DataHora.Caption = Format(Date, "dd/mm/yy") & "   " & Format(sHora, "00") & ":" & Format(sMinuto, "00") & ":" & Format(ssegundo, "00")
'
'    Me.Caption = Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", "Venda") & "Filial: " & giFilialEmpresa & "    Caixa: " & giCodCaixa & "    Operador: " & gsNomeOperador & _
'    "    Data: " & Format(Date, "dd/mm/yyyy") & "     Hora: " & Format(sHora, "00") & ":" & Format(sMinuto, "00") & ":" & Format(ssegundo, "00") & "    Empresa: " & Formata_Campo(ALINHAMENTO_DIREITA, 50, " ", gsNomeEmpresa)
    
    Me.Caption = Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", "Venda") & "Filial: " & giFilialEmpresa & "    Caixa: " & giCodCaixa & "    Operador: " & gsNomeOperador & _
    "    Data: " & Format(Date, "dd/mm/yyyy") & "     Hora: " & Format(dtTime, "Hh:Nn:Ss") & "    Empresa: " & Formata_Campo(ALINHAMENTO_DIREITA, 50, " ", gsNomeEmpresa)
    
    DoEvents
    
    If giOrcamentoECF <> CAIXA_SO_ORCAMENTO Then
    
        If gdtDataAnterior <> Date Then
        
            lErro = AFRAC_DataReducao(dtUltimaReducaoECF)
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Data Ultima Reducao Z")
            If lErro <> SUCESSO Then gError 214031
                    
        
            'se a hora > 1 ou se a data da ultima reducao registrada pelo sistema é menor q a data da ultima reducao do ECF--> fecha o caixa
            If Hour(dtTime) > 1 Or gdtUltimaReducao < dtUltimaReducaoECF Then
            
                If gobjVenda.objCupomFiscal.lNumero = 0 Then
                
                    lErro = CF_ECF("Requisito_XXII")
                    If lErro <> SUCESSO Then gError 210013
                
                    'messagem avisando dá mudança de dia
                    vbMsgBox = Rotina_AvisoECF(vbOKOnly, AVISO_INICIALIZAR_SISTEMA_AGORA)
                    
                    lErro = CF_ECF("Caixa_Executa_Fechamento")
                    If lErro <> SUCESSO Then gError 118007
                    
'                    gdtDataHoje = Date
                    
                    lErro = CF_ECF("Carrega_Arquivo_FonteDados")
                    If lErro <> SUCESSO Then gError 210197
                    
                    lErro = CF_ECF("Requisito_VII_Item8_E2")
                    If lErro <> SUCESSO Then gError 210198
                    
                    lErro = CF_ECF("Caixa_Executa_Abertura")
                    If lErro <> SUCESSO Then gError 118014
                    
                    If giOrcamentoECF <> CAIXA_SO_ORCAMENTO Then
                    
                        lErro = AFRAC_AbrirDia(Date)
                        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Abertura do Dia")
                        If lErro <> SUCESSO Then gError 118015
                    
                    End If
                    
                    'Função que Faz a Abertura de Sessão
                    lErro = CF_ECF("Sessao_Executa_Abertura")
                    If lErro <> SUCESSO Then gError 108016
                End If
            ElseIf Minute(dtTime) = 30 Or Minute(dtTime) = 0 Then
                If Minute(dtTime) <> gsMinutoAnt Then
                    'messagem avisando dá mudança de dia
                    vbMsgBox = Rotina_AvisoECF(vbOKOnly, AVISO_INICIALIZAR_SISTEMA)
                    gsMinutoAnt = Minute(dtTime)
                End If
            End If
        End If
    
    End If
    
    Exit Sub
    
Erro_Timer1_Timer:
    
    Select Case gErr
        
        Case 118007, 118014 To 118016, 126153, 126155, 126156, 126157, 210013, 210197, 210198, 214031
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175691)

    End Select
    
    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
Dim bCancel As Boolean
    
    If KeyCode = 13 Then
    
        If Me.ActiveControl Is ProdutoNomeRed And Len(Trim(Quantidade.Text)) = 0 Then
            Quantidade.SetFocus
        Else
            KeyCode = vbKeyTab
        End If

'        If Me.ActiveControl Is ProdutoNomeRed Then
'            If Len(ProdutoNomeRed.Text) = 0 Or Len(Quantidade.Text) = 0 Then
'                Quantidade.SetFocus
'            Else
'                Call ProdutoNomeRed_Validate(bCancel)
'            End If
'        ElseIf Me.ActiveControl Is Quantidade Then
'            If Len(ProdutoNomeRed.Text) = 0 Or Len(Quantidade.Text) = 0 Then
'                ProdutoNomeRed.SetFocus
'            Else
'                Call Quantidade_Validade(bCancel)
'        End If
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is CodVendedor Then
            Call LabelVendedor_Click
        ElseIf Me.ActiveControl Is CodCliente Then
            Call LabelCliente_Click
        ElseIf Me.ActiveControl Is Orcamento Then
            Call LabelOrcamento_Click
        ElseIf Me.ActiveControl Is ProdutoNomeRed Then
            Call BotaoProdutos_Click
        End If
    End If
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = vbKeyF4 Then
        Call BotaoCancelaItemAtual_Click
    End If

    If KeyCode = vbKeyF5 Then
        Call BotaoPreco_Click
    End If

    If KeyCode = vbKeyF6 Then
        Call BotaoCancelaItem_Click
    End If

    If KeyCode = vbKeyF7 Then
        Call BotaoSuspender_Click
    End If

    If KeyCode = vbKeyF8 Then
        Call BotaoFechar_Click
    End If
  
    If KeyCode = vbKeyF9 Then
        Call BotaoProdutos_Click
    End If

    If KeyCode = vbKeyF10 Then
        Call BotaoAbrirGaveta_Click
    End If

    If KeyCode = vbKeyEscape Then
        Call BotaoCancelaCupom_Click
    End If

    If KeyCode = vbKeyF11 Then
        Call BotaoPagamento_Click
    End If

    If KeyCode = vbKeyF12 Then
        Call LabelOrcamento_Click
    End If
    
    If KeyCode = vbKeyF1 And Shift = 2 Then
        Call BotaoAtualizar_Click
    End If

End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    'Libera a referência da tela
    Set gobjVenda = Nothing
    
    If LeitoraCodBarras.PortOpen = True Then LeitoraCodBarras.PortOpen = False
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer
Dim lErro As Long
Dim lNum As Long
Dim objOperador As New ClassOperador
Dim iCodGerente As Integer

On Error GoTo Erro_Form_QueryUnload

    If ListCF.ListIndex <> -1 Or giUltimaLinhaGrid > 1 Then

        Timer1.Enabled = False
    
        lErro = CF_ECF("Requisito_XXII")
        If lErro <> SUCESSO Then gError 210014
    
        lNum = Retorna_Count_ItensCupom
        
        
        If gobjVenda.objCupomFiscal.lNumero <> 0 And gobjVenda.iTipo = OPTION_CF Then
            'Envia aviso perguntando se deseja cancelar a venda
            vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_CANCELA_VENDA)
    
            If vbMsgRes = vbYes Then
                'Se for Necessário a autorização do Gerente para abertura do Caixa
                If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then
            
                    'Chama a Tela de Senha
                    Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)
            
                    'Sai de Função se a Tela de Login não Retornar ok
                    If giRetornoTela <> vbOK Then gError 102506
                    
                    'Se Operador for Gerente
                    iCodGerente = objOperador.iCodigo
            
                End If
    
                'Cancelar o Cupom de Venda
                lErro = AFRAC_CancelarCupom(Me, gobjVenda)
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Cancelar Cupom")
                If lErro <> SUCESSO Then gError 99617
                
                Call Move_Dados_Memoria_1
                
                'Realiza as operações necessárias para gravar
                lErro = CF_ECF("Grava_Venda_Arquivo", gobjVenda)
                If lErro <> SUCESSO Then gError 204534
                
                lErro = CF_ECF("Grava_CancelamentoCupom_Arquivo", gobjVenda)
                If lErro <> SUCESSO Then gError 204535
                
                
            Else
                Cancel = True
            End If
            
        End If
        
        'Se tiver um orçamento na tela
        If lNum > 0 And (gobjVenda.iTipo = OPTION_DAV Or gobjVenda.iTipo = OPTION_PREVENDA) Then
            
            'Envia aviso que o orçamento será cancelado
            vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_CANCELA_ORCAMENTO)
    
            If vbMsgRes = vbNo Then Cancel = True
        
        End If
        
        If Cancel = True Then Timer1.Enabled = True
    
    End If
    
    Exit Sub
    
Erro_Form_QueryUnload:

    Select Case gErr
                
        Case 99617, 102506, 204534, 204535, 210014
            Cancel = True
            Timer1.Enabled = True
                            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175692)

    End Select
    
    Exit Sub
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object
Dim objOperador As ClassOperador
Dim sOper As String

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Call Form_Load


End Function

Public Function Name() As String

    Name = "VendaNFe"

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

Private Sub LabelCliente_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection
        
    objCliente.sNomeReduzido = CodCliente.Text
        
    'Chama Tela ClienteLista
    Call Chama_TelaECF_Modal("ClienteLista", objCliente)
        
    If giRetornoTela = vbOK Then
    
        Select Case Len(Trim(objCliente.sCgc))
    
            Case STRING_CPF 'CPF
                
                'Formata e coloca na Tela
                CodCliente.Format = "000\.000\.000-00; ; ; "
    
            Case STRING_CGC 'CGC
                
                'Formata e Coloca na Tela
                CodCliente.Format = "00\.000\.000\/0000-00; ; ; "

        End Select

        CodCliente.Text = objCliente.sCgc
        NomeCliente.Text = objCliente.sNomeReduzido
        gobjVenda.objCupomFiscal.sCPFCGC = objCliente.sCgc
        gobjVenda.objCupomFiscal.sNomeCliente = objCliente.sNomeReduzido
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

Private Sub BotaoAbrirGaveta_Click()
    
Dim iCodigo As Integer
Dim lErro As Long
Dim objMovCaixa As ClassMovimentoCaixa
Dim bAchou As Boolean
Dim iIndice As Integer
Dim sMsg As String
Dim sIndice As String
Dim lNum As Long
Dim sDesc As String
Dim objTiposMeioPagto As ClassTMPLoja
Dim objVenda As New ClassVenda
Dim lSequencial As Long
Dim sLog As String
Dim colRegistro As New Collection
Dim objCliente As ClassCliente
Dim sCPF As String
Dim lNumero As Long
Dim dtDataFinal As Date
Dim vbMsgRes As VbMsgBoxResult
Dim objProdutoNomeRed As Object
Dim sMsgVendedor As String
Dim objVendedor As ClassVendedor
Dim objTela As Object
Dim sOrcamento As String
Dim objOrcamentoOrc As ClassVenda
Dim colOrcamento As New Collection
Dim sRetorno As String

On Error GoTo Erro_BotaoAbrirGaveta_Click
            
    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 210003
            
            
    'Se não há valor para pagar --> erro.
    If StrParaDbl(Subtotal.Caption) = 0 Then gError 99889

    
    'sevé obrigatório o preenchimento do vendedor
    If gobjLojaECF.iVendedorObrigatorio = 1 Then
        If Len(Trim(CodVendedor.Text)) = 0 Then gError 112072
    End If
    
    If gobjVenda.iTipo = OPTION_PREVENDA Or gobjVenda.iTipo = OPTION_DAV Then
    
        gobjVenda.objCupomFiscal.iTipo = gobjVenda.iTipo
        gobjVenda.objCupomFiscal.dtDataEmissao = Date
        gobjVenda.objCupomFiscal.dHoraEmissao = CDbl(Time)
        gobjVenda.objCupomFiscal.dValorTroco = 0
        gobjVenda.objCupomFiscal.lDuracao = 0
        gobjVenda.objCupomFiscal.iFilialEmpresa = giFilialEmpresa
        gobjVenda.objCupomFiscal.iCodCaixa = giCodCaixa
        gobjVenda.objCupomFiscal.iECF = giCodECF
        gobjVenda.objCupomFiscal.iTabelaPreco = gobjLojaECF.iTabelaPreco
        gobjVenda.objCupomFiscal.dValorTotal = StrParaDbl(Subtotal.Caption)
        gobjVenda.objCupomFiscal.dValorProdutos = gobjVenda.objCupomFiscal.dValorTotal
        gobjVenda.objCupomFiscal.dtDataReducao = gdtDataAnterior
        
        'resseta estas variaveis para ficar o deposito todo o pagamento em dinheiro
        Set gobjVenda.colMovimentosCaixa = New Collection
        Set gobjVenda.colCheques = Nothing
        Set gobjVenda.colTroca = Nothing
        Set gobjVenda.objCarne = Nothing
        
        Set objMovCaixa = New ClassMovimentoCaixa
        
        objMovCaixa.dtDataMovimento = Date
        objMovCaixa.dValor = StrParaDbl(Subtotal.Caption)
        objMovCaixa.iFilialEmpresa = giFilialEmpresa
        objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_DINHEIRO
        objMovCaixa.iCaixa = giCodCaixa
        objMovCaixa.iCodOperador = giCodOperador
        objMovCaixa.dHora = CDbl(Time)
        objMovCaixa.iAdmMeioPagto = MEIO_PAGAMENTO_DINHEIRO
                    
        gobjVenda.colMovimentosCaixa.Add objMovCaixa
        
        
        If gobjVenda.iTipo = OPTION_PREVENDA Then
            'grava a prevenda
            lErro = CF_ECF("Grava_Orcamento_ECF1", gobjVenda)
            If lErro <> SUCESSO Then gError 105895
                
        Else
        
'        'Envia aviso perguntando se deseja imprimir o orçamento
'        vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_ORCAMENTO_IMPRESSAO)
        
'        If vbMsgRes = vbYes Then

'         If giImprimeOrc = 1 Then 'PAFECF


            'se for dav é ja tiver sido impresso ==> nao imprime nem altera o DAV
'            If gobjVenda.objCupomFiscal.lNumeroDAV <> 0 Then gError 210504

            objVenda.objCupomFiscal.lNumOrcamento = gobjVenda.objCupomFiscal.lNumOrcamento


            lErro = CF_ECF("OrcamentoECF_Le", objVenda)
            If lErro <> SUCESSO And lErro <> 204690 And lErro <> 210447 Then gError 210893

            'se for dav é ja tiver sido impresso ==> nao imprime nem altera o DAV
            If gobjVenda.iTipo = OPTION_DAV And objVenda.objCupomFiscal.iDAVImpresso <> 0 Then gError 210504

            'se o orcamento ja tiver sido usado para gerar cupom fiscal ==> nao pode alterar
            If lErro = 210447 Then gError 210995


            'se for DAV grava e imprime
            dtDataFinal = gobjVenda.objCupomFiscal.dtDataEmissao

            Set objTela = Me
            
            'le os registros do orcamento e loca o arquivo
'            lErro = CF_ECF("Imprime_OrcamentoECF", dtDataFinal, gobjVenda.objCupomFiscal.lNumOrcamento, objTela, gobjVenda)
'            If lErro <> SUCESSO Then gError 105887
        
            lErro = CF_ECF("Grava_Orcamento_ECF1", gobjVenda)
            If lErro <> SUCESSO Then gError 105887
        
        End If
        
        Set gobjVenda = New ClassVenda
        
        Call Limpa_Tela_Venda
        
    Else
    
        'se for um cupom fiscal
        
        Exibe.Caption = "ENCERRAMENTO DE VENDA"
        Exibe1.Caption = ""
        DoEvents
               
        'Preenche as col's globais
        
        Call Move_Dados_Memoria
            
        'informar se for um CF
        If gobjVenda.iTipo = OPTION_CF Then
            sIndice = TIPOMEIOPAGTOLOJA_DINHEIRO
            'Recolhe a descrição
            For Each objTiposMeioPagto In gcolTiposMeiosPagtos
                'Se o tipo for dinheiro
                If objTiposMeioPagto.iTipo = StrParaInt(sIndice) Then
                    sDesc = objTiposMeioPagto.sDescricao
                    Exit For
                End If
            Next
            
    '        'If gobjlojaecf.iImprimeItemAItem = DESMARCADO Then
    '            lErro = Transforma_Cupom
    '            If lErro <> SUCESSO Then gError 112074
    '        'End If
    '
            lErro = AFRAC_AcrescimoDescontoCupom(0, 0, "", "")
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Acrescimo - Desconto")
            If lErro <> SUCESSO Then gError 109691
                
            'Informa o pagamento
            lErro = AFRAC_FormaPagamento(sDesc, sIndice, Subtotal.Caption, sMsg)
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Forma Pagamento")
            If lErro <> SUCESSO Then gError 99618
                
            If gobjVenda.objCupomFiscal.iVendedor <> 0 Then

                For Each objVendedor In gcolVendedores
                    'verifica se existe o vendedor na col
                    If objVendedor.iCodigo = gobjVenda.objCupomFiscal.iVendedor Then
                        Exit For
                    End If
                Next
            
                sMsgVendedor = TRACO_CAB
                sMsgVendedor = sMsgVendedor & VENDEDOR_ECF_MSG & Formata_Campo(ALINHAMENTO_DIREITA, 38, " ", objVendedor.iCodigo & " - " & objVendedor.sNomeReduzido)
                sMsgVendedor = sMsgVendedor & TRACO_CAB
                
            End If
            
            'se o cupom esta sendo originado de um orcamento
            If gobjVenda.objCupomFiscal.lNumOrcamento <> 0 Then
            
                'se for uma PREVENDA
                If gobjVenda.objCupomFiscal.lNumeroDAV = 0 Then
                    sOrcamento = "PV""" & Format(gobjVenda.objCupomFiscal.lNumOrcamento, "0000000000") & """ "
                Else
                    sOrcamento = "DAV""" & Format(gobjVenda.objCupomFiscal.lNumeroDAV, "0000000000") & """ "
                End If
                
            End If
            
            'Fecha cupom
            lErro = AFRAC_FecharCupom(Me, gobjVenda, False, gobjVenda.objCupomFiscal.sCPFCGC1, NomeCliente.Text, gobjVenda.objCupomFiscal.sEndereco, False, sOrcamento & sMsgVendedor)
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Fechar Cupom")
            If lErro <> SUCESSO Then gError 99619
                        
            'Abri a Gaveta
            lErro = AFRAC_AbrirGaveta()
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Abrir Gaveta")
            If lErro <> SUCESSO Then gError 99613
                            
        End If
            
        'Realiza as operações necessárias para gravar
        lErro = CF_ECF("Grava_Venda_Arquivo", gobjVenda)
        If lErro <> SUCESSO Then gError 109823
           
        'Atribui para a coleção global o objvenda
        gcolVendas.Add gobjVenda
        
            'Atualiza o arquivo
        If gobjVenda.iTipo = OPTION_CF Then Call WritePrivateProfileString(APLICACAO_ECF, "CupomAberto", "0", NOME_ARQUIVO_CAIXA)
                
        Set gobjVenda = New ClassVenda
        
        Call Limpa_Tela_Venda
        
        If gobjLojaECF.iAbreAposFechamento = MARCADO Then
            sCPF = gobjVenda.objCupomFiscal.sCPFCGC1
            lErro = CF_ECF("Abre_Cupom", gobjVenda)
            If lErro <> SUCESSO Then gError 99818
        End If
    
    End If
    
    Exit Sub

Erro_BotaoAbrirGaveta_Click:

    Select Case gErr
                
        Case 99613, 99618, 99619, 109823, 109691, 112074, 99952, 99953, 99901, 99818, 105871, 105887, 105895, 204344, 204345, 210003, 210893
                            
        Case 99810
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_NAO_PREENCHIDO, gErr, Error$)
            
        Case 99889
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_NAO_EXISTENTE, gErr)
            
        Case 112072
            Call Rotina_ErroECF(vbOKOnly, ERRO_VENDEDOR_NAO_PREENCHIDO, gErr, Error$)
            
        Case 210504
            Call Rotina_ErroECF(vbOKOnly, ERRO_DAV_NAO_ALTERADO_DEPOIS_DE_IMPRESSO, gErr)
            
        Case 210995
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_BAIXADO, gErr)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175693)

    End Select
    
    Exit Sub
        
End Sub

Private Sub Alteracoes_CancelamentoOrcamento(objVenda As ClassVenda)

Dim objMovCaixa As ClassMovimentoCaixa
Dim objCheque As ClassChequePre
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
Dim iIndice As Integer
Dim objCarne As ClassCarne
Dim lErro As Long
Dim objAdmMeioPagto As New ClassAdmMeioPagto
Dim lSequencialCaixa As Long
Dim objAliquota As New ClassAliquotaICMS
Dim objItens As ClassItemCupomFiscal
Dim iIndice1 As Integer

    For Each objItens In objVenda.objCupomFiscal.colItens
        For Each objAliquota In gcolAliquotasTotal
            If objItens.dAliquotaICMS = objAliquota.dAliquota Then
                objAliquota.dValorTotalizadoLoja = objAliquota.dValorTotalizadoLoja - ((objItens.dPrecoUnitario * objItens.dQuantidade) * objAliquota.dAliquota)
                Exit For
            End If
        Next
    Next
    
    For iIndice = gcolMovimentosCaixa.Count To 1 Step -1
        Set objMovCaixa = gcolMovimentosCaixa.Item(iIndice)
        If objMovCaixa.lNumIntExt = objVenda.objCupomFiscal.lNumOrcamento Then gcolMovimentosCaixa.Remove (iIndice)
    Next
    
    'Para cada movimento da venda
    For Each objMovCaixa In objVenda.colMovimentosCaixa
        If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_DINHEIRO Then gdSaldoDinheiro = gdSaldoDinheiro - objMovCaixa.dValor
        'Se for de cartao de crédito ou débito especificado
        If (objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAOCREDITO Or objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAODEBITO) And objMovCaixa.iAdmMeioPagto <> 0 Then
            'Busca em gcolCartão a ocorrencia de Cartão nao especificado
            For iIndice = gcolCartao.Count To 1 Step -1
                Set objAdmMeioPagtoCondPagto = gcolCartao.Item(iIndice)
                'Se encontrou
                If objAdmMeioPagtoCondPagto.iAdmMeioPagto = objMovCaixa.iAdmMeioPagto And objAdmMeioPagtoCondPagto.iParcelamento = objMovCaixa.iParcelamento And objAdmMeioPagtoCondPagto.iTipoCartao = objMovCaixa.iTipoCartao Then
                    'Atualiza o saldo do cartão
                    objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                    If objAdmMeioPagtoCondPagto.dSaldo = 0 Then gcolCartao.Remove (iIndice)
                    Exit For
                End If
            Next
        End If
        'Se o omvimento for de cartão de crédito não especificado
        If (objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAOCREDITO) And objMovCaixa.iAdmMeioPagto = 0 Then
            'inclui na col como não especificado
            For Each objAdmMeioPagtoCondPagto In gcolCartao
                If objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = STRING_NAO_DETALHADO_CDEBITO Then
                    'Atualiza o saldo de não especificado
                    objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                    Exit For
                End If
            Next
        End If
        'Se o omvimento for de cartão de débito não especificado
        If (objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAODEBITO) And objMovCaixa.iAdmMeioPagto = 0 Then
            'inclui na col como não especificado
            For Each objAdmMeioPagtoCondPagto In gcolCartao
                If objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = STRING_NAO_DETALHADO_CCREDITO Then
                    'Atualiza o saldo de não especificado
                    objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                    Exit For
                End If
            Next
        End If
    Next
    
    'Para cada movimento
    For iIndice = objVenda.colMovimentosCaixa.Count To 1 Step -1
        'Pega o movimento
        Set objMovCaixa = objVenda.colMovimentosCaixa.Item(iIndice)
        'Se for um recebimento em ticket
        If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_VALETICKET Then
            'Se for não especificado
            If objMovCaixa.iAdmMeioPagto = 0 Then
                'Para cada obj de ticket da coleção global de tickets
                For Each objAdmMeioPagtoCondPagto In gcolTicket
                    'Se for o não especificado
                    If objAdmMeioPagtoCondPagto.iAdmMeioPagto = 0 Then
                        'Atualiza o saldo de não especificado
                        objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                    End If
                Next
            'Se for especificado
            Else
                'Para cada Ticket da coleção global
                For iIndice1 = gcolTicket.Count To 1 Step -1
                    Set objAdmMeioPagtoCondPagto = gcolTicket.Item(iIndice1)
                    'Se encontrou o ticket/parcelamento
                    If objAdmMeioPagtoCondPagto.iAdmMeioPagto = objMovCaixa.iAdmMeioPagto Then
                        'Atualiza o saldo
                        objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                        If objAdmMeioPagtoCondPagto.dSaldo = 0 Then gcolTicket.Remove (iIndice1)
                        'Sinaliza que encontrou
                        Exit For
                    End If
                Next
            End If
        End If
    Next
    
    Set objAdmMeioPagtoCondPagto = New ClassAdmMeioPagtoCondPagto
    
    'Verifica se já existe movimentos de Outros\
    'Para cada MOvimento de Outros
    For iIndice = objVenda.colMovimentosCaixa.Count To 1 Step -1
        'Pega o MOvimento
        Set objMovCaixa = objVenda.colMovimentosCaixa.Item(iIndice)
        'Se for do tipo outros
        If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_OUTROS Then
            'Se for não especificado
            If objMovCaixa.iAdmMeioPagto = 0 Then
                'Para cada pagamento em outros na coleção global
                For Each objAdmMeioPagtoCondPagto In gcolOutros
                    'Se for o não especificado
                    If objAdmMeioPagtoCondPagto.iAdmMeioPagto = 0 Then
                        'Atualiza o saldo não especificado
                        objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                    End If
                Next
            'Se for especificado
            Else
                'Para cada Pagamento em outros na col global
                For iIndice1 = gcolOutros.Count To 1 Step -1
                    Set objAdmMeioPagtoCondPagto = gcolOutros.Item(iIndice1)
                    'Se for do mesmo tipo que o atual
                    If objAdmMeioPagtoCondPagto.iAdmMeioPagto = objMovCaixa.iAdmMeioPagto Then
                        'Atualiza o saldo
                        objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                        If objAdmMeioPagtoCondPagto.dSaldo = 0 Then gcolOutros.Remove (iIndice1)
                        Exit For
                    End If
                Next
            End If
        End If
    Next
        
    'remove o Carne na col global
    If objVenda.objCarne.colParcelas.Count > 0 Then
        For iIndice = 1 To gcolCarne.Count
            Set objCarne = gcolCarne.Item(iIndice)
            If objCarne.lNumIntExt = objVenda.objCupomFiscal.lNumOrcamento Then gcolCarne.Remove (iIndice)
        Next
    End If
    
    'remove o Cheque na col global
    If objVenda.colCheques.Count > 0 Then
        For iIndice = gcolCheque.Count To 1 Step -1
            Set objCheque = gcolCheque.Item(iIndice)
            If objCheque.lNumIntExt = objVenda.objCupomFiscal.lNumOrcamento Then gcolCheque.Remove (iIndice)
        Next
    End If
    
    Exit Sub
    
End Sub

Private Sub Move_Dados_Memoria()

Dim objMovCaixa As ClassMovimentoCaixa
Dim objItens As New ClassItemCupomFiscal
Dim objAliquota As New ClassAliquotaICMS

    Set gobjVenda.colCheques = Nothing
    Set gobjVenda.colTroca = Nothing
    Set gobjVenda.objCarne = Nothing

    Set gobjVenda.colMovimentosCaixa = New Collection
    
    Set objMovCaixa = New ClassMovimentoCaixa
    
    objMovCaixa.dtDataMovimento = Date
    objMovCaixa.dValor = StrParaDbl(Subtotal.Caption)
    objMovCaixa.iFilialEmpresa = giFilialEmpresa
    objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_DINHEIRO
    objMovCaixa.iCaixa = giCodCaixa
    objMovCaixa.iCodOperador = giCodOperador
    objMovCaixa.dHora = CDbl(Time)
    objMovCaixa.iAdmMeioPagto = MEIO_PAGAMENTO_DINHEIRO
                
    gobjVenda.colMovimentosCaixa.Add objMovCaixa
        
    'atualiza o saldo em dinheiro
    gdSaldoDinheiro = gdSaldoDinheiro + StrParaDbl(Subtotal.Caption)
    
    'Joga os outros campos necessários ao obj global
    gobjVenda.objCupomFiscal.dValorTotal = StrParaDbl(Subtotal.Caption)
    gobjVenda.objCupomFiscal.dValorProdutos = gobjVenda.objCupomFiscal.dValorTotal
'    gobjVenda.objCupomFiscal.dtDataEmissao = Date
'    gobjVenda.objCupomFiscal.dHoraEmissao = CDbl(Time)
    gobjVenda.objCupomFiscal.iFilialEmpresa = giFilialEmpresa
    gobjVenda.objCupomFiscal.iCodCaixa = giCodCaixa
    gobjVenda.objCupomFiscal.iECF = giCodECF
    
    For Each objItens In gobjVenda.objCupomFiscal.colItens
        For Each objAliquota In gcolAliquotasTotal
            If objItens.dAliquotaICMS = objAliquota.dAliquota Then
                objAliquota.dValorTotalizadoLoja = objAliquota.dValorTotalizadoLoja + ((objItens.dPrecoUnitario * objItens.dQuantidade) * objAliquota.dAliquota)
                Exit For
            End If
        Next
    Next

'    If OptionDAV.Value = True Then
'        gobjVenda.iTipo = OPTION_DAV
'        gobjVenda.objCupomFiscal.iTipo = OPTION_DAV
'    ElseIf OptionPreVenda.Value = True Then
'        gobjVenda.iTipo = OPTION_PREVENDA
'        gobjVenda.objCupomFiscal.iTipo = OPTION_PREVENDA
'    'modelo nao possui ECF
'    ElseIf giCodModeloECF = 4 Then
'        gobjVenda.objCupomFiscal.iStatus = STATUS_BAIXADO
'        gobjVenda.iTipo = OPTION_CF
'        gobjVenda.objCupomFiscal.iTipo = OPTION_CF
'    Else
        gobjVenda.iTipo = OPTION_CF
        gobjVenda.objCupomFiscal.iTipo = OPTION_CF
'    End If
    
End Sub


Private Sub LeitoraCodBarras_OnComm()

Dim sInput As String
Dim sInput1 As String
Dim iInput As Integer
Dim i As Long
Dim j As Long

On Error GoTo Erro_LeitoraCodBarras_OnComm

    Select Case LeitoraCodBarras.CommEvent
    
        ' Handle each event or error by placing
        ' code below each case statement
        
        ' Errors
    
        Case comEventBreak   ' A Break was received.
            
        Case comEventFrame   ' Framing Error
            MsgBox "Erro de Frame"
            
        Case comEventOverrun   ' Data Lost.
            MsgBox "Dados Perdidos"
      
        Case comEventRxOver   ' Receive buffer overflow.
            MsgBox "Buffer Overflow"
      
        Case comEventRxParity   ' Parity Error.
            MsgBox "Erro de Paridade"
      
        Case comEventTxFull   ' Transmit buffer full.
        
        Case comEventDCB   ' Unexpected error retrieving DCB]
            MsgBox "Erro de DCB"

        ' Events
        
        Case comEvCD   ' Change in the CD line.
      
        Case comEvCTS   ' Change in the CTS line.
      
        Case comEvDSR   ' Change in the DSR line.
      
        Case comEvRing   ' Change in the Ring Indicator.
            
        Case comEvReceive   ' Received RThreshold # of chars.
         
         
                     
'            For i = 1 To 1000000
'               j = j + 1
'            Next

            Call Sleep(1000)
         
            LeitoraCodBarras.InputLen = 0
            sInput = LeitoraCodBarras.Input
            ProdutoNomeRed.Text = left(sInput, Len(sInput) - 1)
            LeitoraCodBarras.InBufferCount = 0
      
            Call ProdutoNomeRed_Validate(False)
      
      
        Case comEvSend   ' There are SThreshold number of characters in the transmit buffer.
      
        Case comEvEOF   ' An EOF charater was found in the input stream
            
    End Select

    Exit Sub
    
Erro_LeitoraCodBarras_OnComm:
    
    Select Case gErr
                
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 175695)

    End Select
    
    Exit Sub

End Sub

Public Sub CGC_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CGC_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CGC, iAlterado)

End Sub

Public Sub CGC_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CGC_Validate
    
    'Se CGC/CPF não foi preenchido -- Exit Sub
    If Len(Trim(CGC.Text)) = 0 Then Exit Sub
    
    Select Case Len(Trim(CGC.Text))

        Case STRING_CPF 'CPF
            
            'Critica Cpf
            lErro = Cpf_Critica(CGC.Text)
            If lErro <> SUCESSO Then gError 199469
            
            'Formata e coloca na Tela
            CGC.Format = "000\.000\.000-00; ; ; "
            CGC.Text = CGC.Text

        Case STRING_CGC 'CGC
            
            'Critica CGC
            lErro = Cgc_Critica(CGC.Text)
            If lErro <> SUCESSO Then gError 199470
            
            'Formata e Coloca na Tela
            CGC.Format = "00\.000\.000\/0000-00; ; ; "
            CGC.Text = CGC.Text

        Case Else
                
            gError 199471

    End Select

    gobjVenda.objCupomFiscal.sCPFCGC1 = CGC.FormattedText
    gobjVenda.objCupomFiscal.sCPFCGC = CGC.Text
    Exit Sub

Erro_CGC_Validate:

    Cancel = True

    Select Case gErr

        Case 199468, 199469

        Case 199471
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC_CPF", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 199472)

    End Select

    Exit Sub

End Sub

Private Sub NomeCliente_Validate(Cancel As Boolean)
    gobjVenda.objCupomFiscal.sNomeCliente = NomeCliente.Text
End Sub

'Private Sub Endereco_Validate(Cancel As Boolean)
'    gobjVenda.objCupomFiscal.sEndereco = Endereco.Text
'End Sub


Private Sub BotaoMesclar_Click()

Dim objOrcamento As New ClassOrcamentoLoja
Dim objVenda As New ClassVenda
Dim colOrcamento As New Collection
Dim iAchou As Integer
Dim lErro As Long
Dim objItens As ClassItemCupomFiscal
Dim iIndice As Integer
Dim objProduto As ClassProduto


On Error GoTo Erro_BotaoMesclar_Click

    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 210011


    'Chama tela de OrçamentoLista
    Call Chama_TelaECF_Modal("OrcamentoLista", objOrcamento)
    
    If giRetornoTela = vbOK Then
        
'        If giPreVenda = 1 And giUsaImpressoraFiscal = 0 Then
'            If Not OptionPreVenda.Value Then OptionPreVenda.Value = True
'        ElseIf giDAV = 1 Then
'            If Not OptionDAV.Value Then OptionDAV.Value = True
'        End If
        
        objVenda.objCupomFiscal.lNumOrcamento = objOrcamento.lNumOrcamento
        
        'Função Que le os orcamentos
        lErro = CF_ECF("OrcamentoECF_Le", objVenda)
        If lErro <> SUCESSO And lErro <> 204690 And lErro <> 210447 Then gError 105857
        
        If lErro = 210447 Then gError 210452
        
        'orcamento nao cadastrado
        If lErro <> SUCESSO Then gError 105858
        
        'descobre o nome reduzido do produto
        For Each objItens In objVenda.objCupomFiscal.colItens
            For iIndice = 1 To gaobjProdutosNome.Count
                Set objProduto = gaobjProdutosNome.Item(iIndice)
                If objItens.sProduto = objProduto.sCodigo Then
                    objItens.sProdutoNomeRed = objProduto.sNomeReduzido
                    Exit For
                End If
            Next
        Next
        
        'Traz ele para a tela
        Set gobjVenda = New ClassVenda
        Call Copia_Venda(gobjVenda, objVenda)
        Call Traz_Orcamento
        
        'se o cupom fiscal estiver ligado, cham OptionCF_Click para transformar o orcamento em cupom
        If OptionCF.Value Then Call OptionCF_Click
        
    End If
    
    Exit Sub

Erro_BotaoMesclar_Click:

    Select Case gErr

        Case 105857, 210011

        Case 105858
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_NAO_CADASTRADO1, gErr, objOrcamento.lNumOrcamento)

        Case 210452
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_BAIXADO, gErr, objVenda.objCupomFiscal.lNumOrcamento)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175673)

    End Select

    Exit Sub

End Sub

Public Sub Form_Resize(iLargura As Integer, iAltura As Integer)


Dim iIndice As Integer
Dim lGridTopOrig As Long

    If Parent.WindowState = 2 Then

        UserControl.Size iLargura, iAltura
        Frame3.Width = iLargura
    
'        ListCF.Width = giLarguraListCF + (iLargura - giLarguraOrig)
'        ListCF.Height = giAlturaListCF + (iAltura - giAlturaOrig)
    
        
    
        Figura.Height = IIf(giAlturaFigura + (iAltura - giAlturaOrig) < 0, 1, giAlturaFigura + (iAltura - giAlturaOrig))
        
        
        Grid.Width = giLarguraGrid + (iLargura - giLarguraOrig)
'        Grid.Height = giAlturaGrid + (iAltura - giAlturaOrig)
        
        Grid1.left = Grid.left
        Grid1.Width = Grid.Width
        ListCF.Width = Grid.Width
        Picture1.Width = Grid1.Width
        
        
        If Grid.Width < 12000 Then
            Grid.Font.Size = 8
            Grid1.Font.Size = 8
            ListCF.Font.Size = 8
            Picture1.Font.Size = 8
        Else
            Grid.Font.Size = 14
            Grid1.Font.Size = 14
            ListCF.Font.Size = 14
            Picture1.Font.Size = 14
        End If
        
        Grid1.Visible = True
        
        If Grid.Width < 8000 Then
            Grid.Visible = False
            ListCF.Visible = True
            Grid1.BorderStyle = flexBorderSingle
            Grid1.Height = 13 * Grid1.RowHeight(0)
            
        Else
            Grid.Visible = True
            ListCF.Visible = False
            Grid1.BorderStyle = flexBorderNone
            Grid1.Height = 11 * Grid1.RowHeight(0)
        
        End If
        
        
        lGridTopOrig = Grid.top
        Grid.top = Grid1.top + Grid1.Height
        ListCF.top = Grid1.top + Grid1.Height - 15
        ListCF.left = Grid1.left
        
        Grid.Height = iAltura - GetSystemMetrics(SM_CYCAPTION) * Screen.TwipsPerPixelX - Grid.top
        ListCF.Height = iAltura - ListCF.top
        
        'Limpa toda a list
        ListCF.Clear
        
        Call Inicia_Cupom_Tela

        
    
    End If
    

End Sub


Private Function ControleMaxTamVisivel(ByVal objControle As Object, sCaracter As String) As Integer
'Retorna o maximo de caracteres que cabem na largura de uma picturebox ou outro controle que tenha hdc
'o fonte utilizado tem que ser de caracteres com a mesma largura

Dim hdc As Long, sTexto As String, szTexto As typeSize, iTam As Integer

    iTam = 0
    
    hdc = objControle.hdc
    If hdc <> 0 Then
        
'        sTexto = "1234567890123456789012345678901234567890"
        sTexto = String(100, sCaracter)
        
        If GetTextExtentPoint32(objControle.hdc, sTexto, Len(sTexto), szTexto) <> 0 Then
            
'            iTam = Fix(objControle.Width * 40 / (szTexto.cx * Screen.TwipsPerPixelX))
            iTam = Fix(objControle.Width * 100 / (szTexto.cx * (Screen.TwipsPerPixelX)))
            
        End If
        
    End If
    
    ControleMaxTamVisivel = iTam
    
End Function


Private Function Proxima_Linha_Grid()

Dim iLinha As Integer

        giUltimaLinhaGrid = giUltimaLinhaGrid + 1
        If giUltimaLinhaGrid + 1 > Grid.Rows Then Grid.Rows = giUltimaLinhaGrid + 1
        
        If giUltimaLinhaGrid + 1 > giLinhasVisiveisGrid Then
            Grid.TopRow = giUltimaLinhaGrid - giLinhasVisiveisGrid + 3
        End If

End Function

Private Sub Move_Dados_Memoria_1()

    gobjVenda.objCupomFiscal.dValorTotal = StrParaDbl(Subtotal.Caption)
    gobjVenda.objCupomFiscal.dValorProdutos = gobjVenda.objCupomFiscal.dValorTotal
    gobjVenda.objCupomFiscal.iFilialEmpresa = giFilialEmpresa
    gobjVenda.objCupomFiscal.iCodCaixa = giCodCaixa
    gobjVenda.objCupomFiscal.iECF = giCodECF

End Sub

Private Sub ItemCupom_AjustaTrib(objItem As ClassItemCupomFiscal)
'Ajusta a tributacao de acordo com o que foi efetivamente vendido
'??? falta tratar ST

Dim objTributacaoDocItem As ClassTributacaoDocItem
Dim dValorLiquido As Double

    Set objTributacaoDocItem = objItem.objTributacaoDocItem

    dValorLiquido = Arredonda_Moeda(Arredonda_Moeda(objItem.dQuantidade * objItem.dPrecoUnitario) - Arredonda_Moeda(objItem.dValorDesconto))
    
    'dados gerais
    objTributacaoDocItem.dQuantidade = objItem.dQuantidade
    objTributacaoDocItem.dPrecoUnitario = objItem.dPrecoUnitario
    
    If objTributacaoDocItem.dTotTrib <> 0 Then
        objTributacaoDocItem.dTotTrib = Arredonda_Moeda(objTributacaoDocItem.dTotTrib * dValorLiquido / (objTributacaoDocItem.dQuantidade * objTributacaoDocItem.dPrecoUnitario))
    End If
    
    'icms
    If objTributacaoDocItem.dICMSAliquota <> 0 Then
        objTributacaoDocItem.dICMSBase = dValorLiquido
        objTributacaoDocItem.dICMSValor = Arredonda_Moeda(objTributacaoDocItem.dICMSBase * objTributacaoDocItem.dICMSAliquota)
    Else
        objTributacaoDocItem.dICMSBase = 0
        objTributacaoDocItem.dICMSValor = 0
    End If
    
    'pis
    If objTributacaoDocItem.dPISAliquota <> 0 Then
        objTributacaoDocItem.dPISBase = dValorLiquido
        objTributacaoDocItem.dPISValor = Arredonda_Moeda(objTributacaoDocItem.dPISBase * objTributacaoDocItem.dPISAliquota)
    Else
        If objTributacaoDocItem.dPISAliquotaValor <> 0 Then
            objTributacaoDocItem.dPISBase = 0 '??? rever
            objTributacaoDocItem.dPISQtde = objItem.dQuantidade  '??? rever pq pode ter que converter unidade de medida
            objTributacaoDocItem.dPISValor = Arredonda_Moeda(objTributacaoDocItem.dPISQtde * objTributacaoDocItem.dPISAliquota)
        Else
            objTributacaoDocItem.dPISBase = 0
            objTributacaoDocItem.dPISValor = 0
    
        End If
    End If
    
    'cofins
    If objTributacaoDocItem.dCOFINSAliquota <> 0 Then
        objTributacaoDocItem.dCOFINSBase = dValorLiquido
        objTributacaoDocItem.dCOFINSValor = Arredonda_Moeda(objTributacaoDocItem.dCOFINSBase * objTributacaoDocItem.dCOFINSAliquota)
    Else
        If objTributacaoDocItem.dCOFINSAliquotaValor <> 0 Then
            objTributacaoDocItem.dCOFINSQtde = objItem.dQuantidade  '??? rever pq pode ter que converter unidade de medida
            objTributacaoDocItem.dCOFINSValor = Arredonda_Moeda(objTributacaoDocItem.dCOFINSQtde * objTributacaoDocItem.dCOFINSAliquota)
        Else
            objTributacaoDocItem.dCOFINSBase = 0
            objTributacaoDocItem.dCOFINSValor = 0
    
        End If
    End If
    
    If objTributacaoDocItem.dISSAliquota <> 0 Then
        objTributacaoDocItem.dISSBase = dValorLiquido
        objTributacaoDocItem.dISSValor = Arredonda_Moeda(objTributacaoDocItem.dISSBase * objTributacaoDocItem.dISSAliquota)
    Else
        objTributacaoDocItem.dISSBase = 0
        objTributacaoDocItem.dISSValor = 0
    End If
    
End Sub


