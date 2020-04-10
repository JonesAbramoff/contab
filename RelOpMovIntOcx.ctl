VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl RelOpMovIntOcx 
   ClientHeight    =   5535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8895
   KeyPreview      =   -1  'True
   ScaleHeight     =   5535
   ScaleWidth      =   8895
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4140
      Index           =   2
      Left            =   165
      TabIndex        =   10
      Top             =   1140
      Visible         =   0   'False
      Width           =   8565
      Begin VB.Frame FrameData 
         Caption         =   "Data"
         Height          =   885
         Left            =   210
         TabIndex        =   40
         Top             =   2010
         Width           =   5700
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   300
            Left            =   2040
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   345
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataInicial 
            Height          =   300
            Left            =   1035
            TabIndex        =   13
            Top             =   345
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDown2 
            Height          =   300
            Left            =   4680
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   345
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataFinal 
            Height          =   300
            Left            =   3675
            TabIndex        =   14
            Top             =   345
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label dIni 
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
            Height          =   240
            Left            =   645
            TabIndex        =   44
            Top             =   390
            Width           =   345
         End
         Begin VB.Label dFim 
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
            Left            =   3270
            TabIndex        =   43
            Top             =   405
            Width           =   360
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Centro de Custo"
         Height          =   1335
         Left            =   210
         TabIndex        =   35
         Top             =   300
         Width           =   5700
         Begin MSMask.MaskEdBox CclInicial 
            Height          =   285
            Left            =   750
            TabIndex        =   11
            Top             =   375
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   503
            _Version        =   393216
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
         Begin MSMask.MaskEdBox CclFinal 
            Height          =   270
            Left            =   750
            TabIndex        =   12
            Top             =   855
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   476
            _Version        =   393216
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin VB.Label DescCclInic 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2145
            TabIndex        =   39
            Top             =   375
            Width           =   3255
         End
         Begin VB.Label DescCclFim 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2145
            TabIndex        =   38
            Top             =   855
            Width           =   3255
         End
         Begin VB.Label LabelCclAte 
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
            Left            =   345
            TabIndex        =   37
            Top             =   900
            Width           =   360
         End
         Begin VB.Label LabelCclDe 
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
            Left            =   390
            TabIndex        =   36
            Top             =   390
            Width           =   315
         End
      End
      Begin MSComctlLib.TreeView TvwCcls 
         Height          =   3540
         Left            =   6075
         TabIndex        =   15
         Top             =   390
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   6244
         _Version        =   393217
         Indentation     =   453
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin VB.Label LabelCcl 
         Caption         =   "Centro de Custo"
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
         Left            =   6075
         TabIndex        =   45
         Top             =   135
         Width           =   1530
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4140
      Index           =   1
      Left            =   165
      TabIndex        =   1
      Top             =   1140
      Width           =   8565
      Begin VB.Frame Frame3 
         Caption         =   "Categoria de Produtos"
         Height          =   1425
         Left            =   225
         TabIndex        =   30
         Top             =   2475
         Width           =   7125
         Begin VB.ComboBox ValorFinal 
            Height          =   315
            Left            =   3630
            TabIndex        =   9
            Top             =   975
            Width           =   2100
         End
         Begin VB.CheckBox TodasCategorias 
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
            Left            =   285
            TabIndex        =   6
            Top             =   255
            Width           =   855
         End
         Begin VB.ComboBox ValorInicial 
            Height          =   315
            Left            =   720
            TabIndex        =   8
            Top             =   975
            Width           =   1950
         End
         Begin VB.ComboBox Categoria 
            Height          =   315
            Left            =   1635
            TabIndex        =   7
            Top             =   525
            Width           =   2745
         End
         Begin VB.Label Label7 
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
            Left            =   675
            TabIndex        =   34
            Top             =   585
            Width           =   855
         End
         Begin VB.Label Label8 
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
            TabIndex        =   33
            Top             =   1020
            Width           =   420
         End
         Begin VB.Label Label6 
            Caption         =   "Ate:"
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
            Left            =   3210
            TabIndex        =   32
            Top             =   1020
            Width           =   555
         End
         Begin VB.Label Label5 
            Caption         =   "Label5"
            Height          =   15
            Left            =   360
            TabIndex        =   31
            Top             =   720
            Width           =   30
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Produtos"
         Height          =   1200
         Left            =   225
         TabIndex        =   25
         Top             =   1110
         Width           =   7095
         Begin MSMask.MaskEdBox ProdutoFinal 
            Height          =   315
            Left            =   750
            TabIndex        =   5
            Top             =   750
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdutoInicial 
            Height          =   315
            Left            =   750
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
            TabIndex        =   29
            Top             =   765
            Width           =   555
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
            Left            =   360
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   28
            Top             =   405
            Width           =   615
         End
         Begin VB.Label DescProdFim 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2295
            TabIndex        =   27
            Top             =   750
            Width           =   3135
         End
         Begin VB.Label DescProdInic 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2295
            TabIndex        =   26
            Top             =   375
            Width           =   3135
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Ordem de Produção"
         Height          =   810
         Left            =   225
         TabIndex        =   16
         Top             =   150
         Width           =   7080
         Begin VB.TextBox OpInicial 
            Height          =   300
            Left            =   765
            TabIndex        =   2
            Top             =   360
            Width           =   1680
         End
         Begin VB.TextBox OpFinal 
            Height          =   300
            Left            =   3645
            TabIndex        =   3
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label LabelOpInicial 
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
            Left            =   390
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   24
            Top             =   390
            Width           =   420
         End
         Begin VB.Label LabelOpFinal 
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
            Left            =   3225
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   23
            Top             =   390
            Width           =   570
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6600
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpMovIntOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpMovIntOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpMovIntOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpMovIntOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   22
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
      Left            =   4755
      Picture         =   "RelOpMovIntOcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpMovIntOcx.ctx":0A96
      Left            =   1425
      List            =   "RelOpMovIntOcx.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   180
      Width           =   2916
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4620
      Left            =   120
      TabIndex        =   46
      Top             =   705
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   8149
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Parte 1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Parte 2"
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
      Left            =   720
      TabIndex        =   47
      Top             =   225
      Width           =   615
   End
End
Attribute VB_Name = "RelOpMovIntOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoOp As AdmEvento
Attribute objEventoOp.VB_VarHelpID = -1
Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1


Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio
Dim giProdInicial As Integer
Dim giCclInicial As Integer
Dim giOp_Inicial As Integer
Dim iFrameAtual As Integer

Private Sub Form_Load()

Dim lErro As Long
Dim colCategoriaProduto As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim sMascaraCclPadrao  As String

On Error GoTo Erro_Form_Load
    
    iFrameAtual = 1
    
    Set objEventoOp = New AdmEvento
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento

    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoInicial)
    If lErro <> SUCESSO Then Error 34360

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoFinal)
    If lErro <> SUCESSO Then Error 34361
    
    'Inicializa Máscara de Ccl
    sMascaraCclPadrao = String(STRING_CCL, 0)

    lErro = MascaraCcl(sMascaraCclPadrao)
    If lErro <> SUCESSO Then Error 52179

    CclInicial.Mask = sMascaraCclPadrao
    
    CclFinal.Mask = sMascaraCclPadrao

   'Inicializa a arvore de Centros de Custo
    lErro = Carga_Arvore_Ccl(TvwCcls.Nodes)
    If lErro <> SUCESSO Then Error 34363
    
    'Le as categorias de produto
    lErro = CF("CategoriasProduto_Le_Todas", colCategoriaProduto)
    If lErro <> SUCESSO And lErro <> 22542 Then Error 34364

    'Preenche CategoriaProduto
    For Each objCategoriaProduto In colCategoriaProduto

        Categoria.AddItem objCategoriaProduto.sCategoria

    Next
        
    Call Define_Padrao

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = Err

    Select Case Err

        Case 34360, 34361, 34363, 34364
        
        Case 52179
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170064)

    End Select

    Exit Sub

End Sub

Sub Define_Padrao()
'Preenche a tela com as opções padrão de FilialEmpresa

Dim lErro As Long

On Error GoTo Erro_Define_Padrao

    giProdInicial = 1
    
    TodasCategorias_Click
    
    TodasCategorias = 1
    
    giCclInicial = 1
    
    giOp_Inicial = 1
    
    DataInicial.PromptInclude = False
    DataInicial.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataInicial.PromptInclude = True
    
    DataFinal.PromptInclude = False
    DataFinal.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataFinal.PromptInclude = True
       
    Exit Sub

Erro_Define_Padrao:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170065)

    End Select

    Exit Sub

End Sub

Private Sub DataFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFinal)

End Sub

Private Sub DataInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicial)

End Sub

Private Sub ProdutoInicial_GotFocus()
'Mostra a arvore de produtos

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_GotFocus

    giProdInicial = 1

    Exit Sub

Erro_ProdutoInicial_GotFocus:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170066)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoFinal_GotFocus()
'Mostra a arvore de produtos

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_GotFocus

    giProdInicial = 0

    Exit Sub

Erro_ProdutoFinal_GotFocus:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170067)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim sDescCcl As String

On Error GoTo Erro_PreencherParametrosNaTela

 Call Limpar_Tela

    lErro = objRelOpcoes.Carregar
    If lErro Then Error 34367

    'pega Ordem de Producao Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TOPINIC", sParam)
    If lErro Then Error 34368

    OpInicial.Text = sParam

    'pega Ordem de Producao Final e exibe
    lErro = objRelOpcoes.ObterParametro("TOPFIM", sParam)
    If lErro Then Error 34369

    OpFinal.Text = sParam
    
    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro Then Error 34370

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then Error 34371

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro Then Error 34372

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then Error 34373
    
    'pega parâmetro TodasCategorias e exibe
    lErro = objRelOpcoes.ObterParametro("NTODASCAT", sParam)
    If lErro Then Error 34374

    TodasCategorias.Value = CInt(sParam)

    'pega parâmetro categoria de produto e exibe
    lErro = objRelOpcoes.ObterParametro("TCATPROD", sParam)
    If lErro Then Error 34375
    
    Categoria.Text = sParam

    'pega parâmetro valor inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TITEMCATPRODINI", sParam)
    If lErro Then Error 34376
    
    ValorInicial.Text = sParam
    
    'pega parâmetro Valor Final e exibe
    lErro = objRelOpcoes.ObterParametro("TITEMCATPRODFIM", sParam)
    If lErro Then Error 34377
    
    ValorFinal.Text = sParam
 
    'pega Ccl Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TCCLINIC", sParam)
    If lErro Then Error 34378

    If sParam <> "" Then
    
        lErro = Obtem_Descricao_Ccl(sParam, sDescCcl)
        If lErro <> SUCESSO Then Error 34379

        CclInicial.PromptInclude = False
        CclInicial.Text = sParam
        CclInicial.PromptInclude = True
        
        DescCclInic.Caption = sDescCcl

    End If
    
    'pega Ccl Final e exibe
    lErro = objRelOpcoes.ObterParametro("TCCLFIM", sParam)
    If lErro Then Error 34380
    
    If sParam <> "" Then
    
        lErro = Obtem_Descricao_Ccl(sParam, sDescCcl)
        If lErro <> SUCESSO Then Error 34381
    
        CclFinal.PromptInclude = False
        CclFinal.Text = sParam
        CclFinal.PromptInclude = True
        
        DescCclFim.Caption = sDescCcl
    
    End If
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then Error 34382

    If sParam <> "07/09/1822" Then
        
        DataInicial.PromptInclude = False
        DataInicial.Text = sParam
        DataInicial.PromptInclude = True

    End If
    
    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then Error 34383

    If sParam <> "07/09/1822" Then
        
        DataFinal.PromptInclude = False
        DataFinal.Text = sParam
        DataFinal.PromptInclude = True
        
    End If
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 34367 To 34383

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170068)

    End Select

    Exit Function

End Function

Private Sub Form_Unload(Cancel As Integer)

    Set objEventoOp = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    
End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82397

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82398

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 82399

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 82397, 82399

        Case 82398
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170069)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82445

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82446

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 82447

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 82445, 82447

        Case 82446
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170070)

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
        If lErro <> SUCESSO Then gError 82493

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 82493

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170071)

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
        If lErro <> SUCESSO Then gError 82492

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 82492

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170072)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 29894
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 34358
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 34358, 34359
        
        Case 29894
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170073)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub


Sub Limpar_Tela()

    Call Limpa_Tela(Me)

    DescProdInic.Caption = ""
    DescProdFim.Caption = ""
    DescCclInic.Caption = ""
    DescCclFim.Caption = ""
    
    ComboOpcoes.SetFocus

End Sub

Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String, sCcl_I As String, sCcl_F As String) As Long
'Formata os produtos retornando em sProd_I e sProd_F
'Verifica se os parâmetros iniciais são maiores que os finais

Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim iCclPreenchida_I As Integer
Dim iCclPreenchida_F As Integer
Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then Error 34385

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then Error 34386

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then Error 34387

    End If

   'data inicial não pode ser maior que a data final
    If Trim(DataInicial.ClipText) <> "" And Trim(DataFinal.ClipText) <> "" Then
    
         If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then Error 34388
    
    End If
    
    'valor inicial não pode ser maior que o valor final
    If Trim(ValorInicial.Text) <> "" And Trim(ValorFinal.Text) <> "" Then
    
         If ValorInicial.Text > ValorFinal.Text Then Error 34389
         
    Else
        
        If Trim(ValorInicial.Text) = "" And Trim(ValorFinal.Text) = "" And TodasCategorias.Value = 0 Then Error 37999
    
    End If
    
    'verifica se o Ccl Inicial é maior que o Ccl Final
    lErro = CF("Ccl_Formata", CclInicial.Text, sCcl_I, iCclPreenchida_I)
    If lErro Then Error 34390

    lErro = CF("Ccl_Formata", CclFinal.Text, sCcl_F, iCclPreenchida_F)
    If lErro Then Error 34391

    If (iCclPreenchida_I = CCL_PREENCHIDA) And (iCclPreenchida_F = CCL_PREENCHIDA) Then

        If sCcl_I > sCcl_F Then Error 34392

    End If
    
    'ordem de produção inicial não pode ser maior que a final
    If Trim(OpInicial.Text) <> "" And Trim(OpFinal.Text) <> "" Then

        If OpInicial.Text > OpFinal.Text Then Error 34393

    End If
          
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = Err

    Select Case Err

        Case 34385
            ProdutoInicial.SetFocus

        Case 34386
            ProdutoFinal.SetFocus

        Case 34387
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", Err)
            ProdutoInicial.SetFocus
             
        Case 34388
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)
            DataInicial.SetFocus
            
        Case 34389
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_INICIAL_MAIOR", Err)
            ValorInicial.SetFocus
            
        Case 34390
            CclInicial.SetFocus

        Case 34391
            CclFinal.SetFocus

        Case 34392
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_INICIAL_MAIOR", Err)
            CclInicial.SetFocus

      
        Case 34393
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OP_INICIAL_MAIOR", Err)
            OpInicial.SetFocus
            
        Case 37999
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_NAO_INFORMADO", Err)
            ValorInicial.SetFocus
    
      
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170074)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

    ComboOpcoes.Text = ""
    Limpar_Tela
    Call Define_Padrao

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame2(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame2(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index

    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170075)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sProd_I As String
Dim sProd_F As String
Dim sCcl_I As String
Dim sCcl_F As String
Dim sOp_I As String
Dim sOp_F As String

On Error GoTo Erro_PreencherRelOp

    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)
    sCcl_I = String(STRING_CCL, 0)
    sCcl_F = String(STRING_CCL, 0)
    
    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F, sCcl_I, sCcl_F)
    If lErro <> SUCESSO Then Error 34398

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 34399
         
    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then Error 34400

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then Error 34401
    
    If DataInicial.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINIC", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 34402

    If DataFinal.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 34403
    
    lErro = objRelOpcoes.IncluirParametro("NTODASCAT", CStr(TodasCategorias.Value))
    If lErro <> AD_BOOL_TRUE Then Error 34404
    
    lErro = objRelOpcoes.IncluirParametro("TCATPROD", Categoria.Text)
    If lErro <> AD_BOOL_TRUE Then Error 34405
    
    lErro = objRelOpcoes.IncluirParametro("TITEMCATPRODINI", ValorInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 34406
    
    lErro = objRelOpcoes.IncluirParametro("TITEMCATPRODFIM", ValorFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 34407
    
    lErro = objRelOpcoes.IncluirParametro("TCCLINIC", sCcl_I)
    If lErro <> AD_BOOL_TRUE Then Error 34408

    lErro = objRelOpcoes.IncluirParametro("TCCLFIM", sCcl_F)
    If lErro <> AD_BOOL_TRUE Then Error 34409

    sOp_I = OpInicial.Text
    lErro = objRelOpcoes.IncluirParametro("TOPINIC", sOp_I)
    If lErro <> AD_BOOL_TRUE Then Error 34410

    sOp_F = OpFinal.Text
    lErro = objRelOpcoes.IncluirParametro("TOPFIM", sOp_F)
    If lErro <> AD_BOOL_TRUE Then Error 34411
    
    If TodasCategorias.Value = 0 Then
        
        gobjRelatorio.sNomeTsk = "MovIntCa"
    Else
        gobjRelatorio.sNomeTsk = "MovInt"
    End If

   
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd_I, sProd_F, sCcl_I, sCcl_F, sOp_I, sOp_F)
    If lErro <> SUCESSO Then Error 34412

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 34398 To 34412

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170076)

    End Select

    Exit Function

End Function


Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 34413

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 34414

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Call BotaoLimpar_Click

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 34413
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 34414

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170077)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 34415

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 34415

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170078)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 34416

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then Error 34417

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 34418

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 34416
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 34417, 34418

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170079)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate

    giProdInicial = 0

    lErro = CF("Produto_Perde_Foco", ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO And lErro <> 27095 Then Error 34419
    
    If lErro <> SUCESSO Then Error 43263

    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True


    Select Case Err

        Case 34419

         Case 43263
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170080)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    giProdInicial = 1

    lErro = CF("Produto_Perde_Foco", ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then Error 34420
    
    If lErro <> SUCESSO Then Error 43264

    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True


    Select Case Err

        Case 34420

         Case 43264
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170081)

    End Select

    Exit Sub

End Sub


Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd_I As String, sProd_F As String, sCcl_I As String, sCcl_F As String, sOp_I As String, sOp_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long


On Error GoTo Erro_Monta_Expressao_Selecao

   If sOp_I <> "" Then sExpressao = "OrdemProducao >= " & Forprint_ConvTexto(sOp_I)

    If sOp_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "OrdemProducao <= " & Forprint_ConvTexto(sOp_F)

    End If


    If sProd_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Produto >= " & Forprint_ConvTexto(sProd_I)

    End If

    If sProd_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Produto <= " & Forprint_ConvTexto(sProd_F)

    End If
    
     If sCcl_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Ccl >= " & Forprint_ConvTexto(sCcl_I)

    End If

    If sCcl_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Ccl <= " & Forprint_ConvTexto(sCcl_F)

    End If
    
    If TodasCategorias.Value = 0 Then
           
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CategoriaProduto = " & Forprint_ConvTexto(Categoria.Text)
            
        If ValorInicial.Text <> "" Then

            If sExpressao <> "" Then sExpressao = sExpressao & " E "
            sExpressao = sExpressao & "ItemCategoriaProduto  >= " & Forprint_ConvTexto(ValorInicial.Text)

        End If
        
        If ValorFinal.Text <> "" Then

            If sExpressao <> "" Then sExpressao = sExpressao & " E "
            sExpressao = sExpressao & "ItemCategoriaProduto <= " & Forprint_ConvTexto(ValorFinal.Text)

        End If
        
    End If
    
    If Trim(DataInicial.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(CDate(DataInicial.Text))

    End If
    
    If Trim(DataFinal.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(CDate(DataFinal.Text))

    End If
 
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If


    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170082)

    End Select

    Exit Function

End Function

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(DataFinal.ClipText) > 0 Then

        sDataFim = DataFinal.Text
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then Error 34421

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True


    Select Case Err

        Case 34421

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170083)

    End Select

    Exit Sub

End Sub


Private Sub DataInicial_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    If Len(DataInicial.ClipText) > 0 Then

        sDataInic = DataInicial.Text
        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then Error 34422

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True


    Select Case Err

        Case 34422

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170084)

    End Select

    Exit Sub

End Sub


Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 34423

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 34423
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170085)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 34424

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 34424
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170086)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 34425

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case Err

        Case 34425
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170087)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 34426

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case Err

        Case 34426
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170088)

    End Select

    Exit Sub

End Sub


Private Sub Categoria_GotFocus()

    If TodasCategorias.Value = 1 Then TodasCategorias.Value = 0
       
End Sub

Private Sub Categoria_Validate(Cancel As Boolean)

    Categoria_Click
 
End Sub


Private Sub Categoria_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As New ClassCategoriaProdutoItem
Dim colCategoria As New Collection

On Error GoTo Erro_Categoria_Click

    If Len(Trim(Categoria.Text)) > 0 Then

        ValorInicial.Clear
        ValorFinal.Clear
        
        'Preenche o objeto com a Categoria
        objCategoriaProduto.sCategoria = Categoria.Text

        'Lê Categoria De Produto no BD
        lErro = CF("CategoriaProduto_Le", objCategoriaProduto)
        If lErro <> SUCESSO And lErro <> 22540 Then Error 34427

        If lErro <> SUCESSO Then Error 34428 'Categoria não está cadastrada

        'Lê os dados de itens de categorias de produto
        lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colCategoria)
        If lErro <> SUCESSO Then Error 34429

        'Preenche Valor Inicial e final
        For Each objCategoriaProdutoItem In colCategoria

            ValorInicial.AddItem (objCategoriaProdutoItem.sItem)
            ValorFinal.AddItem (objCategoriaProdutoItem.sItem)

        Next

    Else
    
        ValorInicial.Text = ""
        ValorFinal.Text = ""
        ValorInicial.Clear
        ValorFinal.Clear

    End If

    Exit Sub

Erro_Categoria_Click:

    Select Case Err

        Case 34427
            Categoria.SetFocus
            
        Case 34428
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_INEXISTENTE", Err)
            Categoria.SetFocus
            
        Case 34429

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170089)

    End Select

    Exit Sub

End Sub


Private Sub ValorInicial_GotFocus()

    If TodasCategorias.Value = 1 Then TodasCategorias.Value = 0
    
End Sub

Private Sub ValorFinal_GotFocus()

    If TodasCategorias.Value = 1 Then TodasCategorias.Value = 0
    
End Sub

Private Sub ValorInicial_Validate(Cancel As Boolean)

    ValorInicial_Click

End Sub

Private Sub ValorFinal_Validate(Cancel As Boolean)

    ValorFinal_Click

End Sub

Private Sub TodasCategorias_Click()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_TodasCategorias_Click

    'Limpa campos
    Categoria.Text = ""
    ValorInicial.Text = ""
    ValorFinal.Text = ""
    
    Exit Sub

Erro_TodasCategorias_Click:

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170090)

    End Select

    Exit Sub

End Sub

Private Sub ValorInicial_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As New ClassCategoriaProdutoItem
Dim colItens As New Collection

On Error GoTo Erro_ValorInicial_Click

    If Len(Trim(ValorInicial.Text)) > 0 Then

        'Tenta selecionar na combo
        lErro = Combo_Item_Igual(ValorInicial)
        If lErro <> SUCESSO Then

            'Preenche o objeto com a Categoria
            objCategoriaProdutoItem.sCategoria = Categoria.Text
            objCategoriaProdutoItem.sItem = ValorInicial.Text

            'Lê Categoria De Produto no BD
            lErro = CF("CategoriaProduto_Le_Item", objCategoriaProdutoItem)
            If lErro <> SUCESSO And lErro <> 22603 Then Error 34430

            'Item da Categoria não está cadastrado
            If lErro <> SUCESSO Then Error 34431
            
        End If

    End If

    Exit Sub

Erro_ValorInicial_Click:

    Select Case Err

        Case 34430
            ValorInicial.SetFocus

        Case 34431
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_INEXISTENTE", Err, objCategoriaProdutoItem.sItem, objCategoriaProdutoItem.sCategoria)
            ValorInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170091)

    End Select

    Exit Sub

End Sub

Private Sub ValorFinal_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As New ClassCategoriaProdutoItem
Dim colItens As New Collection

On Error GoTo Erro_ValorFinal_Click

    If Len(Trim(ValorFinal.Text)) > 0 Then

        'Tenta selecionar na combo
        lErro = Combo_Item_Igual(ValorFinal)
        If lErro <> SUCESSO Then

            'Preenche o objeto com a Categoria
            objCategoriaProdutoItem.sCategoria = Categoria.Text
            objCategoriaProdutoItem.sItem = ValorFinal.Text

            'Lê Categoria De Produto no BD
            lErro = CF("CategoriaProduto_Le_Item", objCategoriaProdutoItem)
            If lErro <> SUCESSO And lErro <> 22603 Then Error 34432
            
            'Item da Categoria não está cadastrado
            If lErro <> SUCESSO Then Error 34433
        End If

    End If

    Exit Sub

Erro_ValorFinal_Click:

    Select Case Err

        Case 34432
            ValorFinal.SetFocus

        Case 34433
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_INEXISTENTE", Err, objCategoriaProdutoItem.sItem, objCategoriaProdutoItem.sCategoria)
            ValorFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170092)

    End Select

    Exit Sub

End Sub


Private Sub CclInicial_GotFocus()
'Mostra a arvore de Ccl

Dim lErro As Long

On Error GoTo Erro_CclInicial_GotFocus

    giCclInicial = 1

    Exit Sub

Erro_CclInicial_GotFocus:

    Select Case Err

        Case 34434

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170093)

    End Select

    Exit Sub

End Sub

Private Sub CclFinal_GotFocus()
'mostra a arvore de Ccl

Dim lErro As Long

On Error GoTo Erro_CclFinal_GotFocus

    giCclInicial = 0

    Exit Sub

Erro_CclFinal_GotFocus:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170094)

    End Select

    Exit Sub

End Sub

Private Sub LabelOpInicial_Click()

Dim lErro As Long
Dim objOp As ClassOrdemDeProducao
Dim colSelecao As Collection


On Error GoTo Erro_LabelOpInicial_Click

    giOp_Inicial = 1

    If Len(Trim(OpInicial.Text)) <> 0 Then

        Set objOp = New ClassOrdemDeProducao
        objOp.sCodigo = OpInicial.Text

    End If

    Call Chama_Tela("OrdProdTodasListaModal", colSelecao, objOp, objEventoOp)
    
    Exit Sub

Erro_LabelOpInicial_Click:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170095)

    End Select

    Exit Sub

End Sub

Private Sub LabelOpFinal_Click()

Dim lErro As Long
Dim colSelecao As Collection
Dim objOp As ClassOrdemDeProducao

On Error GoTo Erro_LabelOpFinal_Click

    giOp_Inicial = 0

    If Len(Trim(OpFinal.Text)) <> 0 Then

        Set objOp = New ClassOrdemDeProducao
        objOp.sCodigo = OpFinal.Text

    End If

    Call Chama_Tela("OrdProdTodasListaModal", colSelecao, objOp, objEventoOp)
   
   Exit Sub

Erro_LabelOpFinal_Click:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170096)

    End Select

    Exit Sub

End Sub


Private Sub objEventoOp_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOp As New ClassOrdemDeProducao

On Error GoTo Erro_objEventoOp_evSelecao

    Set objOp = obj1

    If giOp_Inicial = 1 Then

        OpInicial.Text = objOp.sCodigo
        
    Else

        OpFinal.Text = objOp.sCodigo

    End If

    Exit Sub

Erro_objEventoOp_evSelecao:

    Select Case Err

       Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170097)

    End Select

    Exit Sub

End Sub


Function Obtem_Descricao_Ccl(sCcl As String, sDescCcl As String) As Long
'recebe em sCcl o Ccl no formato do Bd
'retorna em sDescCcl a descrição do Ccl ( que será formatado para tela )

Dim lErro As Long, iCclPreenchida As Integer
Dim objCcl As New ClassCcl
Dim sCopia As String

On Error GoTo Erro_Obtem_Descricao_Ccl

    sCopia = sCcl
    sDescCcl = String(STRING_CCL_DESCRICAO, 0)
    sCcl = String(STRING_CCL_MASK, 0)

    'determina qual Ccl deve ser lido
    objCcl.sCcl = sCopia

    lErro = Mascara_MascararCcl(sCopia, sCcl)
    If lErro <> SUCESSO Then Error 34436

    'verifica se a conta está preenchida
    lErro = CF("Ccl_Formata", sCcl, sCopia, iCclPreenchida)
    If lErro <> SUCESSO Then Error 34437

    If iCclPreenchida = CCL_PREENCHIDA Then

        'verifica se a Ccl existe
        lErro = CF("Ccl_Le", objCcl)
        If lErro <> SUCESSO Then Error 34438

        sDescCcl = objCcl.sDescCcl

    Else

        sCcl = ""
        sDescCcl = ""

    End If

    Obtem_Descricao_Ccl = SUCESSO

    Exit Function

Erro_Obtem_Descricao_Ccl:

    Obtem_Descricao_Ccl = Err

    Select Case Err

        Case 34436
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, sCopia)

        Case 34437

        Case 34438

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170098)

    End Select

    Exit Function

End Function

Function Ccl_Perde_Foco(Ccl As Object, Desc As Object) As Long
'recebe MaskEdBox do Ccl e o label da descrição

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclFormatada As String
Dim iCclPreenchida As Integer

On Error GoTo Erro_Ccl_Perde_Foco

    sCclFormatada = String(STRING_CCL, 0)

    Desc.Caption = ""

    lErro = CF("Ccl_Formata", Ccl.Text, sCclFormatada, iCclPreenchida)
    If lErro Then Error 34439

    If iCclPreenchida = CCL_PREENCHIDA Then

        'verifica se a Ccl existe
        objCcl.sCcl = sCclFormatada
        lErro = CF("Ccl_Le", objCcl)
        If lErro <> SUCESSO And lErro <> 5599 Then Error 34440

        If lErro = 5599 Then Error 34441

        Desc.Caption = objCcl.sDescCcl

    End If

    Ccl_Perde_Foco = SUCESSO

    Exit Function

Erro_Ccl_Perde_Foco:

    Ccl_Perde_Foco = Err

    Select Case Err

        Case 34439

        Case 34440

        Case 34441
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", Err, Ccl.Text)
            Ccl.SetFocus


        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170099)

    End Select

    Exit Function

End Function

Private Function Carga_Arvore_Ccl(colNodes As Nodes) As Long
'move os dados de centro de custo/lucro do banco de dados para a arvore colNodes. /m

Dim objNode As Node
Dim colCcl As New Collection
Dim objCcl As ClassCcl
Dim lErro As Long
Dim sCclMascarado As String
Dim sCcl As String
Dim sCclPai As String
    
On Error GoTo Erro_Carga_Arvore_Ccl
    
    'leitura dos centro de custo/lucro no BD
    lErro = CF("Ccl_Le_Todos", colCcl)
    If lErro <> SUCESSO Then Error 34442
    
    'para cada centro de custo encontrado no bd
    For Each objCcl In colCcl
        
        sCclMascarado = String(STRING_CCL, 0)

        'coloca a mascara no centro de custo
        lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then Error 34443

        sCcl = "C" & objCcl.sCcl

        sCclPai = String(STRING_CCL, 0)
        
        'retorna o centro de custo/lucro "pai" da centro de custo/lucro em questão, se houver
        lErro = Mascara_RetornaCclPai(objCcl.sCcl, sCclPai)
        If lErro <> SUCESSO Then Error 54504
        
        'se o centro de custo/lucro possui um centro de custo/lucro "pai"
        If Len(Trim(sCclPai)) > 0 Then

            sCclPai = "C" & sCclPai
            
            'adiciona o centro de custo como filho do centro de custo pai
            Set objNode = colNodes.Add(colNodes.Item(sCclPai), tvwChild, sCcl)

        Else
        
            'se o centro de custo/lucro não possui centro de custo/lucro "pai", adiciona na árvore sem pai
            Set objNode = colNodes.Add(, tvwLast, sCcl)
            
        End If
        
        'coloca o texto do nó que acabou de ser inserido
        objNode.Text = sCclMascarado & SEPARADOR & objCcl.sDescCcl
        
    Next
    
    Carga_Arvore_Ccl = SUCESSO

    Exit Function

Erro_Carga_Arvore_Ccl:

    Carga_Arvore_Ccl = Err

    Select Case Err

        Case 54504
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCclPai", Err, objCcl.sCcl)

        Case 34442

        Case 34443
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, objCcl.sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 170100)

    End Select
    
    Exit Function

End Function

''Function Carga_Arvore_Ccl(colNodes As Nodes) As Long
'''move os dados de centro de custo/lucro do banco de dados para a arvore colNodes.
''
''Dim objNode As Node
''Dim colCcl As New Collection
''Dim objCcl As ClassCcl
''Dim lErro As Long
''Dim sCclMascarado As String
''
''On Error GoTo Erro_Carga_Arvore_Ccl
''
''    lErro = CF("Ccl_Le_Todos",colCcl)
''    If lErro <> SUCESSO Then Error 34442
''
''    For Each objCcl In colCcl
''
''        sCclMascarado = String(STRING_CCL, 0)
''
''        lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
''        If lErro <> SUCESSO Then Error 34443
''
''        Set objNode = colNodes.Add(, , "C" & objCcl.sCcl, sCclMascarado & SEPARADOR & objCcl.sDescCcl)
''
''    Next
''
''    Carga_Arvore_Ccl = SUCESSO
''
''    Exit Function
''
''Erro_Carga_Arvore_Ccl:
''
''    Carga_Arvore_Ccl = Err
''
''    Select Case Err
''
''        Case 34442
''
''        Case 34443
''            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, objCcl.sCcl)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 170101)
''
''    End Select
''
''    Exit Function
''
''End Function



Private Sub CclFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CclFinal_Validate

    giCclInicial = 0

    lErro = Ccl_Perde_Foco(CclFinal, DescCclFim)
    If lErro <> SUCESSO Then Error 34444

    Exit Sub

Erro_CclFinal_Validate:

    Cancel = True


    Select Case Err

        Case 34444

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170102)

    End Select

    Exit Sub

End Sub

Private Sub CclInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CclInicial_Validate

    giCclInicial = 1

    lErro = Ccl_Perde_Foco(CclInicial, DescCclInic)
    If lErro <> SUCESSO Then Error 34445

    Exit Sub

Erro_CclInicial_Validate:

    Cancel = True


    Select Case Err

        Case 34445

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170103)

    End Select

    Exit Sub

End Sub


Private Sub TvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)

Dim sCcl As String
Dim sCclMascarado As String
Dim lErro As Long
Dim lPosicaoSeparador As Long

On Error GoTo Erro_TvwCcls_NodeClick

    sCcl = right(Node.Key, Len(Node.Key) - 1)

    sCclMascarado = String(STRING_CCL, 0)

    lErro = Mascara_MascararCcl(sCcl, sCclMascarado)
    If lErro <> SUCESSO Then Error 34446

    If giCclInicial = 1 Then
        CclInicial.PromptInclude = False
        CclInicial.Text = sCclMascarado
        CclInicial.PromptInclude = True
    Else
        CclFinal.PromptInclude = False
        CclFinal.Text = sCclMascarado
        CclFinal.PromptInclude = True
    End If

    'Preenche a Descricao do centro de custo/lucro
    lPosicaoSeparador = InStr(Node.Text, SEPARADOR)

    If giCclInicial = 1 Then
        DescCclInic.Caption = Mid(Node.Text, lPosicaoSeparador + 1)
    Else
        DescCclFim.Caption = Mid(Node.Text, lPosicaoSeparador + 1)
    End If

    Exit Sub

Erro_TvwCcls_NodeClick:

    Select Case Err

        Case 34446
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 170104)

    End Select

    Exit Sub

End Sub


Private Function Valida_OrdProd(sCodigoOP As String) As Long
Dim objOp As New ClassOrdemDeProducao
Dim lErro As Long

On Error GoTo Erro_Valida_OrdProd

    objOp.iFilialEmpresa = giFilialEmpresa
    objOp.sCodigo = sCodigoOP
   
    'busca ordem de produção aberta
    lErro = CF("OrdemDeProducao_Le_SemItens", objOp)
    If lErro <> SUCESSO And lErro <> 34455 Then

        Error 34447

    Else
    
        'se não existe ordem de produção aberta
        If lErro <> SUCESSO Then

            'busca ordem de produção baixada
            lErro = CF("OPBaixada_Le_SemItens", objOp)
            If lErro <> SUCESSO And lErro <> 34459 Then Error 34448

            If lErro = 34459 Then Error 34449
            
        End If

    End If

    
    Valida_OrdProd = SUCESSO

    Exit Function

Erro_Valida_OrdProd:

    Valida_OrdProd = Err

    Select Case Err

                   
        Case 34447, 34448
            
        Case 34449
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OP_INEXISTENTE", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170105)

    End Select

    Exit Function

End Function


Private Sub OpInicial_Validate(Cancel As Boolean)
Dim lErro As Long

On Error GoTo Erro_OpInicial_Validate

    giOp_Inicial = 1

    If Len(Trim(OpInicial.Text)) <> 0 Then

        lErro = Valida_OrdProd(OpInicial.Text)
        If lErro <> SUCESSO Then Error 34450
        
    End If

    Exit Sub

Erro_OpInicial_Validate:

    Cancel = True


    Select Case Err

        Case 34450
                   
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170106)

    End Select

    Exit Sub

End Sub

Private Sub OpFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_OpFinal_Validate

    giOp_Inicial = 0

    If Len(Trim(OpFinal.Text)) <> 0 Then

        lErro = Valida_OrdProd(OpFinal.Text)
        If lErro <> SUCESSO Then Error 34451
    
    End If

    Exit Sub

Erro_OpFinal_Validate:

    Cancel = True


    Select Case Err

        Case 34451
                   
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170107)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_MOVIMENTO_INTERNO
    Set Form_Load_Ocx = Me
    Caption = "Relação das Movimentações Internas"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpMovInt"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is OpInicial Then
            Call LabelOpInicial_Click
        ElseIf Me.ActiveControl Is OpFinal Then
            Call LabelOpFinal_Click
        ElseIf Me.ActiveControl Is ProdutoInicial Then
            Call LabelProdutoDe_Click
        ElseIf Me.ActiveControl Is ProdutoFinal Then
            Call LabelProdutoAte_Click
        End If
        
    
    End If

End Sub



Private Sub dIni_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dIni, Source, X, Y)
End Sub

Private Sub dIni_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dIni, Button, Shift, X, Y)
End Sub

Private Sub dFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dFim, Source, X, Y)
End Sub

Private Sub dFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dFim, Button, Shift, X, Y)
End Sub

Private Sub DescCclInic_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescCclInic, Source, X, Y)
End Sub

Private Sub DescCclInic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescCclInic, Button, Shift, X, Y)
End Sub

Private Sub DescCclFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescCclFim, Source, X, Y)
End Sub

Private Sub DescCclFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescCclFim, Button, Shift, X, Y)
End Sub

Private Sub LabelCclAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCclAte, Source, X, Y)
End Sub

Private Sub LabelCclAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCclAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCclDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCclDe, Source, X, Y)
End Sub

Private Sub LabelCclDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCclDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCcl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCcl, Source, X, Y)
End Sub

Private Sub LabelCcl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCcl, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
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

Private Sub LabelOpInicial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelOpInicial, Source, X, Y)
End Sub

Private Sub LabelOpInicial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelOpInicial, Button, Shift, X, Y)
End Sub

Private Sub LabelOpFinal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelOpFinal, Source, X, Y)
End Sub

Private Sub LabelOpFinal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelOpFinal, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub


Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub

