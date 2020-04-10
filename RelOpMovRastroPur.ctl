VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpMovRastro 
   ClientHeight    =   7425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8595
   ScaleHeight     =   7425
   ScaleWidth      =   8595
   Begin VB.CheckBox ExibirEndereco 
      Caption         =   "Exibir o endereço dos clientes"
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
      TabIndex        =   51
      Top             =   7125
      Value           =   1  'Checked
      Width           =   7980
   End
   Begin VB.Frame Frame3 
      Caption         =   "Vendedores"
      Height          =   675
      Left            =   120
      TabIndex        =   46
      Top             =   5400
      Width           =   8235
      Begin VB.OptionButton OptVendIndir 
         Caption         =   "Vendas Indiretas"
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
         Left            =   2220
         TabIndex        =   47
         Top             =   225
         Width           =   2010
      End
      Begin VB.OptionButton OptVendDir 
         Caption         =   "Vendas Diretas"
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
         Left            =   465
         TabIndex        =   48
         Top             =   225
         Value           =   -1  'True
         Width           =   2340
      End
      Begin MSMask.MaskEdBox Vendedor 
         Height          =   300
         Left            =   5385
         TabIndex        =   49
         Top             =   285
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelVendedor 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor:"
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
         Left            =   4440
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   50
         Top             =   330
         Width           =   885
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Endereço"
      Height          =   645
      Left            =   120
      TabIndex        =   43
      Top             =   4710
      Width           =   5670
      Begin MSMask.MaskEdBox Cidade 
         Height          =   315
         Left            =   1215
         TabIndex        =   44
         Top             =   210
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   12
         PromptChar      =   " "
      End
      Begin VB.Label LabelCidade 
         AutoSize        =   -1  'True
         Caption         =   "Cidade:"
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
         Left            =   480
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   45
         Top             =   255
         Width           =   660
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Região"
      Height          =   975
      Left            =   120
      TabIndex        =   36
      Top             =   6105
      Width           =   8235
      Begin MSMask.MaskEdBox RegiaoInicial 
         Height          =   315
         Left            =   840
         TabIndex        =   37
         Top             =   180
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox RegiaoFinal 
         Height          =   315
         Left            =   840
         TabIndex        =   38
         Top             =   555
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label RegiaoAte 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   2520
         TabIndex        =   42
         Top             =   555
         Width           =   5655
      End
      Begin VB.Label RegiaoDe 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   2520
         TabIndex        =   41
         Top             =   180
         Width           =   5655
      End
      Begin VB.Label LabelRegiaoDe 
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
         Left            =   480
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   40
         Top             =   225
         Width           =   360
      End
      Begin VB.Label LabelRegiaoAte 
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
         Left            =   435
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   39
         Top             =   600
         Width           =   435
      End
   End
   Begin VB.Frame FrameNF 
      Caption         =   "Lote"
      Height          =   1005
      Left            =   120
      TabIndex        =   32
      Top             =   3705
      Width           =   5670
      Begin VB.ComboBox FilialOP 
         Height          =   315
         Left            =   1200
         TabIndex        =   12
         Top             =   585
         Width           =   2805
      End
      Begin MSMask.MaskEdBox LoteInicial 
         Height          =   300
         Left            =   495
         TabIndex        =   10
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox LoteFinal 
         Height          =   300
         Left            =   3360
         TabIndex        =   11
         Top             =   225
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Filial OP:"
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
         Left            =   375
         TabIndex        =   35
         Top             =   645
         Width           =   780
      End
      Begin VB.Label LabelLoteInicial 
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
         Left            =   135
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   34
         Top             =   270
         Width           =   315
      End
      Begin VB.Label LabelLoteFinal 
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
         Left            =   2955
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   33
         Top             =   285
         Width           =   360
      End
   End
   Begin VB.ListBox Almoxarifados 
      Height          =   4155
      ItemData        =   "RelOpMovRastroPur.ctx":0000
      Left            =   5820
      List            =   "RelOpMovRastroPur.ctx":0002
      TabIndex        =   9
      Top             =   1215
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Height          =   525
      Left            =   120
      TabIndex        =   29
      Top             =   990
      Width           =   5670
      Begin VB.OptionButton OptAlmoxarifado 
         Caption         =   "Almoxarifado"
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
         Left            =   120
         TabIndex        =   6
         Top             =   180
         Width           =   1425
      End
      Begin VB.OptionButton OptFilial 
         Caption         =   "Filial"
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
         Left            =   2670
         TabIndex        =   7
         Top             =   180
         Width           =   915
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data de Movimento do Estoque"
      Height          =   660
      Left            =   120
      TabIndex        =   25
      Top             =   2025
      Width           =   5670
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   1935
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataMovInicial 
         Height          =   300
         Left            =   930
         TabIndex        =   15
         Top             =   255
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
         Height          =   315
         Left            =   4545
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataMovFinal 
         Height          =   300
         Left            =   3540
         TabIndex        =   17
         Top             =   255
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3135
         TabIndex        =   27
         Top             =   315
         Width           =   360
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
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   540
         TabIndex        =   26
         Top             =   285
         Width           =   345
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produtos"
      Height          =   975
      Left            =   120
      TabIndex        =   20
      Top             =   2700
      Width           =   5670
      Begin MSMask.MaskEdBox ProdutoFinal 
         Height          =   315
         Left            =   495
         TabIndex        =   14
         Top             =   570
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
         Left            =   495
         TabIndex        =   13
         Top             =   210
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   105
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   24
         Top             =   615
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   165
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   23
         Top             =   240
         Width           =   315
      End
      Begin VB.Label DescProdFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2010
         TabIndex        =   22
         Top             =   570
         Width           =   3570
      End
      Begin VB.Label DescProdInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2010
         TabIndex        =   21
         Top             =   210
         Width           =   3585
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpMovRastroPur.ctx":0004
      Left            =   960
      List            =   "RelOpMovRastroPur.ctx":0006
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   413
      Width           =   2910
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
      Left            =   4080
      Picture         =   "RelOpMovRastroPur.ctx":0008
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   270
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5910
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   270
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpMovRastroPur.ctx":010A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpMovRastroPur.ctx":0288
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpMovRastroPur.ctx":07BA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpMovRastroPur.ctx":0944
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Almoxarifado 
      Height          =   315
      Left            =   1335
      TabIndex        =   8
      Top             =   1650
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Label LabelAlmoxarifado 
      AutoSize        =   -1  'True
      Caption         =   "Almoxarifados"
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
      Left            =   5835
      TabIndex        =   31
      Top             =   945
      Width           =   1185
   End
   Begin VB.Label lblAlmoxarifado 
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
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   1695
      Width           =   1185
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
      Left            =   255
      TabIndex        =   28
      Top             =   443
      Width           =   615
   End
End
Attribute VB_Name = "RelOpMovRastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1
Private WithEvents objEventoLoteInicial As AdmEvento
Attribute objEventoLoteInicial.VB_VarHelpID = -1
Private WithEvents objEventoLoteFinal As AdmEvento
Attribute objEventoLoteFinal.VB_VarHelpID = -1

Private WithEvents objEventoCidade As AdmEvento
Attribute objEventoCidade.VB_VarHelpID = -1
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1
Private WithEvents objEventoRegiaoVenda As AdmEvento
Attribute objEventoRegiaoVenda.VB_VarHelpID = -1

Dim giRegiaoVenda As Integer

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Private Sub Form_Load()

Dim lErro As Long
Dim iOpcao As Integer

On Error GoTo Erro_Form_Load
    
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    Set objEventoLoteInicial = New AdmEvento
    Set objEventoLoteFinal = New AdmEvento
    
    Set objEventoVendedor = New AdmEvento
    Set objEventoCidade = New AdmEvento
    Set objEventoRegiaoVenda = New AdmEvento

    'Inicializa a mascara dos produtos
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoInicial)
    If lErro <> SUCESSO Then gError 85103

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoFinal)
    If lErro <> SUCESSO Then gError 85104
    
    '###########################################
    'Inserido por Wagner 10/08/2006
    'carrega a ListBox Almoxarifados
    lErro = Carrega_Lista_Almoxarifado()
    If lErro <> SUCESSO Then gError 171789
    
    'Carrega FilialOP a partir das Filiais Empresas
    lErro = Carrega_FilialOP()
    If lErro <> SUCESSO Then gError 171790
    '###########################################
     
    OptAlmoxarifado.Value = True
    OptVendDir.Value = True
    giRegiaoVenda = 1
     
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:
   
   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 85103, 85104, 171789, 171790
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170108)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim iIndice As Integer

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 85105
    
    'pega Região de Venda Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TREGIAOVENDAINIC", sParam)
    If lErro <> SUCESSO Then gError 171791

    If StrParaLong(sParam) > 0 Then
        RegiaoInicial.Text = sParam
        Call RegiaoInicial_Validate(bSGECancelDummy)
    End If
    
    'pega Região de Venda Final e exibe
    lErro = objRelOpcoes.ObterParametro("TREGIAOVENDAFIM", sParam)
    If lErro <> SUCESSO Then gError 171791

    If StrParaLong(sParam) > 0 Then
        RegiaoFinal.Text = sParam
        Call RegiaoFinal_Validate(bSGECancelDummy)
    End If

    '#############################################################
    'Inserido por Wagner 10/08/2006
    lErro = objRelOpcoes.ObterParametro("NALMOX", sParam)
    If lErro Then gError 171791
    
    If StrParaInt(sParam) <> 0 Then
        Almoxarifado.Text = sParam
        Call Almoxarifado_Validate(bSGECancelDummy)
    Else
        Almoxarifado.Text = ""
    End If
    
    lErro = objRelOpcoes.ObterParametro("NOPTALMOXARIFADO", sParam)
    If lErro Then gError 171792
    
    If StrParaInt(sParam) = MARCADO Then
        OptAlmoxarifado.Value = True
    Else
        OptFilial.Value = True
    End If
    
    lErro = objRelOpcoes.ObterParametro("TLOTEINI", sParam)
    If lErro Then gError 171793
    
    LoteInicial.Text = sParam

    lErro = objRelOpcoes.ObterParametro("TLOTEFIM", sParam)
    If lErro Then gError 171794
    
    LoteFinal.Text = sParam
    
    lErro = objRelOpcoes.ObterParametro("NFILIALOP", sParam)
    If lErro Then gError 171795
    
    If StrParaInt(sParam) <> 0 Then
        Call Combo_Seleciona(FilialOP, StrParaInt(sParam))
    Else
        FilialOP.ListIndex = -1
    End If
    '#############################################################
            
    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro <> SUCESSO Then gError 85106

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 85107

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 85108

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 85109
    
    'Data Inicial
    lErro = objRelOpcoes.ObterParametro("DATAMOVINICIAL", sParam)
    If lErro <> SUCESSO Then gError 85110
    
    'coloca a data Inicial na tela
    Call DateParaMasked(DataMovInicial, CDate(sParam))
    
    'Data Final
    lErro = objRelOpcoes.ObterParametro("DATAMOVFINAL", sParam)
    If lErro <> SUCESSO Then gError 85111
    
    'coloca a data Final na tela
    Call DateParaMasked(DataMovFinal, CDate(sParam))
    
    lErro = objRelOpcoes.ObterParametro("NVENDEDOR", sParam)
    If lErro <> SUCESSO Then gError 85111

    If sParam <> "0" Then
        Vendedor.Text = CInt(sParam)
        Call Vendedor_Validate(bSGECancelDummy)
    End If

    lErro = objRelOpcoes.ObterParametro("NTIPOVEND", sParam)
    If lErro <> SUCESSO Then gError 85111

    If StrParaInt(sParam) = VENDEDOR_DIRETO Then
        OptVendDir.Value = True
    Else
        OptVendIndir.Value = True
    End If
    
    lErro = objRelOpcoes.ObterParametro("TCIDADE", sParam)
    If lErro <> SUCESSO Then gError 85111

    Cidade.Text = sParam
    
    lErro = objRelOpcoes.ObterParametro("NEXIBEENDERECO", sParam)
    If lErro <> SUCESSO Then gError 85111

    If StrParaInt(sParam) = MARCADO Then
        ExibirEndereco.Value = vbChecked
    Else
        ExibirEndereco.Value = vbUnchecked
    End If
    
    PreencherParametrosNaTela = SUCESSO
    
    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 85105 To 85111, 171791 To 171795
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170109)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 85113
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 85112

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 85112
        
        Case 85113
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170110)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 85114

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 85115

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 85116

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 85114, 85116

        Case 85115
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170111)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 85117

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 85118

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 85119

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 85117, 85119

        Case 85118
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170112)

    End Select

    Exit Sub

End Sub

Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String, iTipoVend As Integer) As Long
'Formata os produtos retornando em sProd_I e sProd_F
'Verifica se os parâmetros iniciais são maiores que os finais

Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 85120

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 85121

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambas os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then gError 85122

    End If
    
    '#####################################################################
    'Inserido por Wagner 10/08/2006
    If OptAlmoxarifado.Value = True Then
        'O campo almoxarifado deve ser preenchido
        If Trim(Almoxarifado.Text) = "" Then gError 171803
    End If
    
    If Len(Trim(LoteInicial.Text)) <> 0 And Len(Trim(LoteFinal.Text)) <> 0 Then
        If LoteInicial.Text > LoteFinal.Text Then gError 171788
    End If
    
    If StrParaDate(DataMovInicial.Text) = DATA_NULA Then gError 181822
    
    If StrParaDate(DataMovFinal.Text) = DATA_NULA Then gError 181823
    '#####################################################################
    
    'Se RegiãoInicial e RegiãoFinal estão preenchidos
    If Len(Trim(RegiaoInicial.Text)) > 0 And Len(Trim(RegiaoFinal.Text)) > 0 Then
    
        'Se Região inicial for maior que Região final, erro
        If StrParaLong(RegiaoInicial.Text) > StrParaLong(RegiaoFinal.Text) Then gError 87095
        
    End If
    
    If OptVendDir.Value Then
        iTipoVend = VENDEDOR_DIRETO
    Else
        iTipoVend = VENDEDOR_INDIRETO
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
    
        Case 85120
            ProdutoInicial.SetFocus

        Case 85121
            ProdutoFinal.SetFocus
            
        Case 85122
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            ProdutoInicial.SetFocus

        Case 87095
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGIAOVENDA_INICIAL_MAIOR", gErr)

        Case 171788
            Call Rotina_Erro(vbOKOnly, "ERRO_LOTE_INICIAL_MAIOR", gErr)
            LoteInicial.SetFocus
        
        Case 171803
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_PREENCHIDO1", gErr)
            Almoxarifado.SetFocus
    
        Case 181822
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_NAO_PREENCHIDA", gErr)
            DataMovInicial.SetFocus
        
        Case 181823
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAFINAL_NAO_PREENCHIDA", gErr)
            DataMovFinal.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170113)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 85123
    
    ComboOpcoes.Text = ""
    DescProdInic.Caption = ""
    DescProdFim.Caption = ""
    ComboOpcoes.SetFocus
    
    OptAlmoxarifado.Value = True
    
    FilialOP.ListIndex = -1
    
    RegiaoAte.Caption = ""
    RegiaoDe.Caption = ""
    
    OptVendDir.Value = True
    giRegiaoVenda = 1
        
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 85123
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170114)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    Set objEventoLoteInicial = Nothing
    Set objEventoLoteFinal = Nothing
    
    Set objEventoCidade = Nothing
    Set objEventoVendedor = Nothing
    Set objEventoRegiaoVenda = Nothing
    
End Sub

Private Sub DataMovFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataMovFinal)

End Sub

Private Sub DataMovInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataMovInicial)

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
        If lErro <> SUCESSO Then gError 85123

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 85123

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170115)

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
        If lErro <> SUCESSO Then gError 85124

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 85124

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170116)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional ByVal bExecutando As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sProd_I As String
Dim sProd_F As String
Dim iOpt As Integer
Dim iTipoVend As Integer
Dim lNumIntRel As Long
Dim iExibeEnd As Integer

On Error GoTo Erro_PreencherRelOp
    
    
    'Critica data se  a Inicial e a Final estiverem Preenchida
    If Len(DataMovInicial.ClipText) <> 0 And Len(DataMovFinal.ClipText) <> 0 Then
        'data inicial não pode ser maior que a data final
        If CDate(DataMovInicial.Text) > CDate(DataMovFinal.Text) Then gError 85127
    End If
    
    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)
    
    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F, iTipoVend)
    If lErro <> SUCESSO Then gError 85128
      
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 85129
    
    'se a data não for preenchida não move
    If Trim(DataMovInicial.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DATAMOVINICIAL", DataMovInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DATAMOVINICIAL", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 85130
    
    'se a data não for preenchida não move
    If Trim(DataMovFinal.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DATAMOVFINAL", DataMovFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DATAMOVFINAL", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 85131
    
    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then gError 85132

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then gError 85133
    
    '#############################################################
    'Inserido por Wagner 10/08/2006
    lErro = objRelOpcoes.IncluirParametro("NALMOX", CStr(Codigo_Extrai(Almoxarifado.Text)))
    If lErro <> AD_BOOL_TRUE Then gError 171796
        
    lErro = objRelOpcoes.IncluirParametro("TALMOXARIFADO", Almoxarifado.Text)
    If lErro <> AD_BOOL_TRUE Then gError 171797
    
    If OptAlmoxarifado.Value Then
        iOpt = MARCADO
    Else
        iOpt = DESMARCADO
    End If
    
    lErro = objRelOpcoes.IncluirParametro("NOPTALMOXARIFADO", CStr(iOpt))
    If lErro <> AD_BOOL_TRUE Then gError 171798
    
    lErro = objRelOpcoes.IncluirParametro("NFILIALOP", CStr(Codigo_Extrai(FilialOP.Text)))
    If lErro <> AD_BOOL_TRUE Then gError 171799

    lErro = objRelOpcoes.IncluirParametro("TFILIALOP", FilialOP.Text)
    If lErro <> AD_BOOL_TRUE Then gError 171800

    lErro = objRelOpcoes.IncluirParametro("TLOTEINI", LoteInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 171801
    
    lErro = objRelOpcoes.IncluirParametro("TLOTEFIM", LoteFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 171802
    '############################################################
    
    lErro = objRelOpcoes.IncluirParametro("TCIDADE", Cidade.Text)
    If lErro <> AD_BOOL_TRUE Then gError 171802

    lErro = objRelOpcoes.IncluirParametro("TVENDEDOR", Vendedor.Text)
    If lErro <> AD_BOOL_TRUE Then gError 171802

    lErro = objRelOpcoes.IncluirParametro("NVENDEDOR", Codigo_Extrai(Vendedor.Text))
    If lErro <> AD_BOOL_TRUE Then gError 171802

    lErro = objRelOpcoes.IncluirParametro("NTIPOVEND", CStr(iTipoVend))
    If lErro <> AD_BOOL_TRUE Then gError 171802
    
    lErro = objRelOpcoes.IncluirParametro("TREGIAOVENDAINIC", RegiaoInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 171802

    lErro = objRelOpcoes.IncluirParametro("TREGIAOVENDAFIM", RegiaoFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 171802
    
    If bExecutando Then
    
        lErro = CF("RelOpFatProdVend_Prepara", iTipoVend, Codigo_Extrai(Vendedor.Text), lNumIntRel)
        If lErro <> SUCESSO Then gError 171802
    
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError 171802
    
    End If
    
    If ExibirEndereco.Value = vbChecked Then
        iExibeEnd = MARCADO
    Else
        iExibeEnd = DESMARCADO
    End If
    
    lErro = objRelOpcoes.IncluirParametro("NEXIBEENDERECO", CStr(iExibeEnd))
    If lErro <> AD_BOOL_TRUE Then gError 171802
   
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd_I, sProd_F)
    If lErro <> SUCESSO Then gError 85134

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 85127
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)

        Case 85128 To 85134, 171796 To 171802
        
        Case 85125, 85126
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170117)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 85135
    
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 85136
        
        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError 85137
    
        ComboOpcoes.Text = ""
        DescProdInic.Caption = ""
        DescProdFim.Caption = ""
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 85135
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 85136, 85137

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170118)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 85138 '47262
    
    If Len(Trim(Vendedor.Text)) > 0 Then
        gobjRelatorio.sNomeTsk = "MovRast2"
    Else
        gobjRelatorio.sNomeTsk = "MovRast"
    End If
 
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 85138

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170119)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 85139

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 85140

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 85141

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 85142
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 85139
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus
           
        Case 85140, 85141, 85142
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170120)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate

    lErro = CF("Produto_Perde_Foco", ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 85143
    
    If lErro <> SUCESSO Then gError 85144

    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True


    Select Case gErr

        Case 85143
        
        Case 85144
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
          
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170121)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    lErro = CF("Produto_Perde_Foco", ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 85145
    
    If lErro <> SUCESSO Then gError 85146

    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True


    Select Case gErr

        Case 85145
        
        Case 85146
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170122)

    End Select

    Exit Sub

End Sub


Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd_I As String, sProd_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""
    
    If sProd_I <> "" Then sExpressao = "Produto >= " & Forprint_ConvTexto(sProd_I)

    If sProd_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Produto <= " & Forprint_ConvTexto(sProd_F)

    End If
   
    If Trim(DataMovInicial.ClipText) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & " DataMov >= " & Forprint_ConvData(CDate(DataMovInicial.Text))
    End If

    If Trim(DataMovFinal.ClipText) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & " DataMov <= " & Forprint_ConvData(CDate(DataMovFinal.Text))
    End If
    
    '######################################################################
    'Inserido por Wagner 10/08/2006
    If LoteInicial.Text <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Lote >= " & Forprint_ConvTexto(LoteInicial.Text)

    End If
    
    If LoteFinal.Text <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Lote <= " & Forprint_ConvTexto(LoteFinal.Text)

    End If
    
    If Len(Trim(FilialOP.Text)) <> 0 Then
    
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & " FilialOP = " & Forprint_ConvInt(Codigo_Extrai(FilialOP.Text))
   
    End If
    
    If Len(Trim(Almoxarifado.Text)) <> 0 Then
    
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & " Almoxarifado = " & Forprint_ConvInt(Codigo_Extrai(Almoxarifado.Text))
   
    End If
    '######################################################################

    If Len(Trim(RegiaoInicial.Text)) <> 0 Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Regiao >= " & Forprint_ConvInt(Codigo_Extrai(RegiaoInicial.Text))
    
    End If
    
    If Len(Trim(RegiaoFinal.Text)) <> 0 Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Regiao <= " & Forprint_ConvInt(Codigo_Extrai(RegiaoFinal.Text))

    End If
    
    If Len(Trim(Cidade.Text)) <> 0 Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Cidade = " & Forprint_ConvTexto(Cidade.Text)

    End If
    
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If
 
    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170123)

    End Select

    Exit Function

End Function

Private Sub DataMovFinal_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_DataMovFinal_Validate

    If Len(DataMovFinal.ClipText) > 0 Then

        sDataFim = DataMovFinal.Text
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then gError 85147

    End If

    Exit Sub

Erro_DataMovFinal_Validate:

    Cancel = True


    Select Case gErr

        Case 85147

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170124)

    End Select

    Exit Sub

End Sub

Private Sub DataMovInicial_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_DataMovInicial_Validate

    If Len(DataMovInicial.ClipText) > 0 Then

        sDataInic = DataMovInicial.Text
        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then gError 85148

    End If

    Exit Sub

Erro_DataMovInicial_Validate:

    Cancel = True


    Select Case gErr

        Case 85148

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170125)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataMovInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 85149
    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 85149
            DataMovInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170126)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataMovInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 85150

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 85150
            DataMovInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170127)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataMovFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 85151
    
    Exit Sub

Erro_UpDown2_DownClick:

    Select Case gErr

        Case 85151
            DataMovFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170128)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataMovFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 85152

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case gErr

        Case 85152
            DataMovFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170129)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    'Parent.HelpContextID = IDH_RELOP_ANALISE_MOVIMENTO_ESTOQUE
    Set Form_Load_Ocx = Me
    Caption = "Movimentos de Rastreamento"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpMovRastroOCX"
    
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
    
        If Me.ActiveControl Is ProdutoInicial Then
            Call LabelProdutoDe_Click
        ElseIf Me.ActiveControl Is ProdutoFinal Then
            Call LabelProdutoAte_Click
        ElseIf Me.ActiveControl Is RegiaoInicial Then
            Call LabelRegiaoDe_Click
        ElseIf Me.ActiveControl Is RegiaoFinal Then
            Call LabelRegiaoAte_Click
        ElseIf Me.ActiveControl Is Vendedor Then
            Call LabelVendedor_Click
        ElseIf Me.ActiveControl Is Cidade Then
            Call LabelCidade_Click
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


Private Sub DescProdInic_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescProdInic, Source, X, Y)
End Sub

Private Sub DescProdInic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescProdInic, Button, Shift, X, Y)
End Sub

Private Sub DescProdFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescProdFim, Source, X, Y)
End Sub

Private Sub DescProdFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescProdFim, Button, Shift, X, Y)
End Sub

Private Sub LabelProdutoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProdutoDe, Source, X, Y)
End Sub

Private Sub LabelProdutoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProdutoDe, Button, Shift, X, Y)
End Sub

Private Sub LabelProdutoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProdutoAte, Source, X, Y)
End Sub

Private Sub LabelProdutoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProdutoAte, Button, Shift, X, Y)
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

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

'####################################################
'Inserido por Wagner 10/08/2006
Private Sub LblAlmoxarifado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(lblAlmoxarifado, Source, X, Y)
End Sub

Private Sub LblAlmoxarifado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(lblAlmoxarifado, Button, Shift, X, Y)
End Sub

Private Sub OptAlmoxarifado_Click()
    Almoxarifado.Enabled = True
    Almoxarifados.Enabled = True
End Sub

Private Sub OptFilial_Click()
    Almoxarifado.Enabled = False
    Almoxarifados.Enabled = False
End Sub

Private Sub Almoxarifado_Validate(Cancel As Boolean)

Dim lErro As Long

Dim objAlmoxarifado As New ClassAlmoxarifado
'Dim sContaEnxuta As String

On Error GoTo Erro_Almoxarifado_Validate

    If Len(Trim(Almoxarifado.Text)) > 0 Then
    
        If Codigo_Extrai(Almoxarifado.Text) <> 0 Then
            Almoxarifado.Text = Codigo_Extrai(Almoxarifado.Text)
        End If
        
        'Le o almoxarifado pelo código ou pelo nome reduzido e joga o nome reduzido em Almoxarifado.Text
        lErro = TP_Almoxarifado_Le(Almoxarifado, objAlmoxarifado)
        If lErro <> SUCESSO Then gError 171804
                
        Almoxarifado = objAlmoxarifado.iCodigo & SEPARADOR & objAlmoxarifado.sNomeReduzido
        
        'se o almoxarifado não pertencer a filial em questão ==> erro
        If objAlmoxarifado.iFilialEmpresa <> giFilialEmpresa Then gError 171805
        
    End If
    
    Exit Sub

Erro_Almoxarifado_Validate:

    Cancel = True

    Select Case gErr

        Case 171804

        Case 171805
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_FILIAL_DIFERENTE", gErr, objAlmoxarifado.iCodigo, giFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 181825)

    End Select

    Exit Sub

End Sub

Private Sub Almoxarifados_DblClick()
'Preenche Almoxarifado Final ou Inicial com o almoxarifado selecionado

Dim sListBoxItem As String
Dim lErro As Long

On Error GoTo Erro_Almoxarifados_DblClick

   'Guarda a string selecionada na ListBox Almoxarifados
    sListBoxItem = Almoxarifados.List(Almoxarifados.ListIndex)
    
    Almoxarifado.Text = sListBoxItem

    Exit Sub

Erro_Almoxarifados_DblClick:

    Select Case gErr

    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181826)

    End Select

    Exit Sub

End Sub

Private Function Carrega_Lista_Almoxarifado() As Long
'Carrega a ListBox Almoxarifados

Dim lErro As Long
Dim colAlmoxarifados As New Collection
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_Carrega_Lista_Almoxarifado
    
    'Lê Códigos e NomesReduzidos da tabela Almoxarifado e devolve na coleção
    lErro = CF("Almoxarifados_Le_FilialEmpresa", giFilialEmpresa, colAlmoxarifados)
    If lErro <> SUCESSO Then gError 171806

    'Preenche a ListBox AlmoxarifadoList com os objetos da coleção
    For Each objAlmoxarifado In colAlmoxarifados
        Almoxarifados.AddItem objAlmoxarifado.iCodigo & SEPARADOR & objAlmoxarifado.sNomeReduzido
        Almoxarifados.ItemData(Almoxarifados.NewIndex) = objAlmoxarifado.iCodigo
    Next

    Carrega_Lista_Almoxarifado = SUCESSO

    Exit Function

Erro_Carrega_Lista_Almoxarifado:

    Carrega_Lista_Almoxarifado = gErr

    Select Case gErr

        Case 171806

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181827)

    End Select

    Exit Function

End Function

Private Sub LabelLoteInicial_Click()

Dim colSelecao As New Collection
Dim objRastroLote As New ClassRastreamentoLote

    objRastroLote.sCodigo = LoteInicial.Text

    'Chama tela de Browse de RastreamentoLote
    Call Chama_Tela("RastroLoteLista1", colSelecao, objRastroLote, objEventoLoteInicial)

End Sub

Private Sub objEventoLoteInicial_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRastroLote As ClassRastreamentoLote

On Error GoTo Erro_objEventoLoteInicial_evSelecao

    Set objRastroLote = obj1
    
    'Coloca produto na tela
    LoteInicial.PromptInclude = False
    LoteInicial.Text = objRastroLote.sCodigo
    LoteInicial.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoLoteInicial_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181828)

    End Select

    Exit Sub

End Sub

Private Sub LabelLoteFinal_Click()

Dim colSelecao As New Collection
Dim objRastroLote As New ClassRastreamentoLote

    objRastroLote.sCodigo = LoteFinal.Text

    'Chama tela de Browse de RastreamentoLote
    Call Chama_Tela("RastroLoteLista1", colSelecao, objRastroLote, objEventoLoteFinal)

End Sub

Private Sub objEventoLoteFinal_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRastroLote As ClassRastreamentoLote

On Error GoTo Erro_objEventoLoteFinal_evSelecao

    Set objRastroLote = obj1
    
    'Coloca produto na tela
    LoteFinal.PromptInclude = False
    LoteFinal.Text = objRastroLote.sCodigo
    LoteFinal.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoLoteFinal_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181829)

    End Select

    Exit Sub

End Sub

Private Sub FilialOP_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_FilialOP_Validate

    'Se não estiver preenchida ou alterada pula a crítica
    If Len(Trim(FilialOP.Text)) = 0 Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialOP, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 171807

    'Nao encontrou o item com o código informado
    If lErro = 6730 Then gError 171808

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 171809

    Exit Sub

Erro_FilialOP_Validate:

    Cancel = True

    Select Case gErr

        Case 171807

        Case 171808
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, iCodigo)

        Case 171809
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, FilialOP.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181830)

    End Select

    Exit Sub

End Sub

Private Function Carrega_FilialOP() As Long
'Carrega FilialOP a partir das Filiais Empresas

Dim lErro As Long
Dim objFiliais As AdmFiliais

On Error GoTo Erro_Carrega_FilialOP

    For Each objFiliais In gcolFiliais

        If objFiliais.iCodFilial <> EMPRESA_TODA Then
        
            FilialOP.AddItem CStr(objFiliais.iCodFilial) & SEPARADOR & objFiliais.sNome
            FilialOP.ItemData(FilialOP.NewIndex) = objFiliais.iCodFilial
    
        End If
        
    Next
        
    Carrega_FilialOP = SUCESSO

    Exit Function

Erro_Carrega_FilialOP:

    Carrega_FilialOP = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181831)

    End Select

    Exit Function

End Function
'####################################################

Private Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Vendedor_Validate

    If Len(Trim(Vendedor.Text)) > 0 Then
   
        'Tenta ler o vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le2(Vendedor, objVendedor, 0)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
    
    Exit Sub

Erro_Vendedor_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO2", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169098)

    End Select

End Sub

Private Sub LabelVendedor_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection
    
    'Preenche com o Vendedor da tela
    objVendedor.iCodigo = Codigo_Extrai(Vendedor.Text)
    
    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1
    
    'Preenche campo Vendedor
    Vendedor.Text = CStr(objVendedor.iCodigo)
    Call Vendedor_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub LabelCidade_Click()

Dim objCidade As New ClassCidades
Dim colSelecao As Collection

    objCidade.sDescricao = Cidade.Text

    'Chama a Tela de browse
    Call Chama_Tela("CidadeLista", colSelecao, objCidade, objEventoCidade)

End Sub

Private Sub objEventoCidade_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCidade As ClassCidades

On Error GoTo Erro_objEventoCidade_evSelecao

    Set objCidade = obj1

    Cidade.Text = objCidade.sDescricao

    Me.Show

    Exit Sub

Erro_objEventoCidade_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202974)

    End Select

    Exit Sub

End Sub

Private Sub Cidade_Validate(Cancel As Boolean)

Dim lErro As Long, objCidade As New ClassCidades
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Cidade_Validate

    If Len(Trim(Cidade.Text)) = 0 Then Exit Sub

    objCidade.sDescricao = Cidade.Text
    
    lErro = CF("Cidade_Le_Nome", objCidade)
    If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError ERRO_SEM_MENSAGEM

    If lErro <> SUCESSO Then gError 202976

    Exit Sub

Erro_Cidade_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 202976
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CIDADE")
            If vbMsgRes = vbYes Then
                Call Chama_Tela("CidadeCadastro", objCidade)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202977)

    End Select

    Exit Sub

End Sub

Private Sub LabelRegiaoAte_Click()
    
Dim objRegiaoVenda As New ClassRegiaoVenda
Dim colSelecao As New Collection
    
    giRegiaoVenda = 0
    
    'Se o tipo está preenchido
    If Len(Trim(RegiaoFinal.Text)) > 0 Then
        
        'Preenche com o tipo da tela
        objRegiaoVenda.iCodigo = CInt(RegiaoFinal.Text)
    
    End If
    
    'Chama Tela RegiãoVendaLista
    Call Chama_Tela("RegiaoVendaLista", colSelecao, objRegiaoVenda, objEventoRegiaoVenda)
    
End Sub

Private Sub LabelRegiaoDe_Click()

Dim objRegiaoVenda As New ClassRegiaoVenda
Dim colSelecao As New Collection
    
    giRegiaoVenda = 1
    
    'Se o tipo está preenchido
    If Len(Trim(RegiaoInicial.Text)) > 0 Then
        
        'Preenche com o tipo da tela
        objRegiaoVenda.iCodigo = CInt(RegiaoInicial.Text)
        
    End If
    
    'Chama Tela RegiãoVendaLista
    Call Chama_Tela("RegiaoVendaLista", colSelecao, objRegiaoVenda, objEventoRegiaoVenda)

End Sub

Private Sub objEventoRegiaoVenda_evSelecao(obj1 As Object)

Dim objRegiaoVenda As New ClassRegiaoVenda

    Set objRegiaoVenda = obj1
    
    'Preenche campo Tipo de produto
    If giRegiaoVenda = 1 Then
        RegiaoInicial.Text = objRegiaoVenda.iCodigo
        RegiaoDe.Caption = objRegiaoVenda.sDescricao
    Else
        RegiaoFinal.Text = objRegiaoVenda.iCodigo
        RegiaoAte.Caption = objRegiaoVenda.sDescricao
    End If

    Me.Show

    Exit Sub

End Sub


Private Sub RegiaoFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objRegiaoVenda As New ClassRegiaoVenda

On Error GoTo Erro_RegiaoFinal_Validate

    giRegiaoVenda = 0
                                
    lErro = RegiaoVenda_Perde_Foco(RegiaoFinal, RegiaoAte)
    If lErro <> SUCESSO And lErro <> 87199 Then gError 87202
       
    If lErro = 87199 Then gError 87203
        
    Exit Sub

Erro_RegiaoFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 87202
        
        Case 87203
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_CADASTRADA", gErr, objRegiaoVenda.iCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169083)

    End Select

    Exit Sub

End Sub

Private Sub RegiaoInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sRegiaoInicial As String
Dim objRegiaoVenda As New ClassRegiaoVenda

On Error GoTo Erro_RegiaoInicial_Validate

    giRegiaoVenda = 1
                
    lErro = RegiaoVenda_Perde_Foco(RegiaoInicial, RegiaoDe)
    If lErro <> SUCESSO And lErro <> 87199 Then gError 87200
       
    If lErro = 87199 Then gError 87201
    
    Exit Sub

Erro_RegiaoInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 87200
        
        Case 87201
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_CADASTRADA", gErr, objRegiaoVenda.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169084)

    End Select

    Exit Sub

End Sub

Public Function RegiaoVenda_Perde_Foco(Regiao As Object, Desc As Object) As Long
'recebe MaskEdBox da Região de Venda e o label da descrição

Dim lErro As Long
Dim objRegiaoVenda As New ClassRegiaoVenda

On Error GoTo Erro_RegiaoVenda_Perde_Foco

        
    If Len(Trim(Regiao.Text)) > 0 Then
        
            objRegiaoVenda.iCodigo = StrParaInt(Regiao.Text)
        
            lErro = CF("RegiaoVenda_Le", objRegiaoVenda)
            If lErro <> SUCESSO And lErro <> 16137 Then gError 87198
        
            If lErro = 16137 Then gError 87199

        Desc.Caption = objRegiaoVenda.sDescricao

    Else

        Desc.Caption = ""

    End If

    RegiaoVenda_Perde_Foco = SUCESSO

    Exit Function

Erro_RegiaoVenda_Perde_Foco:

    RegiaoVenda_Perde_Foco = gErr

    Select Case gErr

        Case 87198
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGIOESVENDAS", gErr, objRegiaoVenda.iCodigo)

        Case 87199
            'Erro tratado na rotina chamadora
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169086)

    End Select

    Exit Function

End Function

