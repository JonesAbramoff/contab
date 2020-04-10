VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpRealPrev 
   ClientHeight    =   6930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   KeyPreview      =   -1  'True
   ScaleHeight     =   6930
   ScaleWidth      =   6765
   Begin VB.CheckBox EmpresaToda 
      Caption         =   "Consolidar Empresa Toda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2280
      TabIndex        =   2
      Top             =   835
      Width           =   2595
   End
   Begin VB.Frame FrameData 
      Caption         =   "Período"
      Height          =   720
      Left            =   150
      TabIndex        =   43
      Top             =   1320
      Width           =   4092
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   330
         Left            =   1590
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   312
         Left            =   624
         TabIndex        =   3
         Top             =   252
         Width           =   972
         _ExtentX        =   1693
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   330
         Left            =   3435
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   312
         Left            =   2460
         TabIndex        =   4
         Top             =   255
         Width           =   972
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label dIni 
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
         Height          =   252
         Left            =   240
         TabIndex        =   47
         Top             =   300
         Width           =   396
      End
      Begin VB.Label dFim 
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
         Height          =   252
         Left            =   2040
         TabIndex        =   46
         Top             =   300
         Width           =   456
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Região"
      Height          =   1245
      Left            =   150
      TabIndex        =   38
      Top             =   2160
      Width           =   6285
      Begin MSMask.MaskEdBox RegiaoInicial 
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Top             =   315
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox RegiaoFinal 
         Height          =   315
         Left            =   1065
         TabIndex        =   7
         Top             =   765
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin VB.Label RegiaoAte 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   2505
         TabIndex        =   42
         Top             =   765
         Width           =   3120
      End
      Begin VB.Label RegiaoDe 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   2505
         TabIndex        =   41
         Top             =   315
         Width           =   3120
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
         Left            =   720
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   40
         Top             =   360
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
         Left            =   660
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   39
         Top             =   810
         Width           =   435
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Vendedores"
      Height          =   840
      Left            =   120
      TabIndex        =   34
      Top             =   7200
      Visible         =   0   'False
      Width           =   5775
      Begin MSMask.MaskEdBox VendedorInicial 
         Height          =   300
         Left            =   585
         TabIndex        =   20
         Top             =   255
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox VendedorFinal 
         Height          =   300
         Left            =   3240
         TabIndex        =   21
         Top             =   255
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelVendedorAte 
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
         Left            =   2850
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   36
         Top             =   315
         Width           =   360
      End
      Begin VB.Label LabelVendedorDe 
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
         Left            =   195
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   35
         Top             =   300
         Width           =   315
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clientes"
      Height          =   720
      Left            =   150
      TabIndex        =   31
      Top             =   3525
      Width           =   6285
      Begin MSMask.MaskEdBox ClienteInicial 
         Height          =   300
         Left            =   1080
         TabIndex        =   8
         Top             =   255
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ClienteFinal 
         Height          =   300
         Left            =   3720
         TabIndex        =   9
         Top             =   255
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
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
         Left            =   675
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   33
         Top             =   270
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3315
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   32
         Top             =   315
         Width           =   360
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Filiais"
      Height          =   840
      Left            =   150
      TabIndex        =   27
      Top             =   5880
      Width           =   6285
      Begin VB.ComboBox FilialEmpresaFinal 
         Height          =   315
         Left            =   3600
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   345
         Width           =   2040
      End
      Begin VB.ComboBox FilialEmpresaInicial 
         Height          =   315
         Left            =   960
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   360
         Width           =   2040
      End
      Begin VB.Label Label6 
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
         Left            =   3180
         TabIndex        =   29
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   570
         TabIndex        =   28
         Top             =   390
         Width           =   315
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpRealPrevInpal.ctx":0000
      Left            =   900
      List            =   "RelOpRealPrevInpal.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   405
      Width           =   2916
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
      Left            =   4920
      Picture         =   "RelOpRealPrevInpal.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   840
      Width           =   1575
   End
   Begin VB.CheckBox Devolucoes 
      Caption         =   "Inclui Devoluções"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4560
      TabIndex        =   5
      Top             =   1560
      Width           =   1875
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produtos"
      Height          =   1290
      Left            =   150
      TabIndex        =   22
      Top             =   4440
      Width           =   6285
      Begin MSMask.MaskEdBox ProdutoInicial 
         Height          =   315
         Left            =   1110
         TabIndex        =   10
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
         Left            =   1110
         TabIndex        =   11
         Top             =   825
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
         Left            =   2655
         TabIndex        =   26
         Top             =   825
         Width           =   3000
      End
      Begin VB.Label DescProdInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2685
         TabIndex        =   25
         Top             =   360
         Width           =   2970
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
         Left            =   720
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   24
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
         Left            =   675
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   23
         Top             =   870
         Width           =   435
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4440
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   165
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1620
         Picture         =   "RelOpRealPrevInpal.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpRealPrevInpal.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpRealPrevInpal.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpRealPrevInpal.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   300
      Left            =   885
      TabIndex        =   1
      Top             =   885
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      PromptChar      =   " "
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
      Left            =   240
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   37
      Top             =   960
      Width           =   660
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
      Left            =   210
      TabIndex        =   30
      Top             =   450
      Width           =   615
   End
End
Attribute VB_Name = "RelOpRealPrev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1

Dim giVendedorInicial As Integer
Dim giClienteInicial As Integer
Dim giRegiaoVenda As Integer


'Novos IDHs
Const IDH_RELOP_REAL_PREVISTOVENDCLI = 0

'Eventos dos Browses
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoPrevVenda As AdmEvento
Attribute objEventoPrevVenda.VB_VarHelpID = -1
Private WithEvents objEventoRegiaoVenda As AdmEvento
Attribute objEventoRegiaoVenda.VB_VarHelpID = -1

Public Sub Form_Load()

Dim lErro As Long
Dim iOpcao As Integer
Dim colCategoriaProduto As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto

On Error GoTo Erro_Form_Load
                     
    Set objEventoVendedor = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoPrevVenda = New AdmEvento
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    Set objEventoRegiaoVenda = New AdmEvento
    
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoInicial)
    If lErro <> SUCESSO Then gError 500227

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoFinal)
    If lErro <> SUCESSO Then gError 500228
  
    'Preenche as combos de filial Empresa guardando no itemData o codigo
    lErro = Carrega_FilialEmpresa()
    If lErro <> SUCESSO Then gError 500230
    
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 500231
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 500227, 500228, 500230, 500231

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar()
    If lErro <> SUCESSO Then gError 500232
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then gError 90164

    Call DateParaMasked(DataInicial, CDate(sParam))
    
    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 90165

    Call DateParaMasked(DataFinal, CDate(sParam))
    
    'pega Região de Venda Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TREGIAOVENDAINIC", sParam)
    If lErro Then gError 90119
    RegiaoInicial.Text = sParam
    Call RegiaoInicial_Validate(bSGECancelDummy)
        
    'pega Região de Venda Final e exibe
    lErro = objRelOpcoes.ObterParametro("TREGIAOVENDAFIM", sParam)
    If lErro Then gError 90120
    RegiaoFinal.Text = sParam
    Call RegiaoFinal_Validate(bSGECancelDummy)
    
    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro <> SUCESSO Then gError 500233

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 500234

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 500235

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 500236
    
    If giFilialEmpresa <> EMPRESA_TODA Then
        
        'Preenche em Branco
        FilialEmpresaInicial.ListIndex = -1
        FilialEmpresaFinal.ListIndex = -1
        
        'desabilita a combo
        FilialEmpresaInicial.Enabled = False
        FilialEmpresaFinal.Enabled = False
        
    Else
        
        'pega parâmetro FilialEmpresa Inicial
        lErro = objRelOpcoes.ObterParametro("NFILIALINIC", sParam)
        If lErro <> SUCESSO Then gError 500237
         
        FilialEmpresaInicial.Text = sParam
        Call FilialEmpresaInicial_Validate(bSGECancelDummy)
             
        'pega parâmetro FilialEmpresa Final
        lErro = objRelOpcoes.ObterParametro("NFILIALFIM", sParam)
        If lErro <> SUCESSO Then gError 500238
        
        FilialEmpresaFinal.Text = sParam
        Call FilialEmpresaFinal_Validate(bSGECancelDummy)
         
    End If
    
    'pega código e exibe
    lErro = objRelOpcoes.ObterParametro("TCODIGO", sParam)
    If lErro <> SUCESSO Then gError 500239

    Codigo.Text = sParam
            
    'Pega  Empresa Toda
    lErro = objRelOpcoes.ObterParametro("TEMPRESATODA", sParam)
    If lErro <> SUCESSO Then gError 90405
    
    If sParam = "1" Then
        EmpresaToda.Value = 1
    Else
        If sParam = "0" Then EmpresaToda.Value = 0
    End If
            
    'pega parametro de devolução e exibe
    lErro = objRelOpcoes.ObterParametro("NDEVOLUCAO", sParam)
    If lErro <> SUCESSO Then gError 500241
    
    If sParam <> "" Then Devolucoes.Value = CInt(sParam)
                    
    'pega Cliente inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCLIENTEINIC", sParam)
    If lErro Then gError 500242
    
    ClienteInicial.Text = sParam
    Call ClienteInicial_Validate(bSGECancelDummy)
    
    'pega  Cliente final e exibe
    lErro = objRelOpcoes.ObterParametro("NCLIENTEFIM", sParam)
    If lErro Then gError 500243
    
    ClienteFinal.Text = sParam
    Call ClienteFinal_Validate(bSGECancelDummy)
          
    'pega vendedor inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NVENDEDORINIC", sParam)
    If lErro Then gError 500244
    
    VendedorInicial.Text = sParam
    Call VendedorInicial_Validate(bSGECancelDummy)
    
    'pega  vendedor final e exibe
    lErro = objRelOpcoes.ObterParametro("NVENDEDORFIM", sParam)
    If lErro Then gError 500245
    
    VendedorFinal.Text = sParam
    Call VendedorFinal_Validate(bSGECancelDummy)
          
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 500232 To 500245
        
        Case 90164, 90165, 90405

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 500246
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes
    
    Caption = gobjRelatorio.sCodRel
    
    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 500247

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
                
        Case 500246
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case 500247
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
        
    Set objEventoCliente = Nothing
    Set objEventoVendedor = Nothing
    Set objEventoPrevVenda = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    Set objEventoRegiaoVenda = Nothing

End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82642

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82643
    
    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 82644

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 82642, 82644

        Case 82643
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82645

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82646

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 82647

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 82645, 82647

        Case 82646
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

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
        If lErro <> SUCESSO Then gError 82649

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 82649

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

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
        If lErro <> SUCESSO Then gError 82650

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 82650

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 500248
    
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 500249
        
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 500248, 500249
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iFilialEmpresa As Integer

On Error GoTo Erro_Codigo_Validate

    'Se o código foi preenchido
    If Len(Trim(Codigo.Text)) > 0 Then
    
        'Verifica se houve escolha por consolidar Empresa_Toda
        If EmpresaToda.Value = 0 Then
            iFilialEmpresa = giFilialEmpresa
        ElseIf EmpresaToda.Value = 1 Then
            iFilialEmpresa = EMPRESA_TODA
        End If
    
        'Verifica se existe uma Previsão de Vendas cadastrada com o código passado
        lErro = PrevVendaMensal_Le_Codigo(Codigo.Text, giFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 90203 Then gError 500291
        
        'Se não encontro PrevVenda, erro
        If lErro = 90203 Then gError 500292
        
    End If
    
    Exit Sub
    
Erro_Codigo_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 500291
        
        Case 500292
            Call Rotina_Erro(vbOKOnly, "ERRO_PREVVENDA_NAO_CADASTRADA", gErr, Codigo.Text)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select
    
    Exit Sub
    

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sProd_I As String
Dim sProd_F As String
Dim iIndice As Integer
Dim sFilial_I As String
Dim sFilial_F As String
Dim sCliente_I As String
Dim sCliente_F As String
Dim sVend_I As String
Dim sVend_F As String
Dim iFilialEmpresa As Integer
Dim sCheckEmpToda As String

On Error GoTo Erro_PreencherRelOp

    'Se o código não foi preenchido, erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 500282
    
    If gobjRelatorio.sCodRel = "Acompanhamento de Orçamento" Then
    
         'Verifica se houve escolha por consolidar Empresa_Toda
        If EmpresaToda.Value = 0 Then
            iFilialEmpresa = giFilialEmpresa
            sCheckEmpToda = "0"
            gobjRelatorio.sNomeTsk = "AcomOrca"
        ElseIf EmpresaToda.Value = 1 Then
            iFilialEmpresa = EMPRESA_TODA
            sCheckEmpToda = "1"
            gobjRelatorio.sNomeTsk = "AcOrcaET"
        End If
    
    End If
    
    'Pode Verifica se existe uma Previsão Mensal de Vendas cadastrada com o código passado
    lErro = PrevVendaMensal_Le_Codigo(Codigo.Text, iFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 90203 Then gError 90411
    
    'Se não encontro PrevVenda, erro
    If lErro = 90203 Then gError 90396
    
    'Se a data inicial não foi preenchida, erro
    If Len(DataInicial.ClipText) = 0 Then gError 90298
    
    'Se a data final não foi preenchida, erro
    If Len(DataFinal.ClipText) = 0 Then gError 90299

    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)
       
    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F, sFilial_I, sFilial_F, sCliente_I, sCliente_F, sVend_I, sVend_F)
    If lErro <> SUCESSO Then gError 500250
    
    lErro = objRelOpcoes.Limpar()
    If lErro <> AD_BOOL_TRUE Then gError 500251
    
    lErro = objRelOpcoes.IncluirParametro("TCODIGO", Codigo.Text)
    If lErro <> AD_BOOL_TRUE Then gError 500258
    
    lErro = objRelOpcoes.IncluirParametro("TEMPRESATODA", sCheckEmpToda)
    If lErro <> AD_BOOL_TRUE Then gError 90412
    
    If Trim(DataInicial.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINIC", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 90134
    
    If Trim(DataFinal.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 90135
    
    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then gError 500252

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then gError 500253
             
    lErro = objRelOpcoes.IncluirParametro("NFILIALINIC", sFilial_I)
    If lErro <> AD_BOOL_TRUE Then gError 500254
    
    lErro = objRelOpcoes.IncluirParametro("TFILIALINIC", FilialEmpresaInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 500255

    lErro = objRelOpcoes.IncluirParametro("NFILIALFIM", sFilial_F)
    If lErro <> AD_BOOL_TRUE Then gError 500256
    
    lErro = objRelOpcoes.IncluirParametro("TFILIALFIM", FilialEmpresaFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 500257
    
    lErro = objRelOpcoes.IncluirParametro("NDEVOLUCAO", CInt(Devolucoes.Value))
    If lErro <> AD_BOOL_TRUE Then gError 500260
                      
    lErro = objRelOpcoes.IncluirParametro("NVENDEDORINIC", sVend_I)
    If lErro <> AD_BOOL_TRUE Then gError 500262
    
    lErro = objRelOpcoes.IncluirParametro("TVENDEDORINIC", VendedorInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 500263

    lErro = objRelOpcoes.IncluirParametro("NVENDEDORFIM", sVend_F)
    If lErro <> AD_BOOL_TRUE Then gError 500264
    
    lErro = objRelOpcoes.IncluirParametro("TVENDEDORFIM", VendedorFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 500265
         
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEINIC", sCliente_I)
    If lErro <> AD_BOOL_TRUE Then gError 500266
    
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEINIC", ClienteInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 500267

    lErro = objRelOpcoes.IncluirParametro("NCLIENTEFIM", sCliente_F)
    If lErro <> AD_BOOL_TRUE Then gError 500268
    
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEFIM", ClienteFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 500269
       
    lErro = objRelOpcoes.IncluirParametro("TREGIAOVENDAINIC", RegiaoInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 90166

    lErro = objRelOpcoes.IncluirParametro("TREGIAOVENDAFIM", RegiaoFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 90167
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd_I, sProd_F, sFilial_I, sFilial_F, sCliente_I, sCliente_F, sVend_I, sVend_F)
    If lErro <> SUCESSO Then gError 500261
            
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 500250 To 500269
        
        Case 90166, 90167, 90134, 90135, 90411, 90412

        Case 500282
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus
            
        Case 90298
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_NAO_PREENCHIDA", gErr)
                    
        Case 90299
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAFINAL_NAO_PREENCHIDA", gErr)
                      
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click
    
    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 500270

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 500271

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError 500272
        
        lErro = Define_Padrao()
        If lErro <> SUCESSO Then gError 500273
            
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 500270
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 500271, 500272, 500273

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 500274

'    If giFilialEmpresa <> EMPRESA_TODA Then gobjRelatorio.sNomeTsk = "FatReIn"
'    If giFilialEmpresa = EMPRESA_TODA Then gobjRelatorio.sNomeTsk = "FatReEIn"

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 500274

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 500275

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 500276

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 500277

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 500278
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 500275
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 500276, 500277, 500278

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd_I As String, sProd_F As String, sFilial_I As String, sFilial_F As String, sCliente_I As String, sCliente_F As String, sVend_I As String, sVend_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long
Dim iRegVendaInc As Integer
Dim iRegVendaFin As Integer

On Error GoTo Erro_Monta_Expressao_Selecao

    iRegVendaInc = Codigo_Extrai(RegiaoInicial.Text)
    iRegVendaFin = Codigo_Extrai(RegiaoFinal.Text)

    If Len(Trim(RegiaoInicial.Text)) <> 0 Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "RegiaoVenda >= " & Forprint_ConvInt(iRegVendaInc)
    
    End If
    
    If Len(Trim(RegiaoFinal.Text)) <> 0 Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "RegiaoVenda <= " & Forprint_ConvInt(iRegVendaFin)

    End If


    If sProd_I <> "" Then sExpressao = "Produto >= " & Forprint_ConvTexto(sProd_I)

    If sProd_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Produto <= " & Forprint_ConvTexto(sProd_F)

    End If
        
    If sFilial_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilialEmpresa >= " & Forprint_ConvInt(CInt(sFilial_I))

    End If
    
    If sFilial_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilialEmpresa <= " & Forprint_ConvInt(CInt(sFilial_F))

    End If
    
    If sVend_I <> "" Then
   
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = "Vendedor >= " & Forprint_ConvInt(CInt(sVend_I))
        
    End If
    
    If sVend_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Vendedor <= " & Forprint_ConvInt(CInt(sVend_F))

    End If
   
   
    If sCliente_I <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Cliente >= " & Forprint_ConvInt(CInt(sCliente_I))
        
    End If

    If sCliente_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Cliente <= " & Forprint_ConvInt(CInt(sCliente_F))

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String, sFilial_I As String, sFilial_F As String, sCliente_I, sCliente_F, sVend_I, sVend_F) As Long
'Formata os produtos retornando em sProd_I e sProd_F
'Verifica se os parâmetros iniciais são maiores que os finais

Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim lErro As Long
Dim sExpressao As String

On Error GoTo Erro_Formata_E_Critica_Parametros


    'Critica data se  a Inicial e a Final estiverem Preenchida
    If Len(DataInicial.ClipText) <> 0 And Len(DataFinal.ClipText) <> 0 Then
    
        'data inicial não pode ser maior que a data final
        If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then gError 90168
    
        If Year(DataInicial.Text) <> Year(DataFinal.Text) Then gError 90169
    
    End If

    'Se RegiãoInicial e RegiãoFinal estão preenchidos
    If Len(Trim(RegiaoInicial.Text)) > 0 And Len(Trim(RegiaoFinal.Text)) > 0 Then
    
        'Se Região inicial for maior que Região final, erro
        If CLng(RegiaoInicial.Text) > CLng(RegiaoFinal.Text) Then gError 90170
        
    End If
   
    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 500279

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 500280

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then gError 500281

    End If
        
    If giFilialEmpresa <> EMPRESA_TODA Then
        
        sFilial_I = ""
        sFilial_F = ""
        
    Else
    
        'critica FilialEmpresa Inicial e Final
        If FilialEmpresaInicial.ListIndex <> -1 Then
            sFilial_I = CStr(FilialEmpresaInicial.ItemData(FilialEmpresaInicial.ListIndex))
        Else
            sFilial_I = ""
        End If
        
        If FilialEmpresaFinal.ListIndex <> -1 Then
            sFilial_F = CStr(FilialEmpresaFinal.ItemData(FilialEmpresaFinal.ListIndex))
        Else
            sFilial_F = ""
        End If
                
        If sFilial_I <> "" And sFilial_F <> "" Then
            
            If CInt(sFilial_I) > CInt(sFilial_F) Then gError 500283
            
        End If
    
    End If
                                
    If sExpressao <> "" Then sExpressao = sExpressao & " E "
    sExpressao = sExpressao & "NDEVOLUCOES = " & Forprint_ConvInt(CInt(Devolucoes.Value))
                                
    'critica vendedor Inicial e Final
    If VendedorInicial.Text <> "" Then
        sVend_I = CStr(Codigo_Extrai(VendedorInicial.Text))
    Else
        sVend_I = ""
    End If
    
    If VendedorFinal.Text <> "" Then
        sVend_F = CStr(Codigo_Extrai(VendedorFinal.Text))
    Else
        sVend_F = ""
    End If
    
    If sVend_I <> "" And sVend_F <> "" Then
        
        If CInt(sVend_I) > CInt(sVend_F) Then gError 500284
        
    End If
   
    'critica Cliente Inicial e Final
    If ClienteInicial.Text <> "" Then
        sCliente_I = CStr(LCodigo_Extrai(ClienteInicial.Text))
    Else
        sCliente_I = ""
    End If
    
    If ClienteFinal.Text <> "" Then
        sCliente_F = CStr(LCodigo_Extrai(ClienteFinal.Text))
    Else
        sCliente_F = ""
    End If
            
    If sCliente_I <> "" And sCliente_F <> "" Then
        
        If CInt(sCliente_I) > CInt(sCliente_F) Then gError 500285
        
    End If
        
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                     
        Case 90168
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            
        Case 90169
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ANO_DIFERENTE", gErr)
        
        Case 90170
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGIAOVENDA_INICIAL_MAIOR", gErr)
        
        Case 500279
            ProdutoInicial.SetFocus

        Case 500280
            ProdutoFinal.SetFocus

        Case 500281
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            ProdutoInicial.SetFocus
            
        Case 500283
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_INICIAL_MAIOR", gErr)
            FilialEmpresaInicial.SetFocus
            
        Case 500284
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_INICIAL_MAIOR", gErr)
            VendedorInicial.SetFocus
                
        Case 500285
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)
            ClienteInicial.SetFocus
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Function

End Function

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate

    lErro = CF("Produto_Perde_Foco", ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 500296
    
    If lErro = 27095 Then gError 500297

    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 500296
            
        Case 500297
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    lErro = CF("Produto_Perde_Foco", ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 500298
    
    If lErro = 27095 Then gError 500299
    
    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 500298

        Case 500299
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub FilialEmpresaInicial_Validate(Cancel As Boolean)
'Busca a filial com código digitado na lista FilialEmpresa

Dim lErro As Long
Dim iCodigo As Integer
Dim iIndice As Integer

On Error GoTo Erro_FilialEmpresaInicial_Validate
    
    'se uma opcao da lista estiver selecionada, OK
    If FilialEmpresaInicial.ListIndex <> -1 Then Exit Sub
    
    If Len(Trim(FilialEmpresaInicial.Text)) = 0 Then Exit Sub
    
    lErro = Combo_Seleciona(FilialEmpresaInicial, iCodigo)
    If lErro <> SUCESSO Then gError 500300
        
    Exit Sub

Erro_FilialEmpresaInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 500300
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub FilialEmpresaFinal_Validate(Cancel As Boolean)
'Busca a filial com código digitado na lista FilialEmpresa

Dim lErro As Long
Dim iCodigo As Integer
Dim iIndice As Integer

On Error GoTo Erro_FilialEmpresaFinal_Validate
    
    'se uma opcao da lista estiver selecionada, OK
    If FilialEmpresaFinal.ListIndex <> -1 Then Exit Sub
    
    If Len(Trim(FilialEmpresaFinal.Text)) = 0 Then Exit Sub
    
    lErro = Combo_Seleciona(FilialEmpresaFinal, iCodigo)
    If lErro <> SUCESSO Then gError 500301
    
    Exit Sub

Erro_FilialEmpresaFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 500301
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Function Carrega_FilialEmpresa() As Long
'Carrega as Combos FilialEmpresaInicial e FilialEmpresaFinal

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim iIndice As Integer
Dim colCodigoDescricao As New AdmColCodigoNome

On Error GoTo Erro_Carrega_FilialEmpresa

    'Lê Códigos e NomesReduzidos da tabela FilialEmpresa e devolve na coleção
    lErro = CF("Cod_Nomes_Le", "FiliaisEmpresa", "FilialEmpresa", "Nome", STRING_FILIAL_NOME, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 500302
    
    'preenche as combos iniciais e finais
    For Each objCodigoNome In colCodigoDescricao
        
        If objCodigoNome.iCodigo <> 0 Then
            FilialEmpresaInicial.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
            FilialEmpresaInicial.ItemData(FilialEmpresaInicial.NewIndex) = objCodigoNome.iCodigo
    
            FilialEmpresaFinal.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
            FilialEmpresaFinal.ItemData(FilialEmpresaFinal.NewIndex) = objCodigoNome.iCodigo
        End If
    
    Next

    Carrega_FilialEmpresa = SUCESSO

    Exit Function

Erro_Carrega_FilialEmpresa:

    Carrega_FilialEmpresa = gErr

    Select Case gErr

        Case 500302

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Define_Padrao() As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Define_Padrao
        
    giClienteInicial = 1
    giVendedorInicial = 1
    giRegiaoVenda = 1
        
    If giFilialEmpresa <> EMPRESA_TODA Then

        FilialEmpresaInicial.ListIndex = -1
        FilialEmpresaFinal.ListIndex = -1
    
        FilialEmpresaInicial.Enabled = False
        FilialEmpresaFinal.Enabled = False
        
    Else
    
        FilialEmpresaInicial.ListIndex = -1
        FilialEmpresaFinal.ListIndex = -1
    
    End If
   
    ComboOpcoes.Text = ""
    Devolucoes.Value = 0
    EmpresaToda.Value = 0
    DescProdInic.Caption = ""
    DescProdFim.Caption = ""
    RegiaoAte.Caption = ""
    RegiaoDe.Caption = ""
    
    Define_Padrao = SUCESSO

    Exit Function

Erro_Define_Padrao:

    Define_Padrao = gErr

    Select Case gErr
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Sub VendedorInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_VendedorInicial_Validate

    If Len(Trim(VendedorInicial.Text)) > 0 Then
   
        'Tenta ler o vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le2(VendedorInicial, objVendedor, 0)
        If lErro <> SUCESSO Then gError 500303

    End If
    
    giVendedorInicial = 1
    
    Exit Sub

Erro_VendedorInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 500303

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    Exit Sub
    
End Sub

Private Sub VendedorFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_VendedorFinal_Validate

    If Len(Trim(VendedorFinal.Text)) > 0 Then

        'Tenta ler o vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le2(VendedorFinal, objVendedor, 0)
        If lErro <> SUCESSO Then gError 500304

    End If
    
    giVendedorInicial = 0
 
    Exit Sub

Erro_VendedorFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 500304
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    Exit Sub
    
End Sub

Private Sub ClienteInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteInicial_Validate

    If Len(Trim(ClienteInicial.Text)) > 0 Then
   
        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteInicial, objCliente, 0)
        If lErro <> SUCESSO Then gError 500305

    End If
    
    giClienteInicial = 1
    
    Exit Sub

Erro_ClienteInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 500305

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    Exit Sub
    
End Sub

Private Sub ClienteFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteFinal_Validate

    If Len(Trim(ClienteFinal.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteFinal, objCliente, 0)
        If lErro <> SUCESSO Then gError 500306

    End If
    
    giClienteInicial = 0
 
    Exit Sub

Erro_ClienteFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 500306

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

End Sub

Private Sub LabelClienteAte_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    giClienteInicial = 0
    
    If Len(Trim(ClienteFinal.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteFinal.Text)
    End If
    
    'Chama Tela ClienteLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub LabelClienteDe_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    giClienteInicial = 1

    If Len(Trim(ClienteInicial.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteInicial.Text)
    End If
    
    'Chama Tela ClienteLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub LabelVendedorAte_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection

    giVendedorInicial = 0
    
    If Len(Trim(VendedorFinal.Text)) > 0 Then
        'Preenche com o Vendedor da tela
        objVendedor.iCodigo = Codigo_Extrai(VendedorFinal.Text)
    End If
    
    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Private Sub LabelVendedorDe_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection

    giVendedorInicial = 1
    
    If Len(Trim(VendedorInicial.Text)) > 0 Then
        'Preenche com o Vendedor da tela
        objVendedor.iCodigo = Codigo_Extrai(VendedorInicial.Text)
    End If
    
    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    'Preenche campo Cliente
    If giClienteInicial = 1 Then
        ClienteInicial.Text = CStr(objCliente.lCodigo)
        Call ClienteInicial_Validate(bSGECancelDummy)
    Else
        ClienteFinal.Text = CStr(objCliente.lCodigo)
        Call ClienteFinal_Validate(bSGECancelDummy)
    End If

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1
    
    'Preenche campo Vendedor
    If giVendedorInicial = 1 Then
        VendedorInicial.Text = CStr(objVendedor.iCodigo)
        Call VendedorInicial_Validate(bSGECancelDummy)
    Else
        VendedorFinal.Text = CStr(objVendedor.iCodigo)
        Call VendedorFinal_Validate(bSGECancelDummy)
    End If

    Me.Show

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim objPrevVendaMensal As New ClassPrevVendaMensal
Dim colSelecao As Collection

    If Len(Trim(Codigo.Text)) > 0 Then
        
        'Preenche com o cliente da tela
        objPrevVendaMensal.sCodigo = Codigo.Text
    End If
    
    'Chama Tela ClienteLista
    Call Chama_Tela("PrevVMensalCodLista", colSelecao, objPrevVendaMensal, objEventoPrevVenda)

End Sub

Private Sub objEventoPrevVenda_evSelecao(obj1 As Object)

Dim objPrevVendaMensal As ClassPrevVendaMensal

    Set objPrevVendaMensal = obj1
    
    Codigo.Text = objPrevVendaMensal.sCodigo
    
    Me.Show

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is RegiaoInicial Then
            Call LabelRegiaoDe_Click
        ElseIf Me.ActiveControl Is RegiaoFinal Then
            Call LabelRegiaoAte_Click
        ElseIf Me.ActiveControl Is ClienteInicial Then
            Call LabelClienteDe_Click
        ElseIf Me.ActiveControl Is ClienteFinal Then
            Call LabelClienteAte_Click
        ElseIf Me.ActiveControl Is VendedorInicial Then
            Call LabelVendedorDe_Click
        ElseIf Me.ActiveControl Is VendedorFinal Then
            Call LabelVendedorAte_Click
        ElseIf Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        ElseIf Me.ActiveControl Is ProdutoInicial Then
            Call LabelProdutoDe_Click
        ElseIf Me.ActiveControl Is ProdutoFinal Then
            Call LabelProdutoAte_Click
        End If
    End If
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_REAL_PREVISTOVENDCLI
    Set Form_Load_Ocx = Me
    Caption = "Faturamento Real x Previsto"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpRealPrev"
    
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

Private Sub LabelRegiaoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelRegiaoDe, Source, X, Y)
End Sub

Private Sub LabelRegiaoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelRegiaoDe, Button, Shift, X, Y)
End Sub

Private Sub LabelRegiaoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelRegiaoAte, Source, X, Y)
End Sub

Private Sub LabelRegiaoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelRegiaoAte, Button, Shift, X, Y)
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

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
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

Private Sub LabelVendedorDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelVendedorDe, Source, X, Y)
End Sub

Private Sub LabelVendedorAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelVendedorAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigo, Source, X, Y)
End Sub

Private Sub LabelCodigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigo, Button, Shift, X, Y)
End Sub

Private Sub dFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dFim, Source, X, Y)
End Sub

Private Sub dFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dFim, Button, Shift, X, Y)
End Sub

Private Sub dIni_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dIni, Source, X, Y)
End Sub

Private Sub dIni_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dIni, Button, Shift, X, Y)
End Sub

Public Function RegiaoVenda_Perde_Foco(Regiao As Object, Desc As Object) As Long
'recebe MaskEdBox da Região de Venda e o label da descrição

Dim lErro As Long
Dim objRegiaoVenda As New ClassRegiaoVenda

On Error GoTo Erro_RegiaoVenda_Perde_Foco

        
    If Len(Trim(Regiao.Text)) > 0 Then
        
            objRegiaoVenda.iCodigo = StrParaInt(Regiao.Text)
        
            lErro = CF("RegiaoVenda_Le", objRegiaoVenda)
            If lErro <> SUCESSO And lErro <> 16137 Then gError 90171
        
            If lErro = 16137 Then gError 90172

        Desc.Caption = objRegiaoVenda.sDescricao

    Else

        Desc.Caption = ""

    End If

    RegiaoVenda_Perde_Foco = SUCESSO

    Exit Function

Erro_RegiaoVenda_Perde_Foco:

    RegiaoVenda_Perde_Foco = gErr

    Select Case gErr

        Case 90171
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGIOESVENDAS", gErr, objRegiaoVenda.iCodigo)

        Case 90172
            'Erro tratado na rotina chamadora
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Exit Function

End Function

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
    If lErro <> SUCESSO And lErro <> 90172 Then gError 90173
       
    If lErro = 90172 Then gError 90174
        
    Exit Sub

Erro_RegiaoFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 90173
        
        Case 90174
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_CADASTRADA", gErr, objRegiaoVenda.iCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

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
    If lErro <> SUCESSO And lErro <> 90172 Then gError 90175
       
    If lErro = 90172 Then gError 90176
    
    Exit Sub

Erro_RegiaoInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 90175
        
        Case 90176
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_CADASTRADA", gErr, objRegiaoVenda.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub DataFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFinal)

End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(DataFinal.ClipText) > 0 Then

        lErro = Data_Critica(DataFinal.Text)
        If lErro <> SUCESSO Then gError 90177

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 90177

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub DataInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicial)

End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    If Len(DataInicial.ClipText) > 0 Then

        lErro = Data_Critica(DataInicial.Text)
        If lErro <> SUCESSO Then gError 90178

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 90178

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 90179

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 90179
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 90180

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 90180
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 90181

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case gErr

        Case 90181
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 90182

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case gErr

        Case 90182
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

'Subir para RotinasFATUsu
'***Esta função já está também na RelOpDataRefOcx, RelOpPeriodoOcx, RelOpPrevVendaOcx, RelOpRankCliOcx
Function PrevVendaMensal_Le_Codigo(sCodigo As String, iFilialEmpresa As Integer) As Long
'Verifica se a previsão de Vendas Mensal de códio e FilialEmpresa passados existem

Dim lErro As Long
Dim iFilial As Integer
Dim lComando As Long

On Error GoTo Erro_PrevVendaMensal_Le_Codigo

    'Abertura de comandos
    lComando = Comando_Abrir()
    If lErro <> SUCESSO Then gError 90200
    
    If iFilialEmpresa = EMPRESA_TODA Then
    
        'Pesquisa no BD se existe a Previsão de Vendas Mensais com o código passado, para a Empresa toda
        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa FROM PrevVendaMensal WHERE Codigo = ? ", iFilial, sCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 90201
    Else
        'Pesquisa no BD se existe a Previsão de Vendas Mensais com o código passado, para uma FilialEmpresa
        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa FROM PrevVendaMensal WHERE Codigo = ? AND FilialEmpresa = ?", iFilial, sCodigo, iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 90201
    
    End If
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90202
    
    'PrevVendas não encontradas
    If lErro = AD_SQL_SEM_DADOS Then gError 90203
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)
    
    PrevVendaMensal_Le_Codigo = SUCESSO
    
    Exit Function
    
Erro_PrevVendaMensal_Le_Codigo:
    
    PrevVendaMensal_Le_Codigo = gErr
    
    Select Case gErr
        
        Case 90200
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 90201, 90202
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PREVVENDAMENSAL", gErr, sCodigo)
        
        Case 90203 'PrevVendas não cadastrada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
    
    End Select
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function



