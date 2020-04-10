VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpFatClienteProd 
   ClientHeight    =   7035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7920
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   7035
   ScaleWidth      =   7920
   Begin VB.Frame Frame6 
      Caption         =   "PF, PJ ou Ambos"
      Height          =   480
      Left            =   3075
      TabIndex        =   50
      Top             =   3765
      Width           =   4680
      Begin VB.OptionButton OptPFPJ 
         Caption         =   "Ambos"
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
         Left            =   3420
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   1080
      End
      Begin VB.OptionButton OptPFPJ 
         Caption         =   "Somente PJ"
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
         Left            =   1695
         TabIndex        =   11
         Top             =   240
         Width           =   1485
      End
      Begin VB.OptionButton OptPFPJ 
         Caption         =   "Somente PF"
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
         Left            =   210
         TabIndex        =   10
         Top             =   240
         Width           =   1485
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Vendedores"
      Height          =   675
      Left            =   120
      TabIndex        =   48
      Top             =   4260
      Width           =   7620
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   225
         Value           =   -1  'True
         Width           =   2340
      End
      Begin MSMask.MaskEdBox Vendedor 
         Height          =   300
         Left            =   5385
         TabIndex        =   15
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
         TabIndex        =   49
         Top             =   330
         Width           =   885
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Endereço"
      Height          =   540
      Left            =   3075
      TabIndex        =   46
      Top             =   3240
      Width           =   4665
      Begin MSMask.MaskEdBox Cidade 
         Height          =   315
         Left            =   855
         TabIndex        =   9
         Top             =   165
         Width           =   3645
         _ExtentX        =   6429
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
         Left            =   180
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   47
         Top             =   210
         Width           =   660
      End
   End
   Begin VB.Frame FrameCategoriaCliente 
      Caption         =   "Categoria dos Clientes"
      Height          =   1575
      Left            =   120
      TabIndex        =   43
      Top             =   4980
      Width           =   7425
      Begin VB.ListBox CategoriaClienteItens 
         Height          =   960
         Left            =   180
         Style           =   1  'Checkbox
         TabIndex        =   18
         Top             =   555
         Width           =   7065
      End
      Begin VB.ComboBox CategoriaCliente 
         Height          =   315
         Left            =   2145
         TabIndex        =   17
         Top             =   210
         Width           =   2745
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
         Left            =   195
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   15
         Left            =   360
         TabIndex        =   45
         Top             =   720
         Width           =   30
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
         Left            =   1230
         TabIndex        =   44
         Top             =   255
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5685
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpFatClienteProdPur.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpFatClienteProdPur.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpFatClienteProdPur.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpFatClienteProdPur.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
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
      Height          =   360
      Left            =   180
      TabIndex        =   19
      Top             =   6585
      Width           =   4125
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produtos"
      Height          =   1035
      Left            =   120
      TabIndex        =   37
      Top             =   2160
      Width           =   7605
      Begin MSMask.MaskEdBox ProdutoInicial 
         Height          =   315
         Left            =   735
         TabIndex        =   5
         Top             =   210
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
         Left            =   735
         TabIndex        =   6
         Top             =   615
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
         Left            =   315
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   41
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
         Left            =   345
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   40
         Top             =   255
         Width           =   315
      End
      Begin VB.Label DescProdInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2295
         TabIndex        =   39
         Top             =   210
         Width           =   5190
      End
      Begin VB.Label DescProdFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2295
         TabIndex        =   38
         Top             =   615
         Width           =   5205
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clientes"
      Height          =   615
      Left            =   135
      TabIndex        =   34
      Top             =   1500
      Width           =   7575
      Begin MSMask.MaskEdBox ClienteInicial 
         Height          =   300
         Left            =   630
         TabIndex        =   3
         Top             =   225
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ClienteFinal 
         Height          =   300
         Left            =   3255
         TabIndex        =   4
         Top             =   225
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
         Left            =   195
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   36
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
         Left            =   2835
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   35
         Top             =   300
         Width           =   360
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   750
      Left            =   120
      TabIndex        =   29
      Top             =   645
      Width           =   5505
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   1590
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   300
         Left            =   630
         TabIndex        =   1
         Top             =   285
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   315
         Left            =   4215
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   300
         Left            =   3240
         TabIndex        =   2
         Top             =   285
         Width           =   960
         _ExtentX        =   1693
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2835
         TabIndex        =   33
         Top             =   345
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   32
         Top             =   315
         Width           =   345
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpFatClienteProdPur.ctx":0994
      Left            =   885
      List            =   "RelOpFatClienteProdPur.ctx":0996
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   270
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
      Left            =   5940
      Picture         =   "RelOpFatClienteProdPur.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   810
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Filiais"
      Height          =   1005
      Left            =   120
      TabIndex        =   26
      Top             =   3240
      Width           =   2880
      Begin VB.ComboBox FilialEmpresaFinal 
         Height          =   315
         Left            =   720
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   585
         Width           =   2040
      End
      Begin VB.ComboBox FilialEmpresaInicial 
         Height          =   315
         Left            =   720
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   225
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
         Left            =   300
         TabIndex        =   28
         Top             =   630
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
         Left            =   330
         TabIndex        =   27
         Top             =   270
         Width           =   315
      End
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
      Left            =   195
      TabIndex        =   42
      Top             =   315
      Width           =   615
   End
End
Attribute VB_Name = "RelOpFatClienteProd"
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
Private WithEvents objEventoCidade As AdmEvento
Attribute objEventoCidade.VB_VarHelpID = -1
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1

Dim giClienteInicial As Integer
Dim giProdInicial As Integer

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1

Private Sub Form_Load()

Dim lErro As Long
Dim colCategoriaProduto As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto

On Error GoTo Erro_Form_Load

    Set objEventoCliente = New AdmEvento
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    Set objEventoCidade = New AdmEvento
    Set objEventoVendedor = New AdmEvento
    
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoInicial)
    If lErro <> SUCESSO Then Error 37806

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoFinal)
    If lErro <> SUCESSO Then Error 37807

    'Preenche as combos de filial Empresa guardando no itemData o codigo
    lErro = Carrega_FilialEmpresa()
    If lErro <> SUCESSO Then Error 47946
    
    Call Carrega_ComboCategoriaCliente(CategoriaCliente)
    
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then Error 48565
                  
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = Err

    Select Case Err

        Case 37806, 38807
        
        Case 47946, 48565
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168912)

    End Select

    Exit Sub

End Sub

Private Function Define_Padrao() As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Define_Padrao
    
    giClienteInicial = 1
    
    giProdInicial = 1
    
    Devolucoes.Value = 0
    
    If giFilialEmpresa <> EMPRESA_TODA Then

        FilialEmpresaInicial.ListIndex = -1
        FilialEmpresaFinal.ListIndex = -1
    
        FilialEmpresaInicial.Enabled = False
        FilialEmpresaFinal.Enabled = False
        
    Else
    
        FilialEmpresaInicial.ListIndex = -1
        FilialEmpresaFinal.ListIndex = -1
    
    End If
    
    CategoriaClienteTodas.Value = vbChecked
    CategoriaCliente.Enabled = False
    CategoriaClienteItens.Clear
    
    OptPFPJ(0).Value = True
    OptVendDir.Value = True
   
    Define_Padrao = SUCESSO

    Exit Function

Erro_Define_Padrao:

    Define_Padrao = Err

    Select Case Err
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168913)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String, iIndice As Integer
Dim sListCount As String, iIndiceRel As Integer

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 37810
   
    'pega Cliente inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCLIENTEINIC", sParam)
    If lErro <> SUCESSO Then Error 37811
    
    ClienteInicial.Text = sParam
    Call ClienteInicial_Validate(bSGECancelDummy)
    
    'pega  Cliente final e exibe
    lErro = objRelOpcoes.ObterParametro("NCLIENTEFIM", sParam)
    If lErro <> SUCESSO Then Error 37812
    
    ClienteFinal.Text = sParam
    Call ClienteFinal_Validate(bSGECancelDummy)
    
    'Pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro <> SUCESSO Then Error 37813

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then Error 37814

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro <> SUCESSO Then Error 37815

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then Error 37816
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then Error 37817

    Call DateParaMasked(DataInicial, CDate(sParam))
    'DataInicial.PromptInclude = False
    'DataInicial.Text = sParam
    'DataInicial.PromptInclude = True

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then Error 37818

    Call DateParaMasked(DataFinal, CDate(sParam))
    'DataFinal.PromptInclude = False
    'DataFinal.Text = sParam
    'DataFinal.PromptInclude = True
            
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
        If lErro <> SUCESSO Then Error 47604
         
        FilialEmpresaInicial.Text = sParam
        Call FilialEmpresaInicial_Validate(bSGECancelDummy)
             
        'pega parâmetro FilialEmpresa Final
        lErro = objRelOpcoes.ObterParametro("NFILIALFIM", sParam)
        If lErro <> SUCESSO Then Error 47605
        
        FilialEmpresaFinal.Text = sParam
        Call FilialEmpresaFinal_Validate(bSGECancelDummy)
         
    End If
    
    'pega parametro de devolução e exibe
    lErro = objRelOpcoes.ObterParametro("NDEVOLUCAO", sParam)
    If lErro <> SUCESSO Then Error 37819
    
    If sParam <> "" Then Devolucoes.Value = CInt(sParam)
    
    lErro = objRelOpcoes.ObterParametro("NVENDEDOR", sParam)
    If lErro <> SUCESSO Then Error 37819

    If StrParaInt(sParam) <> 0 Then
        Vendedor.Text = CInt(sParam)
        Call Vendedor_Validate(bSGECancelDummy)
    End If

    lErro = objRelOpcoes.ObterParametro("NTIPOVEND", sParam)
    If lErro <> SUCESSO Then Error 37819

    If StrParaInt(sParam) = VENDEDOR_DIRETO Then
        OptVendDir.Value = True
    Else
        OptVendIndir.Value = True
    End If
    
    lErro = objRelOpcoes.ObterParametro("NTIPOPFPJ", sParam)
    If lErro <> SUCESSO Then Error 37819

    OptPFPJ(StrParaInt(sParam)).Value = True
    
    lErro = objRelOpcoes.ObterParametro("TCIDADE", sParam)
    If lErro <> SUCESSO Then Error 37819

    Cidade.Text = sParam
    
    lErro = objRelOpcoes.ObterParametro("TCATEGORIA", sParam)
    If lErro <> SUCESSO Then Error 37819
    
    CategoriaCliente.Text = sParam
    Call CategoriaCliente_Validate(bSGECancelDummy)
    
    If Len(Trim(sParam)) > 0 Then
        CategoriaClienteTodas.Value = vbFalse
    Else
        CategoriaClienteTodas.Value = vbChecked
    End If
    
    'Limpa a Lista
    For iIndice = 0 To CategoriaClienteItens.ListCount - 1
        CategoriaClienteItens.Selected(iIndice) = False
    Next
    
    'Obtem o numero de Regioes selecionados na Lista
    lErro = objRelOpcoes.ObterParametro("NLISTCATCOUNT", sListCount)
    If lErro <> SUCESSO Then Error 37819
    'Percorre toda a Lista
    
    For iIndice = 0 To CategoriaClienteItens.ListCount - 1
        
        If sListCount = "0" Then
            CategoriaClienteItens.Selected(iIndice) = True
        Else
            'Percorre todas as Regieos que foram slecionados
            For iIndiceRel = 1 To StrParaInt(sListCount)
                lErro = objRelOpcoes.ObterParametro("TLISTCAT" & SEPARADOR & iIndiceRel, sParam)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                'Se o cliente não foi excluido
                If sParam = CategoriaClienteItens.List(iIndice) Then
                    'Marca as categorias que foram gravadas
                    CategoriaClienteItens.Selected(iIndice) = True
                End If
            Next
        End If
    Next
          
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 37810 To 37819

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168914)

    End Select

    Exit Function

End Function

Private Sub Form_Unload(Cancel As Integer)

    Set objEventoCliente = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    Set objEventoCidade = Nothing
    Set objEventoVendedor = Nothing
    
End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82540

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82541
    
    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 82542

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 82540, 82542

        Case 82541
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168915)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82543

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82544

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 82545

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 82543, 82545

        Case 82544
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168916)

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
        If lErro <> SUCESSO Then gError 82559

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 82559

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168917)

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
        If lErro <> SUCESSO Then gError 82558

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 82558

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168918)

    End Select

    Exit Sub

End Sub
Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 29884
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 37804
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 37804
        
        Case 29884
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168919)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 47141
    
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then Error 47948
    
    ComboOpcoes.Text = ""
    Devolucoes.Value = 0
    DescProdInic.Caption = ""
    DescProdFim.Caption = ""
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 47948
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168920)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional bExecutando As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sCliente_I As String
Dim sCliente_F As String
Dim sProd_I As String
Dim sProd_F As String
Dim iIndice As Integer
Dim sFilial_I As String
Dim sFilial_F As String
Dim lNumIntRelCat As Long, lNumIntRel As Long
Dim bTodasCat As Boolean, colCategorias As New Collection
Dim iTipoVend As Integer, iTipoPFPJ As Integer

On Error GoTo Erro_PreencherRelOp

    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)
       
    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F, sCliente_I, sCliente_F, sFilial_I, sFilial_F, iTipoVend, iTipoPFPJ)
    If lErro <> SUCESSO Then Error 37823
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 37824
    
    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then Error 37825

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then Error 37826
         
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEINIC", sCliente_I)
    If lErro <> AD_BOOL_TRUE Then Error 37827
    
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEINIC", ClienteInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54838

    lErro = objRelOpcoes.IncluirParametro("NCLIENTEFIM", sCliente_F)
    If lErro <> AD_BOOL_TRUE Then Error 37828
    
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEFIM", ClienteFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54839
    
    If Trim(DataInicial.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINIC", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 37829
    
    If Trim(DataFinal.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 37830
    
    lErro = objRelOpcoes.IncluirParametro("NFILIALINIC", sFilial_I)
    If lErro <> AD_BOOL_TRUE Then Error 47951
    
    lErro = objRelOpcoes.IncluirParametro("TFILIALINIC", FilialEmpresaInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54840

    lErro = objRelOpcoes.IncluirParametro("NFILIALFIM", sFilial_F)
    If lErro <> AD_BOOL_TRUE Then Error 47952
    
    lErro = objRelOpcoes.IncluirParametro("TFILIALFIM", FilialEmpresaFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54841
    
    lErro = objRelOpcoes.IncluirParametro("NDEVOLUCAO", CInt(Devolucoes.Value))
    If lErro <> AD_BOOL_TRUE Then Error 37831
    
    lErro = objRelOpcoes.IncluirParametro("TVENDEDOR", Vendedor.Text)
    If lErro <> AD_BOOL_TRUE Then Error 37831

    lErro = objRelOpcoes.IncluirParametro("NVENDEDOR", Codigo_Extrai(Vendedor.Text))
    If lErro <> AD_BOOL_TRUE Then Error 37831

    lErro = objRelOpcoes.IncluirParametro("NTIPOVEND", CStr(iTipoVend))
    If lErro <> AD_BOOL_TRUE Then Error 37831
    
    lErro = objRelOpcoes.IncluirParametro("NTIPOPFPJ", CStr(iTipoPFPJ))
    If lErro <> AD_BOOL_TRUE Then Error 37831
    
    lErro = objRelOpcoes.IncluirParametro("TCIDADE", Cidade.Text)
    If lErro <> AD_BOOL_TRUE Then Error 37831
    
    lErro = objRelOpcoes.IncluirParametro("TCATEGORIA", CategoriaCliente.Text)
    If lErro <> AD_BOOL_TRUE Then Error 37831

    bTodasCat = True
    For iIndice = 0 To CategoriaClienteItens.ListCount - 1
        If Not CategoriaClienteItens.Selected(iIndice) Then
            bTodasCat = False
            Exit For
        End If
    Next
    
    If Not bTodasCat Then
        'Percorre toda a Lista
        For iIndice = 0 To CategoriaClienteItens.ListCount - 1
            If CategoriaClienteItens.Selected(iIndice) = True Then
                colCategorias.Add CategoriaClienteItens.List(iIndice)
                'Inclui todas as Regioes que foram slecionados
                lErro = objRelOpcoes.IncluirParametro("TLISTCAT" & SEPARADOR & colCategorias.Count, CategoriaClienteItens.List(iIndice))
                If lErro <> AD_BOOL_TRUE Then Error 37831
            End If
        Next
    End If
    
    'Inclui o numero de Clientes selecionados na Lista
    lErro = objRelOpcoes.IncluirParametro("NLISTCATCOUNT", CStr(colCategorias.Count))
    If lErro <> AD_BOOL_TRUE Then Error 37831
    
    If bExecutando Then
    
        lErro = CF("RelFiltroFilCliCat_Prepara", CategoriaCliente.Text, colCategorias, lNumIntRelCat, StrParaLong(sCliente_I), StrParaLong(sCliente_F), Cidade.Text, Codigo_Extrai(Vendedor.Text), Codigo_Extrai(Vendedor.Text), iTipoVend, gsUsuario, 0, 0, iTipoPFPJ)
        If lErro <> SUCESSO Then Error 37831
    
    End If
    
    lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRelCat))
    If lErro <> AD_BOOL_TRUE Then Error 37831
       
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd_I, sProd_F, sCliente_I, sCliente_F, sFilial_I, sFilial_F)
    If lErro <> SUCESSO Then Error 37832
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 37823 To 37832
        
        Case 47951, 47952, 54838, 54839, 54840, 54841

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168921)

    End Select

    Exit Function

End Function


Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 37833

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 37834

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then Error 47144
        
        lErro = Define_Padrao()
        If lErro <> SUCESSO Then Error 47949
    
        ComboOpcoes.Text = ""
        Devolucoes.Value = 0
        DescProdInic.Caption = ""
        DescProdFim.Caption = ""
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 37833
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 37834, 47144, 47949

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168922)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then Error 37835

'    If giFilialEmpresa <> EMPRESA_TODA Then gobjRelatorio.sNomeTsk = "EstVend"
'    If giFilialEmpresa = EMPRESA_TODA Then gobjRelatorio.sNomeTsk = "EstVendE"
    
    Select Case UCase(gobjRelatorio.sNomeTsk)
    
        Case "ESTVEND7"
            If giFilialEmpresa = EMPRESA_TODA Then gobjRelatorio.sNomeTsk = "EstVend8"
    
        Case "ESTVEND"
            If giFilialEmpresa = EMPRESA_TODA Then gobjRelatorio.sNomeTsk = "EstVendE"
            
    End Select
    
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 37835

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168923)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 37836

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 37837

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 37838

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 47142
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 37836
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 37837, 37838, 47142

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168924)

    End Select

    Exit Sub

End Sub


Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd_I As String, sProd_F As String, sCliente_I As String, sCliente_F As String, sFilial_I As String, sFilial_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long


On Error GoTo Erro_Monta_Expressao_Selecao

   If sProd_I <> "" Then sExpressao = "Produto >= " & Forprint_ConvTexto(sProd_I)

    If sProd_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Produto <= " & Forprint_ConvTexto(sProd_F)

    End If

'   If sCliente_I <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Cliente >= " & Forprint_ConvLong(CLng(sCliente_I))
'
'   End If
'
'   If sCliente_F <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Cliente <= " & Forprint_ConvLong(CLng(sCliente_F))
'
'    End If
    
    If Trim(DataInicial.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(CDate(DataInicial.Text))

    End If
    
    If Trim(DataFinal.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(CDate(DataFinal.Text))

    End If
    
    If sFilial_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilialEmpresa >= " & Forprint_ConvInt(CInt(sFilial_I))

    End If
    
    If sFilial_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilialEmpresa <= " & Forprint_ConvInt(CInt(sFilial_F))

    End If
    
'    If sExpressao <> "" Then sExpressao = sExpressao & " E "
'    sExpressao = sExpressao & "NDEVOLUCOES = " & Forprint_ConvInt(CInt(Devolucoes.Value))
     
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168925)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String, sCliente_I As String, sCliente_F As String, sFilial_I As String, sFilial_F As String, iTipoVend As Integer, iTipoPFPJ As Integer) As Long
'Formata os produtos retornando em sProd_I e sProd_F
'Verifica se os parâmetros iniciais são maiores que os finais

Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim lErro As Long, iIndice As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros
   
    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then Error 37839

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then Error 37840

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then Error 37841

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
        
        If CLng(sCliente_I) > CLng(sCliente_F) Then Error 37842
        
    End If
    
    'data inicial não pode ser maior que a data final
    If Trim(DataInicial.ClipText) <> "" And Trim(DataFinal.ClipText) <> "" Then
    
         If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then Error 37843
    
    End If
    
    If giFilialEmpresa <> EMPRESA_TODA Then
        
        sFilial_I = CStr(giFilialEmpresa)
        sFilial_F = CStr(giFilialEmpresa)
        
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
            
            If CInt(sFilial_I) > CInt(sFilial_F) Then Error 47950
            
        End If
    
    End If
    
    If OptVendDir.Value Then
        iTipoVend = VENDEDOR_DIRETO
    Else
        iTipoVend = VENDEDOR_INDIRETO
    End If
    
    For iIndice = 0 To 2
        If OptPFPJ(iIndice).Value Then
            iTipoPFPJ = iIndice
            Exit For
        End If
    Next
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function


Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = Err

    Select Case Err
                     
         Case 37839
            ProdutoInicial.SetFocus

        Case 37840
            ProdutoFinal.SetFocus

        Case 37841
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", Err)
            ProdutoInicial.SetFocus
       
        Case 37842
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", Err)
            ClienteInicial.SetFocus
        
         Case 37843
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)
            DataInicial.SetFocus
       
        Case 47950
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_INICIAL_MAIOR", Err)
            FilialEmpresaInicial.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168926)

    End Select

    Exit Function

End Function

Private Sub ClienteInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteInicial_Validate

    If Len(Trim(ClienteInicial.Text)) > 0 Then
   
        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteInicial, objCliente, 0)
        If lErro <> SUCESSO Then Error 37845

    End If
    
    giClienteInicial = 1
    
    Exit Sub

Erro_ClienteInicial_Validate:

    Cancel = True


    Select Case Err

        Case 37845
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO_2", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168927)

    End Select

End Sub

Private Sub ClienteFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteFinal_Validate

    If Len(Trim(ClienteFinal.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteFinal, objCliente, 0)
        If lErro <> SUCESSO Then Error 37846

    End If
    
    giClienteInicial = 0
 
    Exit Sub

Erro_ClienteFinal_Validate:

    Cancel = True


    Select Case Err

        Case 37846
             lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", Err, objCliente.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168928)

    End Select

End Sub

Private Sub DataFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFinal)

End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(DataFinal.ClipText) > 0 Then

        sDataFim = DataFinal.Text
        
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then Error 37847

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True


    Select Case Err

        Case 37847

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168929)

    End Select

    Exit Sub

End Sub

Private Sub DataInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicial)

End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    If Len(DataInicial.ClipText) > 0 Then

        sDataInic = DataInicial.Text
        
        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then Error 37848

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True


    Select Case Err

        Case 37848

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168930)

    End Select

    Exit Sub

End Sub

Private Sub LabelClienteAte_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    giClienteInicial = 0
    
    If Len(Trim(ClienteFinal.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteFinal.Text)
    End If
    
    'Chama Tela ClientesLista
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
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

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

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 37849

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 37849
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168931)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 37850

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 37850
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168932)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 37851

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case Err

        Case 37851
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168933)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 37852

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case Err

        Case 37852
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168934)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate

    giProdInicial = 0

    lErro = CF("Produto_Perde_Foco", ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO And lErro <> 27095 Then Error 37866
    
    If lErro <> SUCESSO Then Error 43241

    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True


    Select Case Err

        Case 37866

        Case 43241
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168935)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    giProdInicial = 1

    lErro = CF("Produto_Perde_Foco", ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then Error 37867
    
    If lErro <> SUCESSO Then Error 43242

    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True


    Select Case Err

        Case 37867

        Case 43242
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168936)

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
    If lErro <> SUCESSO Then Error 47947
    
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

    Carrega_FilialEmpresa = Err

    Select Case Err

        Case 47947

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168937)

    End Select

    Exit Function

End Function

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
    If lErro <> SUCESSO Then Error 47953
        
    Exit Sub

Erro_FilialEmpresaInicial_Validate:

    Cancel = True


    Select Case Err

        Case 47953
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168938)

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
    If lErro <> SUCESSO Then Error 47954
    
    Exit Sub

Erro_FilialEmpresaFinal_Validate:

    Cancel = True


    Select Case Err

        Case 47954
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168939)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_FAT_CLIENTE_PRODUTO
    Set Form_Load_Ocx = Me
    Caption = "Faturamento Cliente x Produto"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpFatClienteProd"
    
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
        
        If Me.ActiveControl Is ClienteInicial Then
            Call LabelClienteDe_Click
        ElseIf Me.ActiveControl Is ClienteFinal Then
            Call LabelClienteAte_Click
        ElseIf Me.ActiveControl Is ProdutoInicial Then
            Call LabelProdutoDe_Click
        ElseIf Me.ActiveControl Is ProdutoFinal Then
            Call LabelProdutoAte_Click
        End If
    
    End If

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

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub CategoriaClienteTodas_Click()

Dim lErro As Long

On Error GoTo Erro_CategoriaClienteTodas_Click

    If CategoriaClienteTodas.Value = vbChecked Then
        'Desabilita o combotipo
        CategoriaCliente.ListIndex = -1
        CategoriaCliente.Enabled = False
        CategoriaClienteItens.Clear
    Else
        CategoriaCliente.Enabled = True
    End If

    Call CategoriaCliente_Click

    Exit Sub

Erro_CategoriaClienteTodas_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168911)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaCliente_Click()

Dim lErro As Long

On Error GoTo Erro_CategoriaCliente_Click

    If Len(Trim(CategoriaCliente.Text)) > 0 Then
        CategoriaClienteItens.Enabled = True
        Call Carrega_ComboCategoriaItens(CategoriaCliente, CategoriaClienteItens)
    Else
        CategoriaClienteItens.Clear
        CategoriaClienteItens.Enabled = False
    End If


    Exit Sub

Erro_CategoriaCliente_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168906)

    End Select

    Exit Sub

End Sub

Private Sub Carrega_ComboCategoriaCliente(ByVal objCombo As ComboBox)

Dim lErro As Long
Dim colCategoriaCliente As New Collection
Dim objCategoriaCliente As New ClassCategoriaCliente

On Error GoTo Erro_Carrega_ComboCategoriaCliente

    'Le as categorias de cliente
    lErro = CF("CategoriaCliente_Le_Todos", colCategoriaCliente)
    If lErro <> SUCESSO Then gError 131995

    'Preenche CategoriaCliente
    For Each objCategoriaCliente In colCategoriaCliente

        objCombo.AddItem objCategoriaCliente.sCategoria

    Next
    
    Exit Sub

Erro_Carrega_ComboCategoriaCliente:

    Select Case gErr
    
        Case 131995

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168907)

    End Select

    Exit Sub

End Sub

Private Sub Carrega_ComboCategoriaItens(ByVal objComboCategoria As ComboBox, ByVal objComboItens As ListBox)

Dim lErro As Long
Dim iIndice As Integer
Dim objCategoriaClienteItem As ClassCategoriaClienteItem
Dim objCategoriaCliente As New ClassCategoriaCliente
Dim colCategoria As New Collection

On Error GoTo Erro_Carrega_ComboCategoriaItens

    'Verifica se a CategoriaCliente foi preenchida
    If objComboCategoria.ListIndex <> -1 Then

        objCategoriaCliente.sCategoria = objComboCategoria.Text

        'Lê os dados de Itens da Categoria do Cliente
        lErro = CF("CategoriaCliente_Le_Itens", objCategoriaCliente, colCategoria)
        If lErro <> SUCESSO Then gError 131994

        objComboItens.Enabled = True

        'Limpa os dados de ItemCategoriaCliente
        objComboItens.Clear

        'Preenche ItemCategoriaCliente
        For Each objCategoriaClienteItem In colCategoria

            objComboItens.AddItem objCategoriaClienteItem.sItem

        Next
        
        CategoriaClienteTodas.Value = vbFalse
    
    Else
        
        'Senão Desablita ItemCategoriaCliente
        objComboItens.Clear
        objComboItens.Enabled = False
    
    End If
    
    Exit Sub

Erro_Carrega_ComboCategoriaItens:

    Select Case gErr
    
        Case 131993

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168908)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaCliente_Validate(Cancel As Boolean)

Dim lErro As Long, iCodigo As Integer

On Error GoTo Erro_CategoriaCliente_Validate

    If Len(CategoriaCliente.Text) <> 0 And CategoriaCliente.ListIndex = -1 Then
    
        'pesquisa a categoria na lista
        lErro = Combo_Item_Igual(CategoriaCliente)
        If lErro <> SUCESSO And lErro <> 12253 Then gError 131998
        
        If lErro <> SUCESSO Then gError 131999
    
    End If
    
    'Se a CategoriaCliente estiver em branco desabilita e limpa a combo
    If Len(Trim(CategoriaCliente.Text)) = 0 Then
        CategoriaClienteItens.Clear
        CategoriaClienteItens.Enabled = False
    End If
    
    Exit Sub

Erro_CategoriaCliente_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 131998
         
        Case 131999
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_NAO_CADASTRADA", gErr, CategoriaCliente.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168909)

    End Select

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
