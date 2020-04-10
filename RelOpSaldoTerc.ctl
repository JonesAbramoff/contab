VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpSaldoTercOcx 
   ClientHeight    =   5715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6540
   KeyPreview      =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   6540
   Begin VB.Frame Frame4 
      Caption         =   "Layout"
      Height          =   660
      Left            =   45
      TabIndex        =   43
      Top             =   4965
      Width           =   6360
      Begin VB.ComboBox Layout 
         Height          =   315
         ItemData        =   "RelOpSaldoTerc.ctx":0000
         Left            =   600
         List            =   "RelOpSaldoTerc.ctx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   5565
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Terceiros"
      Height          =   1365
      Left            =   45
      TabIndex        =   28
      Top             =   1500
      Width           =   6360
      Begin VB.Frame FrameTerceiro 
         Caption         =   "Terceiro"
         Enabled         =   0   'False
         Height          =   660
         Left            =   45
         TabIndex        =   33
         Top             =   570
         Width           =   6150
         Begin MSMask.MaskEdBox Cliente 
            Height          =   300
            Left            =   1080
            TabIndex        =   36
            Top             =   225
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   4410
            TabIndex        =   34
            Top             =   225
            Width           =   1680
         End
         Begin MSMask.MaskEdBox Fornecedor 
            Height          =   300
            Left            =   1080
            TabIndex        =   35
            Top             =   225
            Visible         =   0   'False
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label ClienteLabel 
            AutoSize        =   -1  'True
            Caption         =   "Cliente:"
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
            Left            =   405
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   39
            Top             =   285
            Width           =   660
         End
         Begin VB.Label LabelFilial 
            AutoSize        =   -1  'True
            Caption         =   "Filial:"
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
            Left            =   3900
            TabIndex        =   38
            Top             =   285
            Width           =   465
         End
         Begin VB.Label FornecedorLabel 
            AutoSize        =   -1  'True
            Caption         =   "Fornecedor:"
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
            Left            =   45
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   37
            Top             =   285
            Visible         =   0   'False
            Width           =   1035
         End
      End
      Begin VB.OptionButton OptTipoTerc 
         Caption         =   "Fornecedor"
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
         Left            =   4500
         TabIndex        =   32
         Top             =   330
         Width           =   1500
      End
      Begin VB.OptionButton OptTipoTerc 
         Caption         =   "Cliente"
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
         Left            =   3210
         TabIndex        =   31
         Top             =   315
         Width           =   1005
      End
      Begin VB.OptionButton OptTipoTerc 
         Caption         =   "Não Identificado"
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
         Index           =   3
         Left            =   1245
         TabIndex        =   30
         Top             =   315
         Width           =   2130
      End
      Begin VB.OptionButton OptTipoTerc 
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
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   300
         Value           =   -1  'True
         Width           =   1005
      End
   End
   Begin VB.Frame FrameTipoRelat 
      Caption         =   "Tipo de Relatório"
      Height          =   525
      Left            =   45
      TabIndex        =   27
      Top             =   930
      Width           =   4680
      Begin VB.OptionButton OptTipo 
         Caption         =   "De Terceiros conosco"
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
         Index           =   1
         Left            =   2280
         TabIndex        =   9
         Top             =   165
         Width           =   2265
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Nosso em Terceiros"
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
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   210
         Value           =   -1  'True
         Width           =   2715
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4260
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   75
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "RelOpSaldoTerc.ctx":0028
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   630
         Picture         =   "RelOpSaldoTerc.ctx":0182
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpSaldoTerc.ctx":030C
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpSaldoTerc.ctx":083E
         Style           =   1  'Graphical
         TabIndex        =   15
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
      Left            =   4815
      Picture         =   "RelOpSaldoTerc.ctx":09BC
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   840
      Width           =   1575
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpSaldoTerc.ctx":0ABE
      Left            =   840
      List            =   "RelOpSaldoTerc.ctx":0AC0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   150
      Width           =   2916
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produtos"
      Height          =   1065
      Left            =   45
      TabIndex        =   21
      Top             =   2880
      Width           =   6360
      Begin MSMask.MaskEdBox ProdutoFinal 
         Height          =   315
         Left            =   750
         TabIndex        =   2
         Top             =   645
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
         TabIndex        =   1
         Top             =   255
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label DescProdInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2250
         TabIndex        =   25
         Top             =   255
         Width           =   4050
      End
      Begin VB.Label DescProdFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2250
         TabIndex        =   24
         Top             =   645
         Width           =   4050
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
         Left            =   375
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   23
         Top             =   270
         Width           =   315
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
         Left            =   330
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   22
         Top             =   690
         Width           =   360
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Categoria de Produtos"
      Height          =   975
      Left            =   45
      TabIndex        =   16
      Top             =   3975
      Width           =   6360
      Begin VB.ComboBox Categoria 
         Height          =   315
         Left            =   2385
         TabIndex        =   4
         Top             =   180
         Width           =   3780
      End
      Begin VB.ComboBox ValorInicial 
         Height          =   315
         Left            =   600
         TabIndex        =   5
         Top             =   555
         Width           =   2535
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
         TabIndex        =   3
         Top             =   255
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.ComboBox ValorFinal 
         Height          =   315
         Left            =   3645
         TabIndex        =   6
         Top             =   555
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   15
         Left            =   360
         TabIndex        =   20
         Top             =   720
         Width           =   30
      End
      Begin VB.Label Label6 
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
         Left            =   3210
         TabIndex        =   19
         Top             =   600
         Width           =   555
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
         Left            =   225
         TabIndex        =   18
         Top             =   600
         Width           =   420
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
         Left            =   1425
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   315
      Left            =   1860
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   555
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   300
      Left            =   840
      TabIndex        =   41
      Top             =   570
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
      Left            =   300
      TabIndex        =   42
      Top             =   600
      Width           =   480
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
      Left            =   135
      TabIndex        =   26
      Top             =   210
      Width           =   615
   End
End
Attribute VB_Name = "RelOpSaldoTercOcx"
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
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio
Dim giProdInicial As Integer
Dim lCliFornAnt As Integer

Private Sub Form_Load()

Dim lErro As Long
Dim colCategoriaProduto As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objFiliais As AdmFiliais

On Error GoTo Erro_Form_Load

    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoFornecedor = New AdmEvento

    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoInicial)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoFinal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Le as categorias de produto
    lErro = CF("CategoriasProduto_Le_Todas", colCategoriaProduto)
    If lErro <> SUCESSO And lErro <> 22542 Then gError ERRO_SEM_MENSAGEM

    'Preenche CategoriaProduto
    For Each objCategoriaProduto In colCategoriaProduto
        Categoria.AddItem objCategoriaProduto.sCategoria
    Next
    
    Call Define_Padrao

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209660)

    End Select

    Exit Sub

End Sub

Sub Define_Padrao()

On Error GoTo Erro_Define_Padrao

    giProdInicial = 1
    
    Call DateParaMasked(Data, gdtDataAtual)
    
    TodasCategorias.Value = vbChecked
    Call TodasCategorias_Click
    
    OptTipo(0).Value = True
    OptTipoTerc(0).Value = True
    
    Call OptTipoTerc_Click(0)
    lCliFornAnt = 0
    
    Layout.ListIndex = 0
       
    Exit Sub

Erro_Define_Padrao:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209661)

    End Select

    Exit Sub

End Sub

Private Sub Data_GotFocus()
    Call MaskEdBox_TrataGotFocus(Data)
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
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209662)

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
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209663)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 209664

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 209664
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209665)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 209666

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 209666
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209667)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_GotFocus()
    giProdInicial = 1
End Sub

Private Sub ProdutoFinal_GotFocus()
    giProdInicial = 0
End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    Call Limpar_Tela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
       
    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'pega parâmetro TodasCategorias e exibe
    lErro = objRelOpcoes.ObterParametro("NTODASCAT", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    TodasCategorias.Value = CInt(sParam)

    lErro = objRelOpcoes.ObterParametro("NTIPO", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    OptTipo(StrParaInt(sParam)).Value = True
    
    lErro = objRelOpcoes.ObterParametro("NTIPOTERC", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    OptTipoTerc(StrParaInt(sParam)).Value = True
    
    lErro = objRelOpcoes.ObterParametro("NCODTERC", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If StrParaLong(sParam) <> 0 Then
        If OptTipoTerc(TIPO_TERC_CLIENTE).Value Then
            Cliente.Text = sParam
            Call Cliente_Validate(bSGECancelDummy)
        ElseIf OptTipoTerc(TIPO_TERC_FORNECEDOR).Value Then
            Fornecedor.Text = sParam
            Call Fornecedor_Validate(bSGECancelDummy)
        End If
        
        lErro = objRelOpcoes.ObterParametro("NFILIALTERC", sParam)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If StrParaInt(sParam) <> 0 Then
            Filial.Text = sParam
            Call Filial_Validate(bSGECancelDummy)
        End If
    End If
    
    'pega parâmetro categoria de produto e exibe
    lErro = objRelOpcoes.ObterParametro("TCATPROD", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Categoria.Text = sParam

    'pega parâmetro valor inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TITEMCATPRODINI", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    ValorInicial.Text = sParam
    
    'pega parâmetro Valor Final e exibe
    lErro = objRelOpcoes.ObterParametro("TITEMCATPRODFIM", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    ValorFinal.Text = sParam
 
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DDATA", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call DateParaMasked(Data, StrParaDate(sParam))
       
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NLAYOUT", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call Combo_Seleciona_ItemData(Layout, StrParaLong(sParam))
       
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209668)

    End Select

    Exit Function

End Function

Private Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    Set objEventoCliente = Nothing
    Set objEventoFornecedor = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 209669
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case 209669
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209670)

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
    
    Categoria.Text = ""
    ValorInicial.Text = ""
    ValorFinal.Text = ""
    
    Call Define_Padrao
    
    ComboOpcoes.SetFocus

End Sub

Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String, iTipo As Integer, iTipoTerc As Integer, lCodTerc As Long, iFilialTerc As Integer) As Long
'Formata os produtos retornando em sProd_I e sProd_F
'Verifica se os parâmetros iniciais são maiores que os finais

Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim lErro As Long, iIndice As Integer
Dim objCliente As ClassCliente
Dim objFornecedor As ClassFornecedor

On Error GoTo Erro_Formata_E_Critica_Parametros

    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then
        If sProd_I > sProd_F Then gError 209671
    End If

    If StrParaDate(Data.Text) = DATA_NULA Then gError 209672
    
    'valor inicial não pode ser maior que o valor final
    If Trim(ValorInicial.Text) <> "" And Trim(ValorFinal.Text) <> "" Then
        If ValorInicial.Text > ValorFinal.Text Then gError 209673
    Else
        If Trim(ValorInicial.Text) = "" And Trim(ValorFinal.Text) = "" And TodasCategorias.Value = 0 Then gError 209674
    End If
    
    For iIndice = 0 To 1
        If OptTipo(iIndice) Then
            iTipo = iIndice
            Exit For
        End If
    Next
    
    For iIndice = 0 To 3
        If OptTipoTerc(iIndice) Then
            iTipoTerc = iIndice
            Exit For
        End If
    Next
    
    If iTipoTerc = TIPO_TERC_CLIENTE Then
        'instancia o obj cliente
        Set objCliente = New ClassCliente
        
        If Len(Trim(Cliente.Text)) > 0 Then
        
            'preenche o objcliente c/ o nomered do cliente na tela
            objCliente.sNomeReduzido = Trim(Cliente.Text)
            
            'busca o cód. do cliente a apartir do nomereduzido
            lErro = CF("Cliente_Le_NomeReduzido", objCliente)
            If lErro <> SUCESSO And lErro <> 12348 Then gError ERRO_SEM_MENSAGEM
            
            'cliente não cadastrado
            If lErro = 12348 Then gError 209675
            
            'preenche o obj c/ o cód do cliente
            lCodTerc = objCliente.lCodigo
            
        End If
    ElseIf iTipoTerc = TIPO_TERC_FORNECEDOR Then
        'instancia o obj fornecedor
        Set objFornecedor = New ClassFornecedor
        
        If Len(Trim(Fornecedor.Text)) > 0 Then
        
            'carrega o objforncedor com o nomered do fornecedor da tela
            objFornecedor.sNomeReduzido = Trim(Fornecedor.Text)
            
            'busca o cód. do fornecedor a apartir do nomereduzido
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then gError ERRO_SEM_MENSAGEM
        
            'fornecedor não cadastrado
            If lErro = 6681 Then gError 209676
        
            'preenche o obj c/ o cód. do fornecedor
            lCodTerc = objFornecedor.lCodigo
            
        End If
    End If
    iFilialTerc = Codigo_Extrai(Filial.Text)
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
              
        Case ERRO_SEM_MENSAGEM

        Case 209671
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            ProdutoInicial.SetFocus
             
        Case 209672
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
            Data.SetFocus
            
        Case 209673
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_INICIAL_MAIOR", gErr)
            ValorInicial.SetFocus
            
        Case 209674
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_NAO_INFORMADO", gErr)
            ValorInicial.SetFocus
      
        Case 209675
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, Fornecedor.Text)
                    
        Case 209676
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, Cliente.Text)
      
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209677)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()
    ComboOpcoes.Text = ""
    Call Limpar_Tela
End Sub

Private Sub ComboOpcoes_Click()
    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)
    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)
End Sub

Function Traz_Produto_Tela(sProduto As String) As Long
'verifica e preenche o produto inicial e final com sua descriçao de acordo com o último foco
'sProduto deve estar no formato do BD

Dim lErro As Long

On Error GoTo Erro_Traz_Produto_Tela

    If giProdInicial Then

        lErro = CF("Traz_Produto_MaskEd", sProduto, ProdutoInicial, DescProdInic)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Else

        lErro = CF("Traz_Produto_MaskEd", sProduto, ProdutoFinal, DescProdFim)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Traz_Produto_Tela = SUCESSO

    Exit Function

Erro_Traz_Produto_Tela:

    Traz_Produto_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209678)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional ByVal bExecutando As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sProd_I As String
Dim sProd_F As String
Dim iTipo As Integer, iTipoTerc As Integer, lCodTerc As Long, iFilialTerc As Integer
Dim objRelEstTercSel As New ClassRelEstTercSel

On Error GoTo Erro_PreencherRelOp

    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)
    
    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F, iTipo, iTipoTerc, lCodTerc, iFilialTerc)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
         
    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
        
    lErro = objRelOpcoes.IncluirParametro("DDATA", Data.Text)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
       
    lErro = objRelOpcoes.IncluirParametro("NTODASCAT", CStr(TodasCategorias.Value))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("TCATPROD", Categoria.Text)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("TITEMCATPRODINI", ValorInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("TITEMCATPRODFIM", ValorFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("NTIPO", CStr(iTipo))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("NTIPOTERC", CStr(iTipoTerc))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("NCODTERC", CStr(lCodTerc))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("NFILIALTERC", CStr(iFilialTerc))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("NLAYOUT", CStr(Layout.ItemData(Layout.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    If bExecutando Then
    
        objRelEstTercSel.sProdutoDe = sProd_I
        objRelEstTercSel.sProdutoAte = sProd_F
        objRelEstTercSel.dtDataDe = StrParaDate(Data.Text)
        objRelEstTercSel.dtDataAte = DATA_NULA
        objRelEstTercSel.sCategoria = Categoria.Text
        objRelEstTercSel.sCategoriaItemDe = ValorInicial.Text
        objRelEstTercSel.sCategoriaItemAte = ValorFinal.Text
        objRelEstTercSel.iTipo = iTipo
        objRelEstTercSel.iTipoTerc = iTipoTerc
        objRelEstTercSel.iFilialEmpresa = giFilialEmpresa
        objRelEstTercSel.lCodTerc = lCodTerc
        objRelEstTercSel.iFilialTerc = iFilialTerc
    
        lErro = CF("RelSldEstTerc_Prepara", objRelEstTercSel)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(objRelEstTercSel.lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    End If

    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209679)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 209680

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call Limpar_Tela
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 209680
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209681)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If Codigo_Extrai(Layout.Text) = 2 Then
        gobjRelatorio.sNomeTsk = "SldTercC"
    Else
        gobjRelatorio.sNomeTsk = "SldTerc"
    End If

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209682)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 209683

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 209683
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209684)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate

    giProdInicial = 0

    lErro = CF("Produto_Perde_Foco", ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO And lErro <> 27095 Then gError ERRO_SEM_MENSAGEM
    
    If lErro <> SUCESSO Then gError 209685

    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True


    Select Case Err

        Case ERRO_SEM_MENSAGEM

         Case 209685
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209686)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    giProdInicial = 1

    lErro = CF("Produto_Perde_Foco", ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then gError ERRO_SEM_MENSAGEM
    
    If lErro <> SUCESSO Then gError 209687

    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

         Case 209687
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209688)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""
    
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209689)

    End Select

    Exit Function

End Function

Private Sub Data_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_Data_Validate

    If Len(Data.ClipText) > 0 Then

        sDataInic = Data.Text
        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209691)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Data.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209692)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Data.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209693)

    End Select

    Exit Sub

End Sub

Private Sub Categoria_GotFocus()
    If TodasCategorias.Value = 1 Then TodasCategorias.Value = 0
End Sub

Private Sub Categoria_Validate(Cancel As Boolean)
    Call Categoria_Click
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
         If lErro <> SUCESSO And lErro <> 22540 Then gError ERRO_SEM_MENSAGEM

         If lErro <> SUCESSO Then gError 209696 'Categoria não está cadastrada

        'Lê os dados de itens de categorias de produto
        lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colCategoria)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

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

    Select Case gErr
            
        Case 209696
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_INEXISTENTE", gErr)
            Categoria.SetFocus
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209697)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
    
        If Me.ActiveControl Is ProdutoInicial Then
            Call LabelProdutoDe_Click
        ElseIf Me.ActiveControl Is ProdutoFinal Then
            Call LabelProdutoAte_Click
        ElseIf Me.ActiveControl Is Cliente Then
            Call ClienteLabel_Click
        ElseIf Me.ActiveControl Is Fornecedor Then
            Call FornecedorLabel_Click
        End If
        
    End If

End Sub

Private Sub ValorInicial_GotFocus()
    If TodasCategorias.Value = 1 Then TodasCategorias.Value = 0
End Sub

Private Sub ValorFinal_GotFocus()
    If TodasCategorias.Value = 1 Then TodasCategorias.Value = 0
End Sub

Private Sub ValorInicial_Validate(Cancel As Boolean)
    Call ValorInicial_Click
End Sub

Private Sub ValorFinal_Validate(Cancel As Boolean)
    Call ValorFinal_Click
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

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209698)

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
            If lErro <> SUCESSO And lErro <> 22603 Then gError ERRO_SEM_MENSAGEM

            If lErro <> SUCESSO Then gError 209699 'Item da Categoria não está cadastrado

        End If

    End If

    Exit Sub

Erro_ValorInicial_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            ValorInicial.SetFocus

        Case 209699
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_INEXISTENTE", gErr, objCategoriaProdutoItem.sItem, objCategoriaProdutoItem.sCategoria)
            ValorInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209700)

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
            If lErro <> SUCESSO And lErro <> 22603 Then gError ERRO_SEM_MENSAGEM

            If lErro <> SUCESSO Then gError 209701 'Item da Categoria não está cadastrado

        End If

    End If

    Exit Sub

Erro_ValorFinal_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            ValorFinal.SetFocus

        Case 209701
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_INEXISTENTE", gErr, objCategoriaProdutoItem.sItem, objCategoriaProdutoItem.sCategoria)
            ValorFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209702)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_KARDEX
    Set Form_Load_Ocx = Me
    Caption = "Saldo em Estoque - Por Terceiro"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpSaldoTerc"
    
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

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub dIni_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dIni, Source, X, Y)
End Sub

Private Sub dIni_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dIni, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub FornecedorLabel_Click()
'chama o browser referente ao fornecedor

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'se o fornecedor estiver preenchido
    If Len(Trim(Fornecedor.Text)) > 0 Then
        'Preenche nomeReduzido com o fornecedor da tela
        objFornecedor.sNomeReduzido = Trim(Fornecedor.Text)
    End If
    
    'chama a tela c/ a lista de fornecedores
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

End Sub

Private Sub ClienteLabel_Click()
'chama o browser de clientes

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection

    'se o cliente foi preenchido
    If Len(Trim(Cliente.Text)) > 0 Then
        'Prenche o nome reduzido do Cliente
        objCliente.sNomeReduzido = Trim(Cliente.Text)
    End If
    
    'chama a tela de browser clienteslista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)
'verifica se o fornecedor selecionado é valido

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Validate
    
    'Verifica preenchimento de Fornecedor, se não foi preenchido
    If Len(Trim(Fornecedor.Text)) <> 0 Then

        'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
        lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        'Lê coleção de códigos, nomes de Filiais do Fornecedor
        lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    Else
        'limpa a filial e sai da rotina
        Filial.Clear
    
    End If
    
    If lCliFornAnt = REGISTRO_ALTERADO Then

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)
    
        'verifica se foi digitado nome ou cód. do fornecedor
        If colCodigoNome.Count = 1 Or iCodFilial <> 0 Then
            
            If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ
                
            'Seleciona filial na Combo Filial
            Call CF("Filial_Seleciona", Filial, iCodFilial)
            
        End If
        
        lCliFornAnt = 0
        
    End If

    Exit Sub

Erro_Fornecedor_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209703)

    End Select

    Exit Sub

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)
'verifica se o cliente é valido

Dim lErro As Long
Dim iCodFilial As Integer
Dim objCliente As New ClassCliente
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate

    'se o cliente não foi preenchido, sai da rotina
    If Len(Trim(Cliente.Text)) <> 0 Then
    
        'Busca o Cliente no BD
        lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        'busca no bd a relação de filiais referentes ao cliente
        lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    Else
        'limpa a filial e sai da rotina
        Filial.Clear
    End If
    
    If lCliFornAnt = REGISTRO_ALTERADO Then
    
        'Preenche ComboBox de Filiais do cliente
        Call CF("Filial_Preenche", Filial, colCodigoNome)
        
        'verifica se foi digitado nome ou cód. do cliente
        If colCodigoNome.Count = 1 Or iCodFilial <> 0 Then
            
            If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ
                
            'Seleciona filial na Combo Filial
            Call CF("Filial_Seleciona", Filial, iCodFilial)
            
        End If
        
        lCliFornAnt = 0
        
    End If
    
    Exit Sub
        
Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209704)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Filial_Validate(Cancel As Boolean)
'verifica se a filial do cliente\fornecedor é válida

Dim lErro As Long
Dim objFilialCliente As ClassFilialCliente
Dim objFilialFornecedor As ClassFilialFornecedor
Dim iCodigo As Integer

On Error GoTo Erro_Filial_Validate

    'Verifica se foi preenchida a ComboBox Filial
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

   'Verifica se está preenchida com o ítem selecionado na ComboBox Filial
    If Filial.ListIndex >= 0 Then Exit Sub

    'se o tipo de terc. for cliente
    If OptTipoTerc(TIPO_TERC_CLIENTE).Value = True Then
    
        'verifica se o cliente foi preenchido
        If Len(Trim(Cliente.Text)) = 0 Then gError 209705

        'Verifica se existe o ítem na List da Combo. Se existir seleciona.
        lErro = Combo_Seleciona(Filial, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError ERRO_SEM_MENSAGEM
    
        'Nao existe o ítem com o CÓDIGO na List da ComboBox
        If lErro = 6730 Then
    
            'instancia o obj
            Set objFilialCliente = New ClassFilialCliente
    
            'passa o nº preenchido como código
            objFilialCliente.iCodFilial = iCodigo
    
            'Tentativa de leitura da Filial com esse código no BD
            lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Trim(Cliente.Text), objFilialCliente)
            If lErro <> SUCESSO And lErro <> 17660 Then gError ERRO_SEM_MENSAGEM
    
            'Não encontrou Filial no  BD
            If lErro = 17660 Then gError 209706
    
            'Encontrou Filial no BD, coloca no Text da Combo
            Filial.Text = objFilialCliente.iCodFilial & SEPARADOR & objFilialCliente.sNome
    
        End If
            
        'Não existe o ítem com a STRING na List da ComboBox
        If lErro = 6731 Then gError 209707
    
    'senão, é o fornecedor
    ElseIf OptTipoTerc(TIPO_TERC_FORNECEDOR).Value = True Then

        'verifica se o fornecedor foi preenchido
        If Len(Trim(Fornecedor.Text)) = 0 Then gError 209708

        'Verifica se existe o ítem na List da Combo. Se existir seleciona.
        lErro = Combo_Seleciona(Filial, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError ERRO_SEM_MENSAGEM
    
        'Nao existe o ítem com o CÓDIGO na List da ComboBox
        If lErro = 6730 Then
    
            'instancia o obj
            Set objFilialFornecedor = New ClassFilialFornecedor
    
            'passa o nº preenchido como código
            objFilialFornecedor.iCodFilial = iCodigo
    
            'Tentativa de leitura da Filial com esse código no BD
            lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Trim(Fornecedor.Text), objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 18272 Then gError ERRO_SEM_MENSAGEM
    
            'Não encontrou Filial no  BD
            If lErro = 18272 Then gError 209709
    
            'Encontrou Filial no BD, coloca no Text da Combo
            Filial.Text = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome
    
        End If
            
        'Não existe o ítem com a STRING na List da ComboBox
        If lErro = 6731 Then gError 209710

    End If

    Exit Sub

Erro_Filial_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 209705
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 209706, 209707
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)

        Case 209708
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 209709, 209710
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", gErr, Fornecedor.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209711)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)
'traz os dados do item selecionado no browser

Dim objCliente As ClassCliente

    Set objCliente = obj1

    'Preenche o Cliente com o cod. do Cliente selecionado
    Cliente.Text = objCliente.lCodigo
    
    'Dispara o Validate de Cliente p/ a validação do cliente e preencher a filial
    Call Cliente_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)
'traz os dados do item selecionado no browser

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1

    'Coloca o cód. do fornecedor na Tela
    Fornecedor.Text = objFornecedor.lCodigo

    'dispara o validate do fornecedor p/ validar o fornecedor e preencher a filial
    Call Fornecedor_Validate(bSGECancelDummy)

    Me.Show

End Sub

Private Sub OptTipoTerc_Click(Index As Integer)
    Call OptionTipoTerc_Trata
End Sub

Private Sub OptionTipoTerc_Trata()
Dim bExibeCli As Boolean
Dim bExibeForn As Boolean

    bExibeCli = False
    bExibeForn = False
    FrameTerceiro.Enabled = False
    
   If OptTipoTerc(TIPO_TERC_CLIENTE).Value = True Then bExibeCli = True
   If OptTipoTerc(TIPO_TERC_FORNECEDOR).Value = True Then bExibeForn = True

    If bExibeCli Then
        'habilita o cliente
        Cliente.Visible = True
        ClienteLabel.Visible = True
        FrameTerceiro.Enabled = True
    Else
        Cliente.Visible = False
        Cliente.Text = ""
        ClienteLabel.Visible = False
    End If
    
    If bExibeForn Then
        'habilita o fornecedor
        Fornecedor.Visible = True
        FornecedorLabel.Visible = True
        FrameTerceiro.Enabled = True
    Else
        Fornecedor.Visible = False
        Fornecedor.Text = ""
        FornecedorLabel.Visible = False
    End If
    
    If Fornecedor.Text = "" And Cliente.Text = "" Then Filial.Clear

End Sub

Private Sub Cliente_Change()
    lCliFornAnt = REGISTRO_ALTERADO
End Sub

Private Sub Fornecedor_Change()
    lCliFornAnt = REGISTRO_ALTERADO
End Sub
