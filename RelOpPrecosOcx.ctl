VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpPrecosOcx 
   ClientHeight    =   4185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8250
   KeyPreview      =   -1  'True
   ScaleHeight     =   4185
   ScaleWidth      =   8250
   Begin VB.Frame Frame2 
      Caption         =   "Cálculo do ICMS ST"
      Height          =   975
      Left            =   135
      TabIndex        =   26
      Top             =   3120
      Width           =   7950
      Begin VB.CheckBox IEIsento 
         Caption         =   "Isento de Inscrição Estadual"
         Enabled         =   0   'False
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
         Left            =   1470
         TabIndex        =   34
         Top             =   645
         Width           =   2805
      End
      Begin VB.CheckBox IENaoContrib 
         Caption         =   "Não Contribuinte do ICMS"
         Enabled         =   0   'False
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
         Left            =   4875
         TabIndex        =   33
         Top             =   615
         Width           =   2760
      End
      Begin VB.ComboBox Filial 
         Height          =   315
         Left            =   4875
         TabIndex        =   29
         Top             =   255
         Width           =   2145
      End
      Begin MSMask.MaskEdBox Cliente 
         Height          =   300
         Left            =   750
         TabIndex        =   27
         Top             =   270
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label UF 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   750
         TabIndex        =   32
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "UF:"
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
         Left            =   390
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   31
         Top             =   645
         Width           =   315
      End
      Begin VB.Label Label2 
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
         Left            =   4335
         TabIndex        =   30
         Top             =   315
         Width           =   465
      End
      Begin VB.Label LabelCliente 
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
         Left            =   45
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   28
         Top             =   315
         Width           =   660
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6000
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   240
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpPrecosOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpPrecosOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpPrecosOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpPrecosOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produtos"
      Height          =   1035
      Left            =   135
      TabIndex        =   19
      Top             =   1050
      Width           =   7950
      Begin MSMask.MaskEdBox ProdutoFinal 
         Height          =   315
         Left            =   765
         TabIndex        =   3
         Top             =   600
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
         Left            =   765
         TabIndex        =   2
         Top             =   240
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
         Left            =   300
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   23
         Top             =   615
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
         TabIndex        =   22
         Top             =   270
         Width           =   615
      End
      Begin VB.Label DescProdFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2280
         TabIndex        =   21
         Top             =   600
         Width           =   5580
      End
      Begin VB.Label DescProdInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2280
         TabIndex        =   20
         Top             =   240
         Width           =   5580
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Categoria"
      Height          =   990
      Left            =   135
      TabIndex        =   14
      Top             =   2115
      Width           =   7950
      Begin VB.ComboBox ValorFinal 
         Height          =   315
         Left            =   4875
         TabIndex        =   7
         Top             =   570
         Width           =   3000
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
         Left            =   765
         TabIndex        =   4
         Top             =   195
         Width           =   855
      End
      Begin VB.ComboBox ValorInicial 
         Height          =   315
         Left            =   750
         TabIndex        =   6
         Top             =   570
         Width           =   2850
      End
      Begin VB.ComboBox Categoria 
         Height          =   315
         Left            =   2880
         TabIndex        =   5
         Top             =   180
         Width           =   4995
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
         Left            =   1965
         TabIndex        =   18
         Top             =   225
         Width           =   855
      End
      Begin VB.Label Label8 
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
         TabIndex        =   17
         Top             =   615
         Width           =   315
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
         Left            =   4455
         TabIndex        =   16
         Top             =   615
         Width           =   360
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   15
         Left            =   360
         TabIndex        =   15
         Top             =   720
         Width           =   30
      End
   End
   Begin VB.ComboBox TabelaPrecos 
      Height          =   315
      Left            =   1785
      TabIndex        =   1
      Top             =   660
      Width           =   2430
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
      Left            =   4350
      Picture         =   "RelOpPrecosOcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   270
      Width           =   1575
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpPrecosOcx.ctx":0A96
      Left            =   885
      List            =   "RelOpPrecosOcx.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   270
      Width           =   3330
   End
   Begin VB.Label LabelTab 
      AutoSize        =   -1  'True
      Caption         =   "Tabela de Preços:"
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
      Left            =   180
      TabIndex        =   25
      Top             =   705
      Width           =   1575
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
      Left            =   180
      TabIndex        =   24
      Top             =   315
      Width           =   615
   End
End
Attribute VB_Name = "RelOpPrecosOcx"
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
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1

Dim iAlterado As Integer

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio
Dim giProdInicial As Integer

Public Sub Form_Load()

Dim lErro As Long
Dim colCategoriaProduto As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto

On Error GoTo Erro_Form_Load
    
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    Set objEventoCliente = New AdmEvento
    
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoInicial)
    If lErro <> SUCESSO Then Error 37080

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoFinal)
    If lErro <> SUCESSO Then Error 37081

    'Le as categorias de produto
    lErro = CF("CategoriasProduto_Le_Todas", colCategoriaProduto)
    If lErro <> SUCESSO And lErro <> 22542 Then Error 37083

    'Preenche CategoriaProduto
    For Each objCategoriaProduto In colCategoriaProduto

        Categoria.AddItem objCategoriaProduto.sCategoria

    Next
    
    lErro = Carrega_TabelaPrecos()
    If lErro <> SUCESSO Then Error 37084
    
    Call Define_Padrao

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = Err

    Select Case Err

        Case 37080, 37081, 37083, 37084

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171374)

    End Select

    Exit Sub

End Sub

Private Sub Define_Padrao()
'Preenche a tela com as opções padrão de FilialEmpresa

Dim lErro As Long

On Error GoTo Erro_Define_Padrao

    giProdInicial = 1
    
    TodasCategorias_Click
    
    TodasCategorias = 1
                   
    Exit Sub

Erro_Define_Padrao:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171375)

    End Select

    Exit Sub

End Sub


Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim lCliente As Long, iFilial As Integer

On Error GoTo Erro_PreencherParametrosNaTela

 Call Limpar_Tela

    lErro = objRelOpcoes.Carregar
    If lErro Then Error 37085
   
    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro Then Error 37086

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then Error 37087

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro Then Error 37088

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then Error 37089
    
    'pega parâmetro TodasCategorias e exibe
    lErro = objRelOpcoes.ObterParametro("NTODASCAT", sParam)
    If lErro Then Error 37090

    TodasCategorias.Value = CInt(sParam)

    'pega parâmetro categoria de produto e exibe
    lErro = objRelOpcoes.ObterParametro("TCATPROD", sParam)
    If lErro Then Error 37091
    
    Categoria.Text = sParam

    'pega parâmetro valor inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TITEMCATPRODINI", sParam)
    If lErro Then Error 37092
    
    ValorInicial.Text = sParam
    
    'pega parâmetro Valor Final e exibe
    lErro = objRelOpcoes.ObterParametro("TITEMCATPRODFIM", sParam)
    If lErro Then Error 37093
    
    ValorFinal.Text = sParam
    
    'pega parâmetro Valor Final e exibe
    lErro = objRelOpcoes.ObterParametro("TTABPRECO", sParam)
    If lErro Then Error 37094
    
    TabelaPrecos.Text = sParam
    
    lErro = objRelOpcoes.ObterParametro("NCLIENTE", sParam)
    If lErro <> SUCESSO Then Error 37094
    
    lCliente = StrParaLong(sParam)
    
    lErro = objRelOpcoes.ObterParametro("NFILIALCLI", sParam)
    If lErro <> SUCESSO Then Error 37094
    
    iFilial = StrParaLong(sParam)
    
    If lCliente <> 0 Then

        Cliente.Text = CStr(lCliente)
        Call Cliente_Validate(bSGECancelDummy)
        
        Filial.Text = iFilial
        Call Filial_Validate(bSGECancelDummy)
    
    End If
   
    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 37085 To 37094

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171376)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoOp = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    Set objEventoCliente = Nothing
    
End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82600

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82601
    
    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 82602

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 82600, 82602

        Case 82601
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171377)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82603

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82604

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 82605

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 82603, 82605

        Case 82604
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171378)

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
        If lErro <> SUCESSO Then gError 82627

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 82627

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171379)

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
        If lErro <> SUCESSO Then gError 82626

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 82626

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171380)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 29891
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 37078
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 37078
        
        Case 29891
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171381)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub


Sub Limpar_Tela()

    Call Limpa_Tela(Me)
    
    TabelaPrecos.Text = ""

    DescProdInic.Caption = ""
    DescProdFim.Caption = ""
    
    UF.Caption = ""
   
    ComboOpcoes.SetFocus

End Sub

Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String, lCliente As Long, iFilial As Integer) As Long
'Formata os produtos retornando em sProd_I e sProd_F
'Verifica se os parâmetros iniciais são maiores que os finais

Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_Formata_E_Critica_Parametros

    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then Error 37096

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then Error 37097

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then Error 37098

    End If
    
    'valor inicial não pode ser maior que o valor final
    If Trim(ValorInicial.Text) <> "" And Trim(ValorFinal.Text) <> "" Then
    
         If ValorInicial.Text > ValorFinal.Text Then Error 37099
         
     Else
        
        If Trim(ValorInicial.Text) = "" And Trim(ValorFinal.Text) = "" And TodasCategorias.Value = 0 Then Error 38001
    
    End If
    
    'Verifica se a Conta está preenchida
    If Len(Trim(TabelaPrecos.Text)) = 0 Then Error 54503
    
    'Verifica se o Cliente foi preenchido
    If Len(Trim(Cliente.ClipText)) > 0 Then

        objCliente.sNomeReduzido = Cliente.Text

        'Lê o Cliente através do Nome Reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError ERRO_SEM_MENSAGEM

        If lErro = SUCESSO Then
            lCliente = objCliente.lCodigo
        End If
            
    End If

    iFilial = Codigo_Extrai(Filial.Text)
          
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 54503
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_NAO_PREENCHIDA", gErr)
            
        Case 37096
            ProdutoInicial.SetFocus

        Case 37097
            ProdutoFinal.SetFocus

        Case 37098
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            ProdutoInicial.SetFocus
             
        Case 37099
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_INICIAL_MAIOR", gErr)
            ValorInicial.SetFocus
            
        Case 38001
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_NAO_INFORMADO", gErr)
            ValorInicial.SetFocus
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171382)

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

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sProd_I As String
Dim sProd_F As String
Dim iTabPreco As Integer
Dim lCliente As Long, iFilial As Integer

On Error GoTo Erro_PreencherRelOp

    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)
       
    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F, lCliente, iFilial)
    If lErro <> SUCESSO Then Error 37106
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 37107
         
    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then Error 37108

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then Error 37109
         
    lErro = objRelOpcoes.IncluirParametro("NTODASCAT", CStr(TodasCategorias.Value))
    If lErro <> AD_BOOL_TRUE Then Error 37110
    
    lErro = objRelOpcoes.IncluirParametro("TCATPROD", Categoria.Text)
    If lErro <> AD_BOOL_TRUE Then Error 37111
    
    lErro = objRelOpcoes.IncluirParametro("TITEMCATPRODINI", ValorInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 37112
    
    lErro = objRelOpcoes.IncluirParametro("TITEMCATPRODFIM", ValorFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 37113
    
    iTabPreco = Codigo_Extrai(TabelaPrecos.Text)
    
    lErro = objRelOpcoes.IncluirParametro("NTABPRECO", CStr(iTabPreco))
    If lErro <> AD_BOOL_TRUE Then Error 37114
    
    lErro = objRelOpcoes.IncluirParametro("TTABPRECO", TabelaPrecos.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54846
    
    lErro = objRelOpcoes.IncluirParametro("NCLIENTE", CStr(lCliente))
    If lErro <> AD_BOOL_TRUE Then Error 37114
    
    lErro = objRelOpcoes.IncluirParametro("TCLIENTE", Cliente.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54846
    
    lErro = objRelOpcoes.IncluirParametro("NFILIALCLI", CStr(iFilial))
    If lErro <> AD_BOOL_TRUE Then Error 37114
    
    lErro = objRelOpcoes.IncluirParametro("TFILIALCLI", Filial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54846
    
    If TodasCategorias.Value = 0 Then
        gobjRelatorio.sNomeTsk = "precosca"
    Else
        gobjRelatorio.sNomeTsk = "precos"
    End If
        
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd_I, sProd_F)
    If lErro <> SUCESSO Then Error 37115

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 37106 To 37115

        Case 54846
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171383)

    End Select

    Exit Function

End Function


Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 37116

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 37117

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Limpar_Tela
        Call Define_Padrao

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 37116
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 37117

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171384)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 37118

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 37118

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171385)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 37119

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then Error 37120

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 37121

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 37119
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 37120, 37121

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171386)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate

    giProdInicial = 0

    lErro = CF("Produto_Perde_Foco", ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO And lErro <> 27095 Then Error 37122
    
    If lErro <> SUCESSO Then Error 43269

    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True


    Select Case Err

        Case 37122

         Case 43269
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171387)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    giProdInicial = 1

    lErro = CF("Produto_Perde_Foco", ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then Error 37123
    
    If lErro <> SUCESSO Then Error 43270

    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True


    Select Case Err

        Case 37123

         Case 43270
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171388)

    End Select

    Exit Sub

End Sub


Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd_I As String, sProd_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long


On Error GoTo Erro_Monta_Expressao_Selecao

   If sProd_I <> "" Then sExpressao = "Produto >= " & Forprint_ConvTexto(sProd_I)

   If sProd_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Produto <= " & Forprint_ConvTexto(sProd_F)

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
    
''    If Trim(TabelaPrecos.Text) <> "" Then
''
''        If sExpressao <> "" Then sExpressao = sExpressao & " E "
''        sExpressao = sExpressao & "TabelaPreco = " & Forprint_ConvTexto(TabelaPrecos.Text)
''
''    End If
     
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171389)

    End Select

    Exit Function

End Function

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
        If lErro <> SUCESSO And lErro <> 22540 Then Error 37124

        If lErro <> SUCESSO Then Error 37125 'Categoria não está cadastrada

        'Lê os dados de itens de categorias de produto
        lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colCategoria)
        If lErro <> SUCESSO Then Error 37126

        'Preenche Valor Inicial e final
        For Each objCategoriaProdutoItem In colCategoria

            ValorInicial.AddItem (objCategoriaProdutoItem.sItem)
            ValorFinal.AddItem (objCategoriaProdutoItem.sItem)

        Next

    End If

    Exit Sub

Erro_Categoria_Click:

    Select Case Err

        Case 37124
            Categoria.SetFocus
            
        Case 37125
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_INEXISTENTE", Err)
            Categoria.SetFocus
            
        Case 37126

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171390)

    End Select

    Exit Sub

End Sub


Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is ProdutoInicial Then
            Call LabelProdutoDe_Click
        ElseIf Me.ActiveControl Is ProdutoFinal Then
            Call LabelProdutoAte_Click
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171391)

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
            If lErro <> SUCESSO And lErro <> 22603 Then Error 37127

            'Item da Categoria não está cadastrado
            If lErro <> SUCESSO Then Error 37128
            
        End If

    End If

    Exit Sub

Erro_ValorInicial_Click:

    Select Case Err

        Case 37127
            ValorInicial.SetFocus

        Case 37128
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_INEXISTENTE", Err, objCategoriaProdutoItem.sItem, objCategoriaProdutoItem.sCategoria)
            ValorInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171392)

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
            If lErro <> SUCESSO And lErro <> 22603 Then Error 37129
                                    
            'Item da Categoria não está cadastrado
            If lErro <> SUCESSO Then Error 37130
        End If

    End If

    Exit Sub

Erro_ValorFinal_Click:

    Select Case Err

        Case 37129
            ValorFinal.SetFocus

        Case 37130
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_INEXISTENTE", Err, objCategoriaProdutoItem.sItem, objCategoriaProdutoItem.sCategoria)
            ValorFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171393)

    End Select

    Exit Sub

End Sub

Private Function Carrega_TabelaPrecos() As Long
'Carrega a Combo TabelaPrecos

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim iIndice As Integer
Dim colCodigoDescricao As New AdmColCodigoNome

On Error GoTo Erro_Carrega_TabelaPrecos

    'lê códigos e descrições da tabela TabelasDePrecos e devolve na coleção
    lErro = CF("Cod_Nomes_Le", "TabelasDePreco", "Codigo", "Descricao", STRING_TABELA_PRECO_DESCRICAO, colCodigoDescricao, "Ativo = 1")
    If lErro <> SUCESSO Then Error 37131
    
    'preenche a combo
    For Each objCodigoNome In colCodigoDescricao
        
        If objCodigoNome.iCodigo <> 0 Then
            TabelaPrecos.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
            TabelaPrecos.ItemData(TabelaPrecos.NewIndex) = objCodigoNome.iCodigo
             
        End If
    
    Next

    Carrega_TabelaPrecos = SUCESSO

    Exit Function

Erro_Carrega_TabelaPrecos:

    Carrega_TabelaPrecos = Err

    Select Case Err

        'Erro já tratado
        Case 37131

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171394)

    End Select

    Exit Function

End Function


Private Sub TabelaPrecos_Validate(Cancel As Boolean)
'Busca a descricao com código digitado na Combo

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_TabelaPrecos_Validate

    'se uma opcao da lista estiver selecionada, OK
    If TabelaPrecos.ListIndex <> -1 Then Exit Sub

    If Len(Trim(TabelaPrecos.Text)) = 0 Then Exit Sub

    lErro = Combo_Seleciona(TabelaPrecos, iCodigo)
    If lErro <> SUCESSO Then Error 37132

    Exit Sub

Erro_TabelaPrecos_Validate:

    Cancel = True


    Select Case Err

        Case 37132
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELA_PRECO_NAO_CADASTRADA", Err, iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171395)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_PRECOS
    Set Form_Load_Ocx = Me
    Caption = "Lista de Preços"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpPrecos"
    
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

Private Sub LabelTab_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTab, Source, X, Y)
End Sub

Private Sub LabelTab_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTab, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    'Preenche campo Cliente
    Cliente.Text = objCliente.sNomeReduzido

    'Executa o Validate
    Call Cliente_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Public Sub LabelCliente_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    objCliente.sNomeReduzido = Cliente.Text

    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Public Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate

    'Verifica se o Cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        'Busca o Cliente no BD
        lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
        If lErro <> SUCESSO Then gError 185150
                   
        lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
        If lErro <> SUCESSO Then gError 185151

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)
        
        If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", Filial, iCodFilial)

    'Se não estiver preenchido
    ElseIf Len(Trim(Cliente.Text)) = 0 Then

        'Limpa a Combo de Filiais
        Filial.Clear
        
        IEIsento.Value = vbUnchecked
        IENaoContrib.Value = vbUnchecked
        UF.Caption = ""

    End If

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 185150, 185151

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185152)

    End Select

    Exit Sub

End Sub

Private Sub Cliente_GotFocus()
    Call MaskEdBox_TrataGotFocus(Cliente, iAlterado)
End Sub

Private Sub Cliente_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult
Dim objCliente As New ClassCliente

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida ou alterada
    If Len(Trim(Filial.Text)) = 0 Then
        IEIsento.Value = vbUnchecked
        IENaoContrib.Value = vbUnchecked
        UF.Caption = ""
        Exit Sub
    End If

    'Verifica se é uma filial selecionada
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 185153

    'Se não encontrou o CÓDIGO
    If lErro = 6730 Then

        'Verifica se o cliente foi digitado
        If Len(Trim(Cliente.Text)) = 0 Then gError 185154

        sCliente = Cliente.Text
        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe Filial com o código extraído
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 185155

        If lErro = 17660 Then

            'Lê o Cliente
            objCliente.sNomeReduzido = sCliente
            lErro = CF("Cliente_Le_NomeReduzido", objCliente)
            If lErro <> SUCESSO And lErro <> 12348 Then gError 185156

            'Se encontrou o Cliente
            If lErro = SUCESSO Then
                
                objFilialCliente.lCodCliente = objCliente.lCodigo

                gError 185157
            
            End If
            
        End If
        
        If iCodigo <> 0 Then
        
            'Coloca na tela a Filial lida
            Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome
        
        Else
            
            objCliente.lCodigo = 0
            objFilialCliente.iCodFilial = 0
            
        End If
        
        lErro = Trata_FilialCliente
        If lErro <> SUCESSO Then gError 185158
        
    'Não encontrou a STRING
    ElseIf lErro = 6731 Then
        
        'trecho incluido por Leo em 17/04/02
        objCliente.sNomeReduzido = Cliente.Text
        
        'Lê o Cliente
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 185158
        
        If lErro = SUCESSO Then gError 185159
        
    End If

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 185153, 185155

        Case 185154
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 185156, 185158 'tratado na rotina chamada

        Case 185157
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If

        Case 185159
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185160)

    End Select

    Exit Sub


End Sub

Private Sub Filial_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Filial_Click()

Dim lErro As Long

On Error GoTo Erro_Filial_Click

    iAlterado = REGISTRO_ALTERADO

    'Se nenhuma filial foi selecionada, sai.
    If Filial.ListIndex = -1 Then Exit Sub

    'Faz o tratamento para a filial do cliente selecionada
    lErro = Trata_FilialCliente()
    If lErro <> SUCESSO Then gError 23581

    Exit Sub

Erro_Filial_Click:

    Select Case gErr

        Case 23581

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157954)

    End Select

    Exit Sub

End Sub

Function Trata_FilialCliente() As Long

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente
Dim objCliente As New ClassCliente
Dim objEndereco As New ClassEndereco

On Error GoTo Erro_Trata_FilialCliente

    objFilialCliente.iCodFilial = Codigo_Extrai(Filial.Text)
    objCliente.sNomeReduzido = Trim(Cliente.Text)

    lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Trim(Cliente.Text), objFilialCliente)
    If lErro <> SUCESSO And lErro <> 17660 Then gError ERRO_SEM_MENSAGEM
    If lErro = 17660 Then gError 25434

    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError ERRO_SEM_MENSAGEM
    If lErro = 12348 Then gError 26517
    
    objEndereco.lCodigo = objFilialCliente.lEndereco
    
    lErro = CF("Endereco_Le", objEndereco)
    If lErro <> SUCESSO And lErro <> 12309 Then gError ERRO_SEM_MENSAGEM
    
    UF.Caption = objEndereco.sSiglaEstado
    
    If objFilialCliente.iIEIsento = MARCADO Then
        IEIsento.Value = vbChecked
    Else
        IEIsento.Value = vbUnchecked
    End If
    
    If objFilialCliente.iIENaoContrib = MARCADO Then
        IENaoContrib.Value = vbChecked
    Else
        IENaoContrib.Value = vbUnchecked
    End If

    Trata_FilialCliente = SUCESSO

    Exit Function

Erro_Trata_FilialCliente:

    Trata_FilialCliente = gErr

    Select Case gErr

        Case 25434
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA1", gErr, Cliente.Text, objFilialCliente.iCodFilial)

        Case 26515, 26516, 51063, 51064, 51065, 25724, 124056, 131894, 140379, 207414

        Case 26517
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158054)

    End Select

    Exit Function

End Function
