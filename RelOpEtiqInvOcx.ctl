VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl RelOpEtiqInvOcx 
   ClientHeight    =   5205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8790
   KeyPreview      =   -1  'True
   ScaleHeight     =   5205
   ScaleWidth      =   8790
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6480
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpEtiqInvOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpEtiqInvOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpEtiqInvOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpEtiqInvOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Categoria"
      Height          =   1470
      Left            =   135
      TabIndex        =   25
      Top             =   3075
      Width           =   5790
      Begin VB.ComboBox Categoria 
         Height          =   315
         Left            =   1635
         TabIndex        =   6
         Top             =   540
         Width           =   2745
      End
      Begin VB.ComboBox ValorInicial 
         Height          =   315
         Left            =   705
         TabIndex        =   7
         Top             =   1020
         Width           =   1950
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
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox ValorFinal 
         Height          =   315
         Left            =   3375
         TabIndex        =   8
         Top             =   1005
         Width           =   2100
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   15
         Left            =   360
         TabIndex        =   29
         Top             =   720
         Width           =   30
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
         Left            =   2955
         TabIndex        =   28
         Top             =   1065
         Width           =   360
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
         Left            =   330
         TabIndex        =   27
         Top             =   1065
         Width           =   315
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
         Left            =   720
         TabIndex        =   26
         Top             =   585
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produtos"
      Height          =   1350
      Left            =   135
      TabIndex        =   20
      Top             =   1665
      Width           =   5790
      Begin MSMask.MaskEdBox ProdutoFinal 
         Height          =   315
         Left            =   765
         TabIndex        =   4
         Top             =   855
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
         Left            =   780
         TabIndex        =   3
         Top             =   330
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
         Left            =   2370
         TabIndex        =   24
         Top             =   330
         Width           =   3135
      End
      Begin VB.Label DescProdFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2370
         TabIndex        =   23
         Top             =   855
         Width           =   3135
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
         Left            =   420
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   22
         Top             =   360
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
         Left            =   375
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   21
         Top             =   885
         Width           =   360
      End
   End
   Begin VB.TextBox ProximoNumero 
      Height          =   285
      Left            =   1725
      TabIndex        =   9
      Top             =   4695
      Width           =   825
   End
   Begin VB.Frame Frame2 
      Caption         =   "Almoxarifados"
      Height          =   840
      Left            =   120
      TabIndex        =   17
      Top             =   780
      Width           =   5790
      Begin MSMask.MaskEdBox AlmoxarifadoInicial 
         Height          =   315
         Left            =   750
         TabIndex        =   1
         Top             =   345
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox AlmoxarifadoFinal 
         Height          =   315
         Left            =   3450
         TabIndex        =   2
         Top             =   345
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label labelAlmoxarifadoFinal 
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
         Left            =   3000
         TabIndex        =   19
         Top             =   390
         Width           =   360
      End
      Begin VB.Label Label9 
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
         Left            =   360
         TabIndex        =   18
         Top             =   390
         Width           =   315
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpEtiqInvOcx.ctx":0994
      Left            =   1395
      List            =   "RelOpEtiqInvOcx.ctx":0996
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
      Left            =   4590
      Picture         =   "RelOpEtiqInvOcx.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      Width           =   1575
   End
   Begin VB.ListBox Almoxarifados 
      Height          =   3960
      ItemData        =   "RelOpEtiqInvOcx.ctx":0A9A
      Left            =   6060
      List            =   "RelOpEtiqInvOcx.ctx":0A9C
      TabIndex        =   10
      Top             =   1065
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Próximo número:"
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
      Left            =   240
      TabIndex        =   32
      Top             =   4725
      Width           =   1410
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
      Left            =   690
      TabIndex        =   31
      Top             =   315
      Width           =   615
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
      Height          =   195
      Left            =   6060
      TabIndex        =   30
      Top             =   855
      Width           =   1185
   End
End
Attribute VB_Name = "RelOpEtiqInvOcx"
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

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio
Dim giProdInicial As Integer
Dim giAlmoxInicial As Integer

Private Sub AlmoxarifadoInicial_GotFocus()
'Mostra a lista de almoxarifado

Dim lErro As Long

On Error GoTo Erro_AlmoxarifadoInicial_GotFocus

    giAlmoxInicial = 1

    Exit Sub

Erro_AlmoxarifadoInicial_GotFocus:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168707)

    End Select

    Exit Sub

End Sub


Private Sub AlmoxarifadoFinal_GotFocus()
'mostra a lista de almoxarifado

Dim lErro As Long

On Error GoTo Erro_AlmoxarifadoFinal_GotFocus

    giAlmoxInicial = 0

    Exit Sub

Erro_AlmoxarifadoFinal_GotFocus:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168708)

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
        If lErro <> SUCESSO Then gError 82337

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 82337

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168709)

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
        If lErro <> SUCESSO Then gError 82336

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 82336

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168710)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82339

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82340

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 82341

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 82339, 82341

        Case 82340
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168711)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82343

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82344

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 82345

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 82343, 82345

        Case 82344
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168712)

    End Select

    Exit Sub

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168713)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168714)

    End Select

    Exit Sub

End Sub

Private Sub Form_Load()

Dim lErro As Long
Dim colCategoriaProduto As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto

On Error GoTo Erro_Form_Load

    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento

    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoInicial)
    If lErro <> SUCESSO Then Error 37353

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoFinal)
    If lErro <> SUCESSO Then Error 37354

    'carrega a ListBox Almoxarifados
    lErro = Carrega_Lista_Almoxarifado()
    If lErro <> SUCESSO Then Error 37356
   
    'Le as categorias de produto
    lErro = CF("CategoriasProduto_Le_Todas", colCategoriaProduto)
    If lErro <> SUCESSO And lErro <> 22542 Then Error 37357

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

        Case 37353, 37354, 37356, 37357

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168715)

    End Select

    Exit Sub

End Sub

Private Sub Define_Padrao()
'Preenche a tela com as opções padrão

Dim lErro As Long

On Error GoTo Erro_Define_Padrao

    giProdInicial = 1
    
    giAlmoxInicial = 1
    
    TodasCategorias_Click
    
    TodasCategorias = 1
    
    Exit Sub

Erro_Define_Padrao:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168716)

    End Select

    Exit Sub

End Sub


Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

 Call Limpar_Tela

    lErro = objRelOpcoes.Carregar
    If lErro Then Error 37358

   'pega parâmetro Almoxarifado Inicial e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("NALMOXINIC", sParam)
    If lErro Then Error 37359
    
    AlmoxarifadoInicial.Text = sParam
    Call AlmoxarifadoInicial_Validate(bSGECancelDummy)
    
    'pega parâmetro Almoxarifado Final e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("NALMOXFIM", sParam)
    If lErro Then Error 37360
    
    AlmoxarifadoFinal.Text = sParam
    Call AlmoxarifadoFinal_Validate(bSGECancelDummy)
   
    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro Then Error 37361

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then Error 37362

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro Then Error 37363

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then Error 37364
    
    'pega parâmetro TodasCategorias e exibe
    lErro = objRelOpcoes.ObterParametro("NTODASCAT", sParam)
    If lErro Then Error 37365

    TodasCategorias.Value = CInt(sParam)

    'pega parâmetro categoria de produto e exibe
    lErro = objRelOpcoes.ObterParametro("TCATPROD", sParam)
    If lErro Then Error 37366
    
    Categoria.Text = sParam

    'pega parâmetro valor inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TITEMCATPRODINI", sParam)
    If lErro Then Error 37367
    
    ValorInicial.Text = sParam
    
    'pega parâmetro Valor Final e exibe
    lErro = objRelOpcoes.ObterParametro("TITEMCATPRODFIM", sParam)
    If lErro Then Error 37368
    
    ValorFinal.Text = sParam
    
    'pega parâmetro Próximo Número e exibe
    lErro = objRelOpcoes.ObterParametro("NPROXNUMERO", sParam)
    If lErro Then Error 37369
    
    ProximoNumero.Text = sParam
       
       
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 37358 To 37369

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168717)

    End Select

    Exit Function

End Function

Private Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 29896
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 37351

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 37351
        
        Case 29896
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168718)

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
    
    ComboOpcoes.SetFocus

End Sub


Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String, sAlmox_I, sAlmox_F) As Long
'Formata os produtos retornando em sProd_I e sProd_F
'Verifica se os parâmetros iniciais são maiores que os finais

Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then Error 37371

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then Error 37372

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then Error 37373

    End If

    'critica Almoxarifado Inicial e Final
    If AlmoxarifadoInicial.Text <> "" Then
        sAlmox_I = CStr(Codigo_Extrai(AlmoxarifadoInicial.Text))
        
    Else
        sAlmox_I = ""
        
    End If
    
        
    If AlmoxarifadoFinal.Text <> "" Then
        sAlmox_F = CStr(Codigo_Extrai(AlmoxarifadoFinal.Text))
    
    Else
        sAlmox_F = ""
        
    End If
    
    
    If sAlmox_I <> "" And sAlmox_F <> "" Then
          
        If sAlmox_I <> "" And sAlmox_F <> "" Then
        
            If CInt(sAlmox_I) > CInt(sAlmox_F) Then Error 37374
        
        End If
        
    End If
    
    'valor inicial não pode ser maior que o valor final
    If Trim(ValorInicial.Text) <> "" And Trim(ValorFinal.Text) <> "" Then
    
         If ValorInicial.Text > ValorFinal.Text Then Error 37375
         
    Else
        
        If Trim(ValorInicial.Text) = "" And Trim(ValorFinal.Text) = "" And TodasCategorias.Value = 0 Then Error 37993
    
    End If
    
    
          
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = Err

    Select Case Err

        Case 37371
            ProdutoInicial.SetFocus

        Case 37372
            ProdutoFinal.SetFocus

        Case 37373
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", Err)
            ProdutoInicial.SetFocus
         
         Case 37374
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INICIAL_MAIOR", Err)
            AlmoxarifadoInicial.SetFocus
                   
        Case 37375
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_INICIAL_MAIOR", Err)
            ValorInicial.SetFocus
            
        Case 37993
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_NAO_INFORMADO", Err)
            ValorInicial.SetFocus
      
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168719)

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
Dim sAlmox_I As String
Dim sAlmox_F As String

On Error GoTo Erro_PreencherRelOp

    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)
       
    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F, sAlmox_I, sAlmox_F)
    If lErro <> SUCESSO Then Error 37381

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 37382
         
    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then Error 37383

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then Error 37384
        
    lErro = objRelOpcoes.IncluirParametro("NALMOXINIC", sAlmox_I)
    If lErro <> AD_BOOL_TRUE Then Error 37385
        
    lErro = objRelOpcoes.IncluirParametro("TALMOXINICIAL", AlmoxarifadoInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54536
    
    lErro = objRelOpcoes.IncluirParametro("NALMOXFIM", sAlmox_F)
    If lErro <> AD_BOOL_TRUE Then Error 37386
            
    lErro = objRelOpcoes.IncluirParametro("TALMOXFINAL", AlmoxarifadoFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54537

    lErro = objRelOpcoes.IncluirParametro("NTODASCAT", CStr(TodasCategorias.Value))
    If lErro <> AD_BOOL_TRUE Then Error 37387
    
    lErro = objRelOpcoes.IncluirParametro("TCATPROD", Categoria.Text)
    If lErro <> AD_BOOL_TRUE Then Error 37388
    
    lErro = objRelOpcoes.IncluirParametro("TITEMCATPRODINI", ValorInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 37389
    
    lErro = objRelOpcoes.IncluirParametro("TITEMCATPRODFIM", ValorFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 37390
   
    lErro = objRelOpcoes.IncluirParametro("NPROXNUMERO", ProximoNumero.Text)
    If lErro <> AD_BOOL_TRUE Then Error 37391
       
    If TodasCategorias.Value = 0 Then
        
        gobjRelatorio.sNomeTsk = "EtInvCat"
    Else
        gobjRelatorio.sNomeTsk = "EtiqInv"
    End If
   
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sAlmox_I, sAlmox_F, sProd_I, sProd_F)
    If lErro <> SUCESSO Then Error 37392

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 37381 To 37392
        
        Case 54536, 54537

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168720)

    End Select

    Exit Function

End Function


Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 37393

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 37394

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Call BotaoLimpar_Click

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 37393
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 37394

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168721)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 37395

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 37395

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168722)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 37396

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then Error 37397

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 37398

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome
    
    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 37396
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 37397, 37398

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168723)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate

    giProdInicial = 0

    lErro = CF("Produto_Perde_Foco", ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO And lErro <> 27095 Then Error 37399
    
    If lErro <> SUCESSO Then Error 43255

    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True


    Select Case Err

        Case 37399

         Case 43255
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168724)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    giProdInicial = 1

    lErro = CF("Produto_Perde_Foco", ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then Error 37400
    
    If lErro <> SUCESSO Then Error 43256

    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True


    Select Case Err

        Case 37400

         Case 43256
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168725)

    End Select

    Exit Sub

End Sub


Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sAlmox_I As String, sAlmox_F As String, sProd_I As String, sProd_F As String) As Long
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

    If sAlmox_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Almoxarifado >= " & Forprint_ConvInt(CInt(sAlmox_I))

    End If

    If sAlmox_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Almoxarifado <= " & Forprint_ConvInt(CInt(sAlmox_F))

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
     
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168726)

    End Select

    Exit Function

End Function


Private Function Carrega_Lista_Almoxarifado() As Long
'Carrega a ListBox Almoxarifados

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As New AdmCodigoNome

On Error GoTo Erro_Carrega_Lista_Almoxarifado

    'Lê Códigos e NomesReduzidos da tabela Almoxarifado e devolve na coleção
    lErro = CF("Cod_Nomes_Le", "Almoxarifado", "Codigo", "NomeReduzido", STRING_ALMOXARIFADO_NOME_REDUZIDO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 37401

    'Preenche a ListBox Almoxarifados com os objetos da coleção
    For Each objCodigoDescricao In colCodigoDescricao

        Almoxarifados.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        Almoxarifados.ItemData(Almoxarifados.NewIndex) = objCodigoDescricao.iCodigo

    Next

    Carrega_Lista_Almoxarifado = SUCESSO

    Exit Function

Erro_Carrega_Lista_Almoxarifado:

    Carrega_Lista_Almoxarifado = Err

    Select Case Err

        Case 37401

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168727)

    End Select

    Exit Function

End Function

Private Sub Almoxarifados_DblClick()
'Preenche Almoxarifado Final ou Inicial com o almoxarifado selecionado

Dim lErro As Long
Dim sListBoxItem As String
Dim objCodigoDescricao As New AdmCodigoNome
Dim objAlmoxarifado As ClassAlmoxarifado
Dim objAlmoxSelecionado As ClassAlmoxarifado

On Error GoTo Erro_Almoxarifados_DblClick

    'Guarda a string selecionada na ListBox Almoxarifados
    sListBoxItem = Almoxarifados.List(Almoxarifados.ListIndex)
 
    If giAlmoxInicial = 1 Then
    
        AlmoxarifadoInicial.Text = sListBoxItem
        
    Else
        AlmoxarifadoFinal.Text = sListBoxItem

    End If

    Exit Sub

Erro_Almoxarifados_DblClick:

    Select Case Err

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168728)

    End Select

    Exit Sub

End Sub


Private Sub AlmoxarifadoInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_AlmoxarifadoInicial_Validate

    If Len(Trim(AlmoxarifadoInicial.Text)) > 0 Then
   
        'Tenta ler o Almoxarifado (NomeReduzido ou Código)
        lErro = TP_Almoxarifado_Le_ComCodigo(AlmoxarifadoInicial, objAlmoxarifado)
        If lErro <> SUCESSO Then Error 37402

    End If
    
    Exit Sub

Erro_AlmoxarifadoInicial_Validate:

    Cancel = True


    Select Case Err

        Case 37402

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168729)

    End Select

End Sub


Private Sub AlmoxarifadoFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_AlmoxarifadoFinal_Validate

    If Len(Trim(AlmoxarifadoFinal.Text)) > 0 Then

        'Tenta ler o Almoxarifado (NomeReduzido ou Código)
        lErro = TP_Almoxarifado_Le_ComCodigo(AlmoxarifadoFinal, objAlmoxarifado)
        If lErro <> SUCESSO Then Error 37403

    End If
 
    Exit Sub

Erro_AlmoxarifadoFinal_Validate:

    Cancel = True


    Select Case Err

        Case 37403
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168730)

    End Select

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
         If lErro <> SUCESSO And lErro <> 22540 Then Error 37404

         If lErro <> SUCESSO Then Error 37405 'Categoria não está cadastrada

        'Lê os dados de itens de categorias de produto
        lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colCategoria)
        If lErro <> SUCESSO Then Error 37406

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

        Case 37404
            Categoria.SetFocus
            
        Case 37405
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_INEXISTENTE", Err)
            Categoria.SetFocus
            
        Case 37406

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168731)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168732)

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
            If lErro <> SUCESSO And lErro <> 22603 Then Error 37408

            If lErro <> SUCESSO Then Error 37409 'Item da Categoria não está cadastrado

        End If

    End If

    Exit Sub

Erro_ValorInicial_Click:

    Select Case Err
    
        
        Case 37408
            ValorInicial.SetFocus

        Case 37409
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_INEXISTENTE", Err, objCategoriaProdutoItem.sItem, objCategoriaProdutoItem.sCategoria)
            ValorInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168733)

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
            If lErro <> SUCESSO And lErro <> 22603 Then Error 37411

            If lErro <> SUCESSO Then Error 37412 'Item da Categoria não está cadastrado

        End If

    End If

    Exit Sub

Erro_ValorFinal_Click:

    Select Case Err
    
    
        Case 37411
            ValorFinal.SetFocus

        Case 37412
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_INEXISTENTE", Err, objCategoriaProdutoItem.sItem, objCategoriaProdutoItem.sCategoria)
            ValorFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168734)

    End Select

    Exit Sub

End Sub

Private Function Critica_Numero(sNumero As String) As Long

Dim lErro As Long

On Error GoTo Erro_Critica_Numero

    If sNumero <> "" Then
    
        lErro = Long_Critica(sNumero)
        If lErro <> SUCESSO Then Error 37413
 
        If CLng(sNumero) < 0 Then Error 37414
 
    End If
    
    Critica_Numero = SUCESSO

    Exit Function

Erro_Critica_Numero:

    Critica_Numero = Err

    Select Case Err

        Case 37413
            
        Case 37414
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", Err, sNumero)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168735)

    End Select

    Exit Function

End Function


Private Sub ProximoNumero_Validate(Cancel As Boolean)
Dim lErro As Long

On Error GoTo Erro_ProximoNumero_Validate

    If Len(Trim(ProximoNumero.Text)) <> 0 Then
    
        lErro = Critica_Numero(ProximoNumero.Text)
        If lErro <> SUCESSO Then Error 37415
    
    End If
    
    Exit Sub

Erro_ProximoNumero_Validate:

    Cancel = True


    Select Case Err

        Case 37415

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168736)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_ETIQUETA_INVENTARIO
    Set Form_Load_Ocx = Me
    Caption = "Etiquetas Para Inventário"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpEtiqInv"
    
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

Private Sub labelAlmoxarifadoFinal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(labelAlmoxarifadoFinal, Source, X, Y)
End Sub

Private Sub labelAlmoxarifadoFinal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(labelAlmoxarifadoFinal, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LabelAlmoxarifado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelAlmoxarifado, Source, X, Y)
End Sub

Private Sub LabelAlmoxarifado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelAlmoxarifado, Button, Shift, X, Y)
End Sub

