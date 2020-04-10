VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpPedProdutoOcx 
   ClientHeight    =   5175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7980
   KeyPreview      =   -1  'True
   ScaleHeight     =   5175
   ScaleWidth      =   7980
   Begin VB.Frame Frame3 
      Caption         =   "Categoria de Produtos"
      Height          =   1785
      Left            =   120
      TabIndex        =   25
      Top             =   3360
      Width           =   5670
      Begin VB.ComboBox ValorFinal 
         Height          =   315
         Left            =   3420
         TabIndex        =   29
         Top             =   1215
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
         TabIndex        =   28
         Top             =   300
         Width           =   855
      End
      Begin VB.ComboBox ValorInicial 
         Height          =   315
         Left            =   720
         TabIndex        =   27
         Top             =   1230
         Width           =   1950
      End
      Begin VB.ComboBox Categoria 
         Height          =   315
         Left            =   1650
         TabIndex        =   26
         Top             =   660
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
         TabIndex        =   33
         Top             =   705
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
         Left            =   315
         TabIndex        =   32
         Top             =   1275
         Width           =   420
      End
      Begin VB.Label Label2 
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
         Left            =   2970
         TabIndex        =   31
         Top             =   1260
         Width           =   555
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   15
         Left            =   360
         TabIndex        =   30
         Top             =   720
         Width           =   30
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5715
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpPedProdutoOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpPedProdutoOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpPedProdutoOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpPedProdutoOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produtos"
      Height          =   1800
      Left            =   120
      TabIndex        =   18
      Top             =   645
      Width           =   5550
      Begin VB.ComboBox ComboNivel 
         Height          =   315
         ItemData        =   "RelOpPedProdutoOcx.ctx":0994
         Left            =   375
         List            =   "RelOpPedProdutoOcx.ctx":099E
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1320
         Width           =   2940
      End
      Begin MSMask.MaskEdBox ProdutoInicial 
         Height          =   315
         Left            =   735
         TabIndex        =   1
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
         Left            =   735
         TabIndex        =   2
         Top             =   840
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Nivel 
         Height          =   300
         Left            =   4800
         TabIndex        =   4
         Top             =   1350
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         Caption         =   "Até o Nível:"
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
         Left            =   3690
         TabIndex        =   23
         Top             =   1380
         Width           =   1110
      End
      Begin VB.Label DescProdFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2295
         TabIndex        =   22
         Top             =   840
         Width           =   3000
      End
      Begin VB.Label DescProdInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2295
         TabIndex        =   21
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
         Left            =   345
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   20
         Top             =   405
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
         Left            =   315
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   19
         Top             =   885
         Width           =   435
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpPedProdutoOcx.ctx":09CD
      Left            =   1035
      List            =   "RelOpPedProdutoOcx.ctx":09CF
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
      Left            =   5955
      Picture         =   "RelOpPedProdutoOcx.ctx":09D1
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   810
      Width           =   1575
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   750
      Left            =   135
      TabIndex        =   13
      Top             =   2580
      Width           =   5505
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   1890
         TabIndex        =   14
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
         Left            =   930
         TabIndex        =   5
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
         Left            =   4515
         TabIndex        =   15
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
         Left            =   3540
         TabIndex        =   6
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
         Left            =   3135
         TabIndex        =   17
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
         Left            =   540
         TabIndex        =   16
         Top             =   315
         Width           =   345
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
      Left            =   345
      TabIndex        =   24
      Top             =   315
      Width           =   615
   End
End
Attribute VB_Name = "RelOpPedProdutoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Const TODOS_OS_NIVEIS = 0
Const UM_NIVEL = 1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1

Dim giProdInicial As Integer

Public Sub Form_Load()

Dim lErro As Long
Dim colCategoriaProduto As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto

On Error GoTo Erro_Form_Load

    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento

    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoInicial)
    If lErro <> SUCESSO Then Error 37932

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoFinal)
    If lErro <> SUCESSO Then Error 37933

    Call Define_Padrao
                  
   'Le as categorias de produto
    lErro = CF("CategoriasProduto_Le_Todas", colCategoriaProduto)
    If lErro <> SUCESSO And lErro <> 22542 Then Error 47330

    'Preenche CategoriaProduto
    For Each objCategoriaProduto In colCategoriaProduto

        Categoria.AddItem objCategoriaProduto.sCategoria

    Next
    
    TodasCategorias_Click
    TodasCategorias.Value = 1
                      
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = Err

    Select Case Err

        Case 37932, 37933, 37935

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171020)

    End Select

    Exit Sub

End Sub

Private Function Define_Padrao() As Long

Dim lErro As Long

On Error GoTo Erro_Define_Padrao
        
    giProdInicial = 1
    
    ComboNivel.ListIndex = TODOS_OS_NIVEIS
    
    Nivel.Enabled = False
         
    Define_Padrao = SUCESSO

    Exit Function

Erro_Define_Padrao:

    Define_Padrao = Err

    Select Case Err
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171021)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro Then Error 37936
      
    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro Then Error 37939

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then Error 37940

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro Then Error 37941

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then Error 37942
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then Error 37943

    Call DateParaMasked(DataInicial, CDate(sParam))
    'DataInicial.PromptInclude = False
    'DataInicial.Text = sParam
    'DataInicial.PromptInclude = True

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then Error 37944

    Call DateParaMasked(DataFinal, CDate(sParam))
    'DataFinal.PromptInclude = False
    'DataFinal.Text = sParam
    'DataFinal.PromptInclude = True
    
    'pega parâmetro Tipo de Nivel e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("NTIPONIVELPROD", sParam)
    If lErro Then Error 37945
   
    ComboNivel.ListIndex = CInt(sParam)
    
    'pega parâmetro Nivel e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("NNIVELPROD", sParam)
    If lErro Then Error 37946
   
    Nivel.Text = sParam
    
 'pega parâmetro TodasCategorias e exibe
    lErro = objRelOpcoes.ObterParametro("NTODASCAT", sParam)
    If lErro <> SUCESSO Then Error 47336

    TodasCategorias.Value = CInt(sParam)

    'pega parâmetro categoria de produto e exibe
    lErro = objRelOpcoes.ObterParametro("TCATPROD", sParam)
    If lErro <> SUCESSO Then Error 47337
    
    If sParam <> "" Then
    
        Categoria.Text = sParam
    
        Categoria.Text = sParam
        Call Categoria_Validate(bSGECancelDummy)
    
        'pega parâmetro valor inicial e exibe
        lErro = objRelOpcoes.ObterParametro("TITEMCATPRODINI", sParam)
        If lErro <> SUCESSO Then Error 47338
        
        ValorInicial.Text = sParam
        ValorInicial.Enabled = True
        
        'pega parâmetro Valor Final e exibe
        lErro = objRelOpcoes.ObterParametro("TITEMCATPRODFIM", sParam)
        If lErro <> SUCESSO Then Error 47339
    
        ValorFinal.Text = sParam
        ValorFinal.Enabled = True
    End If

    'pega parâmetro valor inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TITEMCATPRODINI", sParam)
    If lErro <> SUCESSO Then Error 47340
    
    ValorInicial.Text = sParam
    
    'pega parâmetro Valor Final e exibe
    lErro = objRelOpcoes.ObterParametro("TITEMCATPRODFIM", sParam)
    If lErro <> SUCESSO Then Error 47341
    
    ValorFinal.Text = sParam
        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 37936 To 37946

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171022)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 29884
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 37930
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 37930
        
        Case 29884
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171023)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    
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
         If lErro <> SUCESSO And lErro <> 22540 Then Error 47373

         If lErro <> SUCESSO Then Error 47374 'Categoria não está cadastrada

        'Lê os dados de itens de categorias de produto
        lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colCategoria)
        If lErro <> SUCESSO Then Error 47375

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

        Case 47373
            Categoria.SetFocus
            
        Case 47374
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_INEXISTENTE", Err)
            Categoria.SetFocus
            
        Case 47375

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171024)

    End Select

    Exit Sub

End Sub

Private Sub Categoria_GotFocus()

    'desmarca todasCategorias
    TodasCategorias.Value = 0

End Sub


Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82574

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82575
    
    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 82576

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 82574, 82576

        Case 82575
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171025)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82577

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82578

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 82579

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 82577, 82579

        Case 82578
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171026)

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
        If lErro <> SUCESSO Then gError 82591

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 82591

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171027)

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
        If lErro <> SUCESSO Then gError 82590

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 82590

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171028)

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
    If lErro <> SUCESSO Then Error 43181
    
    ComboOpcoes.Text = ""
    DescProdInic.Caption = ""
    DescProdFim.Caption = ""
    TodasCategorias_Click
    TodasCategorias = 0
    
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then Error 43217
    
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 43181, 43217
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171029)

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
Dim sNivel As String
Dim sTipoNivel As String

On Error GoTo Erro_PreencherRelOp

    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)
       
    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F)
    If lErro <> SUCESSO Then Error 37950
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 37951
    
    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then Error 37952

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then Error 37953
        
    If DataInicial.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINIC", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 37956

    If DataFinal.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 37957
    
    sTipoNivel = CStr(ComboNivel.ListIndex)
    lErro = objRelOpcoes.IncluirParametro("NTIPONIVELPROD", sTipoNivel)
    If lErro <> AD_BOOL_TRUE Then Error 37958
    
    sNivel = Nivel.Text
    
    If sNivel = "" Then sNivel = "0"
        
    lErro = objRelOpcoes.IncluirParametro("NNIVELPROD", sNivel)
    If lErro <> AD_BOOL_TRUE Then Error 37959
        
     lErro = objRelOpcoes.IncluirParametro("NTODASCAT", CStr(TodasCategorias.Value))
    If lErro <> AD_BOOL_TRUE Then Error 47358
    
    lErro = objRelOpcoes.IncluirParametro("TCATPROD", Categoria.Text)
    If lErro <> AD_BOOL_TRUE Then Error 47359
    
    lErro = objRelOpcoes.IncluirParametro("TITEMCATPRODINI", ValorInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 47360
    
    lErro = objRelOpcoes.IncluirParametro("TITEMCATPRODFIM", ValorFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 47361
    
    If TodasCategorias.Value = 0 Then
        
        gobjRelatorio.sNomeTsk = "rlvndpca"
    Else
        gobjRelatorio.sNomeTsk = "relvendp"
    End If
       
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd_I, sProd_F)
    If lErro <> SUCESSO Then Error 37960
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 37950 To 37960

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171030)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 37961

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 37962

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'Limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then Error 43182
    
        ComboOpcoes.Text = ""
        DescProdInic.Caption = ""
        DescProdFim.Caption = ""
        TodasCategorias_Click
        TodasCategorias = 0
        
        lErro = Define_Padrao()
        If lErro <> SUCESSO Then Error 43818
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 37961
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 37962, 43218

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171031)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 37963

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 37963

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171032)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 37964
    
    If ComboNivel.ListIndex = UM_NIVEL And Len(Trim(Nivel.Text)) = 0 Then Error 37990

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then Error 37965

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 37966

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 43183
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 37964
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 37965, 37966, 43183

        Case 37990
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NIVEL_NAO_INFORMADO", Err)
            Nivel.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171033)

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
    
    If ComboNivel.ListIndex = 1 Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Nivel <= " & Forprint_ConvInt(CInt(Nivel.Text))

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171034)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String) As Long
'Formata os produtos retornando em sProd_I e sProd_F
'Verifica se os parâmetros iniciais são maiores que os finais

Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
   
    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then Error 37967

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then Error 37968

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then Error 37969

    End If
    
   
    'data inicial não pode ser maior que a data final
    If Trim(DataInicial.ClipText) <> "" And Trim(DataFinal.ClipText) <> "" Then
    
         If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then Error 37971
    
    End If
        
   'valor inicial não pode ser maior que o valor final
    If Trim(ValorInicial.Text) <> "" And Trim(ValorFinal.Text) <> "" Then
    
         If ValorInicial.Text > ValorFinal.Text Then Error 47346
         
     Else
        
        If Trim(ValorInicial.Text) = "" And Trim(ValorFinal.Text) = "" And TodasCategorias.Value = 0 Then Error 47347
    
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function


Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = Err

    Select Case Err
                     
         Case 37967
            ProdutoInicial.SetFocus

        Case 37968
            ProdutoFinal.SetFocus

        Case 37969
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", Err)
            ProdutoInicial.SetFocus
              
         Case 37971
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)
            DataInicial.SetFocus
                              
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171035)

    End Select

    Exit Function

End Function


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
        If lErro <> SUCESSO Then Error 37975

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True


    Select Case Err

        Case 37975

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171036)

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
        If lErro <> SUCESSO Then Error 37976

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True


    Select Case Err

        Case 37976

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171037)

    End Select

    Exit Sub

End Sub

Private Sub TodasCategorias_Click()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_TodasCategorias_Click
        
    Categoria.Text = ""
    ValorInicial.Text = ""
    ValorFinal.Text = ""
    
    Exit Sub

Erro_TodasCategorias_Click:

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171038)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 37977

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 37977
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171039)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 37978

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 37978
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171040)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 37979

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case Err

        Case 37979
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171041)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 37980

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case Err

        Case 37980
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171042)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate

    giProdInicial = 0

    lErro = CF("Produto_Perde_Foco", ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO And lErro <> 27095 Then Error 37988
    
    If lErro <> SUCESSO Then Error 43249

    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True


    Select Case Err

        Case 37988

        Case 43249
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171043)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    giProdInicial = 1

    lErro = CF("Produto_Perde_Foco", ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then Error 37989
    
    If lErro <> SUCESSO Then Error 43250

    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True


    Select Case Err

        Case 37989

        Case 43250
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171044)

    End Select

    Exit Sub

End Sub

Private Sub Nivel_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iNivel As Integer

On Error GoTo Erro_Nivel_Validate

    If Len(Trim(Nivel.Text)) = 0 Then Exit Sub
       
    lErro = Inteiro_Critica(Nivel.Text)
    If lErro <> SUCESSO Then Error 37991
    
    iNivel = CInt(Nivel.Text)
    If iNivel < 0 Then Error 37992
   
    Exit Sub

Erro_Nivel_Validate:

    Cancel = True


    Select Case Err

        Case 37991
            
        Case 37992
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NEGATIVO", Err, iNivel)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171045)

    End Select

    Exit Sub

End Sub

Private Sub ComboNivel_Click()

Dim lErro As Long

On Error GoTo Erro_ComboNivel_Click

    If ComboNivel.ListIndex = 1 Then
        Nivel.Enabled = True
        Nivel.SetFocus
    Else
        Nivel.Text = ""
        Nivel.Enabled = False
    End If

    Exit Sub

Erro_ComboNivel_Click:

    Select Case Err

        Case Else
        
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171046)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_PEDIDOS_PRODUTO
    Set Form_Load_Ocx = Me
    Caption = "Pedidos por Produto"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpPedProduto"
    
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



Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
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

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
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
            If lErro <> SUCESSO And lErro <> 22603 Then Error 47376

            'Item da Categoria não está cadastrado
            If lErro <> SUCESSO Then Error 47377
            
        End If

    End If

    Exit Sub

Erro_ValorInicial_Click:

    Select Case Err

        Case 47376
            ValorInicial.SetFocus

        Case 47377
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_INEXISTENTE", Err, objCategoriaProdutoItem.sItem, objCategoriaProdutoItem.sCategoria)
            ValorInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171047)

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
            If lErro <> SUCESSO And lErro <> 22603 Then Error 47378
                                    
            'Item da Categoria não está cadastrado
            If lErro <> SUCESSO Then Error 47379
        End If

    End If

    Exit Sub

Erro_ValorFinal_Click:

    Select Case Err

        Case 47378
            ValorFinal.SetFocus

        Case 47379
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_INEXISTENTE", Err, objCategoriaProdutoItem.sItem, objCategoriaProdutoItem.sCategoria)
            ValorFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171048)

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

