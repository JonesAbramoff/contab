VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpKardexOcx 
   ClientHeight    =   6795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8685
   KeyPreview      =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   8685
   Begin VB.Frame Frame2 
      Height          =   525
      Left            =   180
      TabIndex        =   42
      Top             =   735
      Width           =   5595
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
         TabIndex        =   2
         Top             =   180
         Width           =   915
      End
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
         TabIndex        =   1
         Top             =   180
         Width           =   1425
      End
   End
   Begin VB.Frame FrameTipoRelat 
      Caption         =   "Tipo de Relatório"
      Height          =   1185
      Left            =   150
      TabIndex        =   41
      Top             =   5040
      Width           =   5685
      Begin VB.OptionButton OptTNPoder 
         Caption         =   "Terceiros em Nosso Poder"
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
         Left            =   120
         TabIndex        =   14
         Top             =   810
         Width           =   3045
      End
      Begin VB.OptionButton OptNTerceiros 
         Caption         =   "Nossos em Terceiros"
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
         Left            =   120
         TabIndex        =   13
         Top             =   525
         Width           =   2715
      End
      Begin VB.OptionButton OptTradicional 
         Caption         =   "Tradicional"
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
         TabIndex        =   12
         Top             =   240
         Width           =   2565
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6420
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "RelOpKardexOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   630
         Picture         =   "RelOpKardexOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpKardexOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpKardexOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CheckBox CheckProdSemMovimento 
      Caption         =   "Lista produtos sem movimento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      TabIndex        =   15
      Top             =   6360
      Width           =   3600
   End
   Begin VB.ListBox Almoxarifados 
      Height          =   5130
      ItemData        =   "RelOpKardexOcx.ctx":0994
      Left            =   6030
      List            =   "RelOpKardexOcx.ctx":0996
      TabIndex        =   16
      Top             =   1080
      Width           =   2535
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
      Picture         =   "RelOpKardexOcx.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpKardexOcx.ctx":0A9A
      Left            =   1380
      List            =   "RelOpKardexOcx.ctx":0A9C
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   255
      Width           =   2916
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produtos"
      Height          =   1065
      Left            =   120
      TabIndex        =   33
      Top             =   1740
      Width           =   5655
      Begin MSMask.MaskEdBox ProdutoFinal 
         Height          =   315
         Left            =   750
         TabIndex        =   5
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
         Left            =   780
         TabIndex        =   4
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
         Left            =   2295
         TabIndex        =   37
         Top             =   255
         Width           =   3135
      End
      Begin VB.Label DescProdFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2295
         TabIndex        =   36
         Top             =   645
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
         Left            =   375
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   35
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
         TabIndex        =   34
         Top             =   690
         Width           =   360
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Categoria de Produtos"
      Height          =   1170
      Left            =   120
      TabIndex        =   23
      Top             =   2910
      Width           =   5670
      Begin VB.ComboBox Categoria 
         Height          =   315
         Left            =   2385
         TabIndex        =   7
         Top             =   270
         Width           =   2745
      End
      Begin VB.ComboBox ValorInicial 
         Height          =   315
         Left            =   600
         TabIndex        =   8
         Top             =   705
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
         TabIndex        =   6
         Top             =   300
         Width           =   855
      End
      Begin VB.ComboBox ValorFinal 
         Height          =   315
         Left            =   3315
         TabIndex        =   9
         Top             =   705
         Width           =   2100
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   15
         Left            =   360
         TabIndex        =   32
         Top             =   720
         Width           =   30
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
         Left            =   2880
         TabIndex        =   31
         Top             =   750
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
         TabIndex        =   30
         Top             =   750
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
         TabIndex        =   29
         Top             =   330
         Width           =   855
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   720
      Left            =   150
      TabIndex        =   24
      Top             =   4170
      Width           =   5670
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   1935
         TabIndex        =   25
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
         Left            =   915
         TabIndex        =   10
         Top             =   285
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
         Left            =   4560
         TabIndex        =   26
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
         TabIndex        =   11
         Top             =   285
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3120
         TabIndex        =   28
         Top             =   330
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
         TabIndex        =   27
         Top             =   315
         Width           =   345
      End
   End
   Begin MSMask.MaskEdBox Almoxarifado 
      Height          =   315
      Left            =   1410
      TabIndex        =   3
      Top             =   1350
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
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
      Left            =   180
      TabIndex        =   40
      Top             =   1380
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
      Left            =   675
      TabIndex        =   39
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
      Height          =   345
      Left            =   6030
      TabIndex        =   38
      Top             =   810
      Width           =   1185
   End
End
Attribute VB_Name = "RelOpKardexOcx"
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

Private Sub Form_Load()

Dim lErro As Long
Dim colCategoriaProduto As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objFiliais As AdmFiliais

On Error GoTo Erro_Form_Load

    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento

    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoInicial)
    If lErro <> SUCESSO Then Error 37135

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoFinal)
    If lErro <> SUCESSO Then Error 37136

    'carrega a ListBox Almoxarifados
    lErro = Carrega_Lista_Almoxarifado()
    If lErro <> SUCESSO Then Error 37138
    
    'Le as categorias de produto
    lErro = CF("CategoriasProduto_Le_Todas", colCategoriaProduto)
    If lErro <> SUCESSO And lErro <> 22542 Then Error 37139

    'Preenche CategoriaProduto
    For Each objCategoriaProduto In colCategoriaProduto

        Categoria.AddItem objCategoriaProduto.sCategoria

    Next
    
    '01/11/01 Marcelo
'    For Each objFiliais In gcolFiliais
'
'        If objFiliais.iCodFilial <> 0 Then
'            cmbFilial.AddItem objFiliais.iCodFilial & SEPARADOR & objFiliais.sNome
'            cmbFilial.ItemData(cmbFilial.NewIndex) = objFiliais.iCodFilial
'        End If
'
'    Next
    
    Call Define_Padrao

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = Err

    Select Case Err

        Case 37135, 37136, 37138, 37139

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169531)

    End Select

    Exit Sub

End Sub

Sub Define_Padrao()

Dim lErro As Long

On Error GoTo Erro_Define_Padrao

    giProdInicial = 1
    
    TodasCategorias_Click
    
    TodasCategorias = 1
    
    CheckProdSemMovimento.Value = 0
    
    '''08/10/01 Marcelo
    OptTradicional.Value = True
    
'    '01/11/01 Marcelo
    OptAlmoxarifado.Value = True
       
    Exit Sub

Erro_Define_Padrao:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169532)

    End Select

    Exit Sub

End Sub

Private Sub DataFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFinal)

End Sub

Private Sub DataInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicial)

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
        If lErro <> SUCESSO Then gError 82378

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 82378

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169533)

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
        If lErro <> SUCESSO Then gError 82379

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 82379

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169534)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82380

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82381

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 82382

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 82380, 82382

        Case 82381
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169535)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82383

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82384

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 82385

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 82383, 82385
        
        Case 82384
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169536)

    End Select

    Exit Sub

End Sub

Private Sub OptAlmoxarifado_Click()
    
    'lblFilial.Visible = False
    'cmbFilial.Visible = False
    'lblAlmoxarifado.Visible = True
    Almoxarifado.Enabled = True
    Almoxarifados.Enabled = True
End Sub

Private Sub OptFilial_Click()
    
    'lblFilial.Visible = True
    'cmbFilial.Visible = True
    'lblAlmoxarifado.Visible = False
    Almoxarifado.Enabled = False
    Almoxarifados.Enabled = False
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169537)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169538)

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
    If lErro Then Error 37143
    
    sParam = String(255, 0)
    
        'pega parâmetro Almoxarifado e exibe
        lErro = objRelOpcoes.ObterParametro("NALMOX", sParam)
        If lErro Then Error 37144
        
        'If sParam <> "" Then
            Almoxarifado.Text = sParam
            Call Almoxarifado_Validate(bSGECancelDummy)
            
            '01/11/01 Marcelo INICIO
            'pega parâmetro OptAlmoxarifado e exibe
            lErro = objRelOpcoes.ObterParametro("TALMOXARIFADO1", sParam)
            If lErro Then Error 93706
            
            If sParam <> "" Then
                OptAlmoxarifado.Value = sParam
            End If
            
        'Else
                  
            lErro = objRelOpcoes.ObterParametro("NFILIALEMP", sParam)
            If lErro Then gError 93730
                        
            '01/11/01 Marcelo
            'pega parâmetro OptFilial e exibe
            lErro = objRelOpcoes.ObterParametro("TFILIAL1", sParam)
            If lErro Then Error 93707
            
            If sParam <> "" Then
                OptFilial.Value = sParam
            End If
            
'            lErro = objRelOpcoes.ObterParametro("TFILIALDESC", sParam)
'            If lErro Then gError 93709
        
            'cmbFilial.Text = sParam
        
        'End If
   
    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro Then Error 37145

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then Error 37146

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro Then Error 37147

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then Error 37148
    
    'pega parâmetro TodasCategorias e exibe
    lErro = objRelOpcoes.ObterParametro("NTODASCAT", sParam)
    If lErro Then Error 37149

    TodasCategorias.Value = CInt(sParam)

    '09/10/01 Marcelo INICIO
    'pega parâmetro OptTradicional e exibe
    lErro = objRelOpcoes.ObterParametro("TRELTRADICIONAL", sParam)
    If lErro Then Error 93666
    
    If sParam <> "" Then
        OptTradicional.Value = sParam
    End If
    
    'pega parâmetro OptTradicional e exibe
    lErro = objRelOpcoes.ObterParametro("TRELNTERCEIROS", sParam)
    If lErro Then Error 93667
    
    If sParam <> "" Then
        OptNTerceiros.Value = sParam
    End If
    
    'pega parâmetro OptTradicional e exibe
    lErro = objRelOpcoes.ObterParametro("TRELTNPODER", sParam)
    If lErro Then Error 93668
    
    If sParam <> "" Then
        OptTNPoder.Value = sParam
    End If
    
    '09/10/01 Marcelo FIM
                
    'pega parâmetro categoria de produto e exibe
    lErro = objRelOpcoes.ObterParametro("TCATPROD", sParam)
    If lErro Then gError 37150
    
    Categoria.Text = sParam

    'pega parâmetro valor inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TITEMCATPRODINI", sParam)
    If lErro Then gError 37151
    
    ValorInicial.Text = sParam
    
    'pega parâmetro Valor Final e exibe
    lErro = objRelOpcoes.ObterParametro("TITEMCATPRODFIM", sParam)
    If lErro Then gError 37152
    
    ValorFinal.Text = sParam
 
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then Error 37153

    Call DateParaMasked(DataInicial, CDate(sParam))
    'DataInicial.PromptInclude = False
    'DataInicial.Text = sParam
    'DataInicial.PromptInclude = True

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then Error 37154

    Call DateParaMasked(DataFinal, CDate(sParam))
    'DataFinal.PromptInclude = False
    'DataFinal.Text = sParam
    'DataFinal.PromptInclude = True
        
    'pega parâmetro p/ listar produto sem movimento e exibe
    lErro = objRelOpcoes.ObterParametro("NPRODSEMMOV", sParam)
    If lErro Then Error 37155

    CheckProdSemMovimento.Value = CInt(sParam)
    
       
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 37143 To 37155, 93666 To 93668, 93706, 93707, 93709, 93730

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169539)

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

    If Not (gobjRelatorio Is Nothing) Then Error 29898
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 37133

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 37133
        
        Case 29898
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169540)

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
    
    ComboOpcoes.SetFocus

End Sub

Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String) As Long
'Formata os produtos retornando em sProd_I e sProd_F
'Verifica se os parâmetros iniciais são maiores que os finais

Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then Error 37157

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then Error 37158

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then Error 37159

    End If

   'data inicial não pode ser maior que a data final
    If Trim(DataInicial.ClipText) <> "" And Trim(DataFinal.ClipText) <> "" Then
    
         If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then Error 37160
    
    End If
    
    'valor inicial não pode ser maior que o valor final
    If Trim(ValorInicial.Text) <> "" And Trim(ValorFinal.Text) <> "" Then
    
        If ValorInicial.Text > ValorFinal.Text Then Error 37161
         
    Else
        
        If Trim(ValorInicial.Text) = "" And Trim(ValorFinal.Text) = "" And TodasCategorias.Value = 0 Then Error 37994
         
    End If
        
    '01/11/01 Marcelo
    If OptAlmoxarifado.Value = True Then
    
        'O campo almoxarifado deve ser preenchido
        If Trim(Almoxarifado.Text) = "" Then Error 37995
    Else
        'O campo filial deve ser preenchido
        'If Trim(cmbFilial.Text) = "" Then Error 93708
    End If
    
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                
'        Case 93708
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
'            cmbFilial.SetFocus
              
        Case 37157
            ProdutoInicial.SetFocus

        Case 37158
            ProdutoFinal.SetFocus

        Case 37159
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            ProdutoInicial.SetFocus
             
        Case 37160
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataInicial.SetFocus
            
        Case 37161
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_INICIAL_MAIOR", gErr)
            ValorInicial.SetFocus
            
        Case 37994
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_NAO_INFORMADO", gErr)
            ValorInicial.SetFocus
        
        Case 37995
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_PREENCHIDO1", gErr)
            Almoxarifado.SetFocus
      
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169541)

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

Function Traz_Produto_Tela(sProduto As String) As Long
'verifica e preenche o produto inicial e final com sua descriçao de acordo com o último foco
'sProduto deve estar no formato do BD

Dim lErro As Long

On Error GoTo Erro_Traz_Produto_Tela

    If giProdInicial Then

        lErro = CF("Traz_Produto_MaskEd", sProduto, ProdutoInicial, DescProdInic)
        If lErro <> SUCESSO Then Error 37165

    Else

        lErro = CF("Traz_Produto_MaskEd", sProduto, ProdutoFinal, DescProdFim)
        If lErro <> SUCESSO Then Error 37166

    End If

    Traz_Produto_Tela = SUCESSO

    Exit Function

Erro_Traz_Produto_Tela:

    Traz_Produto_Tela = Err

    Select Case Err

        Case 37165, 37166

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169542)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sProd_I As String
Dim sProd_F As String
Dim sAlmox As String
Dim iAlmox As Integer
Dim iFilial As Integer
Dim objEstoqueMes As New ClassEstoqueMes

On Error GoTo Erro_PreencherRelOp

    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)
    
    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F)
    If lErro <> SUCESSO Then Error 37167

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 37168
         
    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then Error 37169

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then Error 37170
    
    If OptAlmoxarifado.Value = True Then
    
        iAlmox = Codigo_Extrai(Almoxarifado.Text)
    
        lErro = objRelOpcoes.IncluirParametro("NALMOX", CStr(iAlmox))
        If lErro <> AD_BOOL_TRUE Then Error 37171
            
        lErro = objRelOpcoes.IncluirParametro("TALMOXARIFADO", Almoxarifado.Text)
        If lErro <> AD_BOOL_TRUE Then Error 54501
        
        '01/11/01 - Marcelo
        lErro = objRelOpcoes.IncluirParametro("TALMOXARIFADO1", CStr(OptAlmoxarifado.Value))
        If lErro <> AD_BOOL_TRUE Then gError 93704
                       
    Else
                                      
        lErro = objRelOpcoes.IncluirParametro("NFILIALEMP", giFilialEmpresa)
        If lErro <> AD_BOOL_TRUE Then gError 93729
      
        '01/11/01 - Marcelo
        lErro = objRelOpcoes.IncluirParametro("TFILIAL1", CStr(OptFilial.Value))
        If lErro <> AD_BOOL_TRUE Then gError 93705
                
    End If
        
    
    If DataInicial.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINIC", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 37172

    If DataFinal.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 37173
       
    lErro = objRelOpcoes.IncluirParametro("NTODASCAT", CStr(TodasCategorias.Value))
    If lErro <> AD_BOOL_TRUE Then Error 37174
    
    lErro = objRelOpcoes.IncluirParametro("TCATPROD", Categoria.Text)
    If lErro <> AD_BOOL_TRUE Then Error 37175
    
    lErro = objRelOpcoes.IncluirParametro("TITEMCATPRODINI", ValorInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 37176
    
    lErro = objRelOpcoes.IncluirParametro("TITEMCATPRODFIM", ValorFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 37177
    
    lErro = objRelOpcoes.IncluirParametro("NPRODSEMMOV", CStr(CheckProdSemMovimento.Value))
    If lErro <> AD_BOOL_TRUE Then Error 37178
    
    '09/10/01 - Marcelo INICIO
    lErro = objRelOpcoes.IncluirParametro("TRELTRADICIONAL", CStr(OptTradicional.Value))
    If lErro <> AD_BOOL_TRUE Then gError 93663

    lErro = objRelOpcoes.IncluirParametro("TRELNTERCEIROS", CStr(OptNTerceiros.Value))
    If lErro <> AD_BOOL_TRUE Then gError 93664
    
    lErro = objRelOpcoes.IncluirParametro("TRELTNPODER", CStr(OptTNPoder.Value))
    If lErro <> AD_BOOL_TRUE Then gError 93665

    '09/10/01 - Marcelo FIM
            
    objEstoqueMes.iFilialEmpresa = giFilialEmpresa
    
    lErro = CF("EstoqueMes_Le_Apurado", objEstoqueMes)
    If lErro <> SUCESSO And lErro <> 46225 Then Error 54515

    If lErro = 46225 Then
        objEstoqueMes.iAno = 0
        objEstoqueMes.iMes = 0
    End If
        
    lErro = objRelOpcoes.IncluirParametro("NANOAPURADO", objEstoqueMes.iAno)
    If lErro <> AD_BOOL_TRUE Then Error 54516
 
    lErro = objRelOpcoes.IncluirParametro("NMESAPURADO", objEstoqueMes.iMes)
    If lErro <> AD_BOOL_TRUE Then Error 54517
    
    '01/11/01 Marcelo
    If OptAlmoxarifado.Value = True Then
         '''08/10/01 Marcelo incluido o Tipo de Relatório chamado
        If OptTradicional.Value = True Then
                        
            If TodasCategorias.Value = 0 Then
            
                gobjRelatorio.sNomeTsk = "KardexCa"
            Else
                gobjRelatorio.sNomeTsk = "Kardex"
            End If
        
        ElseIf OptNTerceiros.Value = True Then
              
            If TodasCategorias.Value = 0 Then
                gobjRelatorio.sNomeTsk = "KardNTCa"
            Else
                gobjRelatorio.sNomeTsk = "KardexNT"
            End If
            
        ElseIf OptTNPoder.Value = True Then
            
            If TodasCategorias.Value = 0 Then
                gobjRelatorio.sNomeTsk = "KarTNPCa"
            Else
                gobjRelatorio.sNomeTsk = "KardeTNP"
            End If
            
        End If
    Else
        If OptTradicional.Value = True Then

            If TodasCategorias.Value = 0 Then

                gobjRelatorio.sNomeTsk = "KarCaFil"
            Else
                gobjRelatorio.sNomeTsk = "KardeFil"
            End If

        ElseIf OptNTerceiros.Value = True Then

            If TodasCategorias.Value = 0 Then
                gobjRelatorio.sNomeTsk = "Kantcfil"
            Else
                gobjRelatorio.sNomeTsk = "KarNTFil"
            End If

        ElseIf OptTNPoder.Value = True Then

            If TodasCategorias.Value = 0 Then
                gobjRelatorio.sNomeTsk = "KaTNPCFi"
            Else
                gobjRelatorio.sNomeTsk = "KaTNPFil"
            End If

        End If
    End If

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sAlmox, sProd_I, sProd_F)
    If lErro <> SUCESSO Then Error 37179

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 37167 To 37179, 93663 To 93665
        
        Case 93704, 93705, 93729
        
        Case 54501, 54515, 54516, 54517

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169543)

    End Select

    Exit Function

End Function


Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 37180

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 37181

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex
        
        'cmbFilial.ListIndex = -1
        
        'limpa as opções da tela
        Limpar_Tela
        '01/11/01 Marcelo
        Call Define_Padrao
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 37180
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 37181

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169544)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 37182

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 37182

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169545)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 37183

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then Error 37184

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 37185

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 37183
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 37184, 37185

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169546)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate

    giProdInicial = 0

    lErro = CF("Produto_Perde_Foco", ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO And lErro <> 27095 Then Error 37186
    
    If lErro <> SUCESSO Then Error 43257

    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True


    Select Case Err

        Case 37186

         Case 43257
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169547)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    giProdInicial = 1

    lErro = CF("Produto_Perde_Foco", ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then Error 37187
    
    If lErro <> SUCESSO Then Error 43258

    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True


    Select Case Err

        Case 37187

         Case 43258
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169548)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sAlmox As String, sProd_I As String, sProd_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""
    
    If sProd_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = "Produto >= " & Forprint_ConvTexto(sProd_I)

    End If

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
    
    If Trim(DataInicial.ClipText) <> "" Or Trim(DataFinal.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        
        sExpressao = sExpressao & "((TipoRegRel = 2) OU ("
                
        If Trim(DataInicial.ClipText) <> "" Then
        
            sExpressao = sExpressao & "Data >= " & Forprint_ConvData(CDate(DataInicial.Text))

        End If
    
        If Trim(DataFinal.ClipText) <> "" Then
    
            If Trim(DataInicial.ClipText) <> "" Then sExpressao = sExpressao & " E "
            sExpressao = sExpressao & "Data <= " & Forprint_ConvData(CDate(DataFinal.Text))
    
        End If
        
        sExpressao = sExpressao & "))"
        
    End If
    
    If CheckProdSemMovimento.Value = 0 Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "TipoRegRel = 1"

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169549)

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
        If lErro <> SUCESSO Then Error 37188

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True


    Select Case Err

        Case 37188

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169550)

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
        If lErro <> SUCESSO Then Error 37189

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True


    Select Case Err

        Case 37189

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169551)

    End Select

    Exit Sub

End Sub


Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 37190

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 37190
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169552)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 37191

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 37191
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169553)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 37192

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case Err

        Case 37192
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169554)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 37193

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case Err

        Case 37193
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169555)

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
    If lErro <> SUCESSO Then Error 37195

    'Preenche a ListBox AlmoxarifadoList com os objetos da coleção
    For Each objAlmoxarifado In colAlmoxarifados
        Almoxarifados.AddItem objAlmoxarifado.iCodigo & SEPARADOR & objAlmoxarifado.sNomeReduzido
        Almoxarifados.ItemData(Almoxarifados.NewIndex) = objAlmoxarifado.iCodigo
    Next

    Carrega_Lista_Almoxarifado = SUCESSO

    Exit Function

Erro_Carrega_Lista_Almoxarifado:

    Carrega_Lista_Almoxarifado = Err

    Select Case Err

        Case 37195

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169556)

    End Select

    Exit Function

End Function

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

    Select Case Err

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 169557)

    End Select

    Exit Sub

End Sub

Private Sub Almoxarifado_Validate(Cancel As Boolean)

'Dim lErro As Long
'Dim objAlmoxarifado As New ClassAlmoxarifado
'
'On Error GoTo Erro_Almoxarifado_Validate
'
'    If Len(Trim(Almoxarifado.Text)) > 0 Then
'
'        'Tenta ler o Almoxarifado (NomeReduzido ou Código)
'        lErro = TP_Almoxarifado_Le_ComCodigo(Almoxarifado, objAlmoxarifado)
'        If lErro <> SUCESSO Then Error 37196
'
'    End If
'
'    Exit Sub
'
'Erro_Almoxarifado_Validate:
'
'    Cancel = True
'
'
'    Select Case Err
'
'        Case 37196
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 169558)
'
'    End Select

'09/11/01 Marcelo INICIO

Dim lErro As Long

Dim objAlmoxarifado As New ClassAlmoxarifado
'Dim sContaEnxuta As String

On Error GoTo Erro_Almoxarifado_Validate

    If Len(Trim(Almoxarifado.Text)) > 0 Then
        
        'Le o almoxarifado pelo código ou pelo nome reduzido e joga o nome reduzido em Almoxarifado.Text
        lErro = TP_Almoxarifado_Le(Almoxarifado, objAlmoxarifado)
        If lErro <> SUCESSO Then Error 37196
                
        Almoxarifado = objAlmoxarifado.iCodigo & SEPARADOR & objAlmoxarifado.sNomeReduzido
        
        'se o almoxarifado não pertencer a filial em questão ==> erro
        If objAlmoxarifado.iFilialEmpresa <> giFilialEmpresa Then gError 93740
        
    End If
    
    Exit Sub

Erro_Almoxarifado_Validate:

    Cancel = True


    Select Case gErr

        Case 37196

        Case 93740
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_FILIAL_DIFERENTE", gErr, objAlmoxarifado.iCodigo, giFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 169559)

    End Select

    Exit Sub

'09/11/01 Marcelo FIM
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
         If lErro <> SUCESSO And lErro <> 22540 Then Error 37197

         If lErro <> SUCESSO Then Error 37198 'Categoria não está cadastrada

        'Lê os dados de itens de categorias de produto
        lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colCategoria)
        If lErro <> SUCESSO Then Error 37199

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

        Case 37197
            Categoria.SetFocus
            
        Case 37198
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_INEXISTENTE", Err)
            Categoria.SetFocus
            
        Case 37199

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169560)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169561)

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
            If lErro <> SUCESSO And lErro <> 22603 Then Error 37200

            If lErro <> SUCESSO Then Error 37201 'Item da Categoria não está cadastrado

        End If

    End If

    Exit Sub

Erro_ValorInicial_Click:

    Select Case Err

        Case 37200
            ValorInicial.SetFocus

        Case 37201
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_INEXISTENTE", Err, objCategoriaProdutoItem.sItem, objCategoriaProdutoItem.sCategoria)
            ValorInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169562)

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
            If lErro <> SUCESSO And lErro <> 22603 Then Error 37202

            If lErro <> SUCESSO Then Error 37203 'Item da Categoria não está cadastrado

        End If

    End If

    Exit Sub

Erro_ValorFinal_Click:

    Select Case Err

        Case 37202
            ValorFinal.SetFocus

        Case 37203
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_INEXISTENTE", Err, objCategoriaProdutoItem.sItem, objCategoriaProdutoItem.sCategoria)
            ValorFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169563)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_KARDEX
    Set Form_Load_Ocx = Me
    Caption = "Kardex (Lista dos Movimentos)"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpKardex"
    
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

Private Sub LblAlmoxarifado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(lblAlmoxarifado, Source, X, Y)
End Sub

Private Sub LblAlmoxarifado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(lblAlmoxarifado, Button, Shift, X, Y)
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


