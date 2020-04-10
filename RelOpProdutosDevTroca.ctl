VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpProdutosDevTroca 
   ClientHeight    =   6390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7770
   KeyPreview      =   -1  'True
   ScaleHeight     =   6390
   ScaleMode       =   0  'User
   ScaleWidth      =   7770
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
      Left            =   5760
      Picture         =   "RelOpProdutosDevTroca.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   960
      Width           =   1605
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5452
      ScaleHeight     =   495
      ScaleWidth      =   2130
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   120
      Width           =   2190
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1650
         Picture         =   "RelOpProdutosDevTroca.ctx":0102
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1125
         Picture         =   "RelOpProdutosDevTroca.ctx":0280
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "RelOpProdutosDevTroca.ctx":07B2
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpProdutosDevTroca.ctx":093C
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   735
      Left            =   240
      TabIndex        =   34
      Top             =   840
      Width           =   5175
      Begin MSComCtl2.UpDown UpDownDataDe 
         Height          =   300
         Left            =   2130
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataDe 
         Height          =   300
         Left            =   1170
         TabIndex        =   2
         Top             =   285
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataAte 
         Height          =   300
         Left            =   4125
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataAte 
         Height          =   300
         Left            =   3165
         TabIndex        =   3
         Top             =   285
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label LabelDataDe 
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
         Left            =   780
         TabIndex        =   38
         Top             =   338
         Width           =   315
      End
      Begin VB.Label LabelDataAte 
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
         Left            =   2745
         TabIndex        =   37
         Top             =   338
         Width           =   360
      End
   End
   Begin VB.Frame FrameCategoria 
      Caption         =   "Categoria"
      Height          =   1740
      Left            =   240
      TabIndex        =   30
      Top             =   4515
      Width           =   5175
      Begin VB.ComboBox CategoriaProduto 
         Height          =   315
         Left            =   1320
         TabIndex        =   11
         Top             =   720
         Width           =   2820
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
         Left            =   435
         TabIndex        =   10
         Top             =   320
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.ComboBox ItemCatDe 
         Enabled         =   0   'False
         Height          =   315
         Left            =   480
         TabIndex        =   12
         Top             =   1200
         Width           =   1900
      End
      Begin VB.ComboBox ItemCatAte 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3120
         TabIndex        =   13
         Top             =   1200
         Width           =   1900
      End
      Begin VB.Label LabelCategoria 
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
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   360
         TabIndex        =   33
         Top             =   765
         Width           =   930
      End
      Begin VB.Label LabelItemCatAte 
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
         Left            =   2730
         TabIndex        =   32
         Top             =   1260
         Width           =   360
      End
      Begin VB.Label LabelItemCatDe 
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
         TabIndex        =   31
         Top             =   1260
         Width           =   315
      End
   End
   Begin VB.Frame FrameProdutos 
      Caption         =   "Produtos"
      Height          =   1290
      Left            =   240
      TabIndex        =   25
      Top             =   3120
      Width           =   5175
      Begin MSMask.MaskEdBox ProdutoDe 
         Height          =   315
         Left            =   510
         TabIndex        =   8
         Top             =   360
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ProdutoAte 
         Height          =   315
         Left            =   510
         TabIndex        =   9
         Top             =   825
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label ProdutoDescricaoAte 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2100
         TabIndex        =   29
         Top             =   825
         Width           =   2970
      End
      Begin VB.Label ProdutoDescricaoDe 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2100
         TabIndex        =   28
         Top             =   360
         Width           =   2970
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
         Left            =   150
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   27
         Top             =   390
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
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   26
         Top             =   870
         Width           =   360
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpProdutosDevTroca.ctx":0A96
      Left            =   1080
      List            =   "RelOpProdutosDevTroca.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   270
      Width           =   2670
   End
   Begin VB.Frame FrameCaixa 
      Caption         =   "Caixa"
      Height          =   1335
      Left            =   240
      TabIndex        =   21
      Top             =   1680
      Width           =   2501
      Begin MSMask.MaskEdBox CaixaDe 
         Height          =   315
         Left            =   645
         TabIndex        =   4
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   19
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CaixaAte 
         Height          =   315
         Left            =   645
         TabIndex        =   5
         Top             =   855
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   19
         PromptChar      =   " "
      End
      Begin VB.Label LabelCaixaDe 
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
         Left            =   240
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   23
         Top             =   390
         Width           =   315
      End
      Begin VB.Label LabelCaixaAte 
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
         Left            =   195
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   22
         Top             =   915
         Width           =   360
      End
   End
   Begin VB.Frame FrameVendedor 
      Caption         =   "Vendedor"
      Height          =   1335
      Left            =   2914
      TabIndex        =   0
      Top             =   1680
      Width           =   2501
      Begin MSMask.MaskEdBox VendedorDe 
         Height          =   315
         Left            =   645
         TabIndex        =   6
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox VendedorAte 
         Height          =   315
         Left            =   645
         TabIndex        =   7
         Top             =   855
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   " "
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
         Height          =   195
         Left            =   240
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   20
         Top             =   390
         Width           =   315
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
         Height          =   195
         Left            =   195
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   19
         Top             =   915
         Width           =   360
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
      Left            =   360
      TabIndex        =   24
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpProdutosDevTroca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'eventos dos browsers
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoCaixa As AdmEvento
Attribute objEventoCaixa.VB_VarHelpID = -1
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1

'variaveis de controle de browser
Dim giProdInicial As Integer
Dim giVendedorInicial As Integer
Dim giCaixaInicial As Integer

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

'type para ser passado como parâmetro
Type TypeProdutosDevTroca
    sVendI As String
    sVendF As String
    sCaixaI As String
    sCaixaF As String
    sProdI As String
    sProdF As String
End Type

Public Sub Form_Load()

Dim lErro As Long
Dim colCategoriaProduto As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto

On Error GoTo Erro_Form_Load

    'instancia os obj
    Set objEventoProduto = New AdmEvento
    Set objEventoVendedor = New AdmEvento
    Set objEventoCaixa = New AdmEvento
           
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoDe)
    If lErro <> SUCESSO Then gError 116000

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoAte)
    If lErro <> SUCESSO Then gError 116001

    'Le as categorias de produto
    lErro = CF("CategoriasProduto_Le_Todas", colCategoriaProduto)
    If lErro <> SUCESSO And lErro <> 22542 Then gError 116173
    
    'Preenche CategoriaProduto
    For Each objCategoriaProduto In colCategoriaProduto

       CategoriaProduto.AddItem objCategoriaProduto.sCategoria

    Next
                  
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 116000, 116001, 116173
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171821)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoExecutar_Click()
'passa as informações p/ o relatorio

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    'preenche as opções de relatório
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 116002

    'Se é para exibir produtos de todas as categorias, ou seja, não existe o filtro por categoria
    If TodasCategorias.Value = vbChecked Then
    
        'Se a filial ativa for difernete de EMPRESA_TODA
        If giFilialEmpresa <> EMPRESA_TODA Then
        
            'Guarda o nome do relatório sem filtro de categoria executado por filial
            gobjRelatorio.sNomeTsk = "PRDVFI"
        
        'Senão
        Else
            
            'Chama o relatório sem filtro de categoria executado para EMPRESA_TODA
            gobjRelatorio.sNomeTsk = "PRDVET"
    
        End If
    
    'Senão, ou seja, se foi selecionada uma faixa de categorias como filtro
    Else
    
        'Se a filial ativa for diferente de EMPRESA_TODA
        If giFilialEmpresa <> EMPRESA_TODA Then
            
            'Guarda o nome do relatório com filtro de categoria executado por filial
            gobjRelatorio.sNomeTsk = "PRDVCTFI"
        
        'Senão
        Else
        
            'Guarda o nome do relatório com filtro de categoria executado para EMPRESA_TODa
            gobjRelatorio.sNomeTsk = "PRDVCTET"
        
        End If
    
    End If
    
    'Prossegue a execução do relatório
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 116002

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171822)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
 
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'limpa as text box
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 116003
    
    'limpa as combos e as labels
    ProdutoDescricaoDe.Caption = ""
    ProdutoDescricaoAte.Caption = ""
        
    TodasCategorias.Value = vbChecked
        
    'posiciona o cursor na combo opções
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 116003
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171823)

    End Select

    Exit Sub
    
End Sub

Private Sub CategoriaProduto_Click()
        
    'desmarca TodasCategorias
    TodasCategorias.Value = vbUnchecked
    Call CategoriaProduto_Validate(bSGECancelDummy)
        
End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub DataDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataDe)

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)
'Valida a data do campo DataDe

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    If Len(DataDe.ClipText) > 0 Then
        
        'valida a data
        lErro = Data_Critica(DataDe.Text)
        If lErro <> SUCESSO Then gError 116004

    End If

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 116004

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171824)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataAte)

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)
'valida a data do campo DataAte

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    If Len(DataAte.ClipText) > 0 Then

        'valida a data
        lErro = Data_Critica(DataAte.Text)
        If lErro <> SUCESSO Then gError 116005

    End If

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 116005

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171825)

    End Select

    Exit Sub

End Sub

Private Sub ItemCatDe_Validate(Cancel As Boolean)
'valida o item selecionado em ItemCatDe

Dim lErro As Long

On Error GoTo Erro_ItemCatDe_Validate

    'verifica se o item existe
    lErro = ItemCategoriaValidate(ItemCatDe)
    If lErro <> SUCESSO Then gError 116061

    Exit Sub

Erro_ItemCatDe_Validate:

    Cancel = True

    Select Case gErr

        Case 116061
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171826)

    End Select

    Exit Sub

End Sub

Private Sub ItemCatAte_Validate(Cancel As Boolean)
'Valida o item selecionado em ItemCatAte

Dim lErro As Long

On Error GoTo Erro_ItemCatAte_Validate

    'verifica se o item existe
    lErro = ItemCategoriaValidate(ItemCatAte)
    If lErro <> SUCESSO Then gError 116050

    Exit Sub

Erro_ItemCatAte_Validate:

    Cancel = True

    Select Case gErr

        Case 116050
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171827)

    End Select
    
    Exit Sub
    
End Sub

Private Function ItemCategoriaValidate(objCampo As Object) As Long
'verifica se o item selecionado existe e carrega as combo ItemCatDe e Ate

Dim lErro As Long
Dim objCategoriaProdutoItem As ClassCategoriaProdutoItem

On Error GoTo Erro_ItemCategoriaValidate

    If Len(Trim(objCampo)) > 0 Then

        'Tenta selecionar na combo
        lErro = Combo_Item_Igual(objCampo)
        If lErro <> SUCESSO Then

            Set objCategoriaProdutoItem = New ClassCategoriaProdutoItem

            'Preenche o objeto com a Categoria
            objCategoriaProdutoItem.sCategoria = CategoriaProduto.Text
            
            'preenche o obj c/ o parametro (ItemCatDe ou ItemCatAte)
            objCategoriaProdutoItem.sItem = objCampo.Text

            'Lê Categoria De Produto no BD
            lErro = CF("CategoriaProduto_Le_Item", objCategoriaProdutoItem)
            If lErro <> SUCESSO And lErro <> 22603 Then gError 116062
                                    
            'Item da Categoria não está cadastrado
            If lErro = 22603 Then gError 116063
        
        End If

    End If
    
    ItemCategoriaValidate = SUCESSO
    
    Exit Function

Erro_ItemCategoriaValidate:

    ItemCategoriaValidate = gErr
    
    Select Case gErr

        Case 116062

        Case 116063
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_INEXISTENTE", gErr, objCategoriaProdutoItem.sItem, objCategoriaProdutoItem.sCategoria)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171828)

    End Select

    Exit Function

End Function

Private Sub objEventoCaixa_evSelecao(obj1 As Object)
'evento de inclusão de um item selecionado no browser Caixa

Dim objCaixa As ClassCaixa

On Error GoTo Erro_objEventoCaixa_evSelecao

    Set objCaixa = obj1
    
    'Preenche campo Caixa
    If giCaixaInicial = 1 Then
        CaixaDe.Text = CStr(objCaixa.iCodigo)
        CaixaDe_Validate (bSGECancelDummy)
    Else
        CaixaAte.Text = CStr(objCaixa.iCodigo)
        CaixaAte_Validate (bSGECancelDummy)
    End If

    Me.Show

    Exit Sub

Erro_objEventoCaixa_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 171829)

    End Select
    
    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)
'evento de inclusão de um item selecionado no browser Produto

Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1
    
    'Preenche campo Produto
    If giProdInicial = 1 Then
        ProdutoDe.PromptInclude = False
        ProdutoDe.Text = CStr(objProduto.sCodigo)
        ProdutoDe.PromptInclude = True
        ProdutoDe_Validate (bSGECancelDummy)
    Else
        ProdutoAte.PromptInclude = False
        ProdutoAte.Text = CStr(objProduto.sCodigo)
        ProdutoAte.PromptInclude = True
        ProdutoAte_Validate (bSGECancelDummy)
    End If

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 171830)

    End Select
    
    Exit Sub

End Sub

Private Sub VendedorDe_Validate(Cancel As Boolean)
'valida o codigo/ nome reduzido do vendedor

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_VendedorDe_Validate

    'se estiver preenchido
    If Len(Trim(VendedorDe.Text)) > 0 Then
        
        'Tenta ler o vendedor (Código ou nome)
        lErro = TP_Vendedor_Le2(VendedorDe, objVendedor, 0)
        If lErro <> SUCESSO Then gError 116006

    End If
    
    giVendedorInicial = 1
    
    Exit Sub

Erro_VendedorDe_Validate:

    Cancel = True
    
    Select Case gErr

        Case 116006

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171831)

    End Select

End Sub

Private Sub VendedorAte_Validate(Cancel As Boolean)
'valida o codigo/ nome reduzido do vendedor

Dim lErro As Long
Dim objVendedor As ClassVendedor

On Error GoTo Erro_VendedorAte_Validate

    'se estiver preenchido
    If Len(Trim(VendedorAte.Text)) > 0 Then

        Set objVendedor = New ClassVendedor

        'Tenta ler o vendedor (Código ou nome)
        lErro = TP_Vendedor_Le2(VendedorAte, objVendedor, 0)
        If lErro <> SUCESSO Then gError 116007

    End If
    
    giVendedorInicial = 0
 
    Exit Sub

Erro_VendedorAte_Validate:

    Cancel = True
    
    Select Case gErr

        Case 116007
                      
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171832)

    End Select

End Sub

Private Sub ProdutoAte_Validate(Cancel As Boolean)
'valida o codigo do produto

Dim lErro As Long

On Error GoTo Erro_ProdutoAte_Validate

    giProdInicial = 0

    'faz o tratamento do produto
    lErro = CF("Produto_Perde_Foco", ProdutoAte, ProdutoDescricaoAte)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 116008
    
    'Se não encontrou o produto => erro
    If lErro = 27095 Then gError 116009

    Exit Sub

Erro_ProdutoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 116008

        Case 116009
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171833)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoDe_Validate(Cancel As Boolean)
'valida o codigo do produto

Dim lErro As Long

On Error GoTo Erro_ProdutoDe_Validate

    giProdInicial = 1

    'faz o tratamento do produto
    lErro = CF("Produto_Perde_Foco", ProdutoDe, ProdutoDescricaoDe)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 116010
    
    'Se não encontrou o produto => erro
    If lErro = 27095 Then gError 116011

    Exit Sub

Erro_ProdutoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 116010

        Case 116011
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171834)

    End Select

    Exit Sub

End Sub

Private Sub CaixaAte_Validate(Cancel As Boolean)
'valida o codigo/ nome reduzido do caixa

Dim lErro As Long
Dim objCaixa As New ClassCaixa

On Error GoTo Erro_CaixaAte_Validate

    giCaixaInicial = 0

    If Len(Trim(CaixaAte.Text)) > 0 Then

        'preenche o obj c/ o código e a filial
        objCaixa.iCodigo = Codigo_Extrai(CaixaAte.Text)
        objCaixa.iFilialEmpresa = giFilialEmpresa
        
        'Tenta ler a Caixa (Código ou nome)
        lErro = CF("TP_Caixa_Le1", CaixaAte, objCaixa)
        If lErro <> SUCESSO And lErro <> 116175 And lErro <> 116177 Then gError 116163
        
        'codigo inexistente
        If lErro = 116175 Then gError 116169

        'nome_reduzido inexistente
        If lErro = 116177 Then gError 116179

    End If
 
    Exit Sub

Erro_CaixaAte_Validate:

    Cancel = True
    
    Select Case gErr

        Case 116163
            
        Case 116169
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXA_INEXISTENTE", gErr, objCaixa.iCodigo)
            
        Case 116179
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXA_NOMERED_INEXISTENTE", gErr, objCaixa.sNomeReduzido)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171835)

    End Select

    Exit Sub

End Sub

Private Sub CaixaDe_Validate(Cancel As Boolean)
'valida o codigo/ nome reduzido do caixa

Dim lErro As Long
Dim objCaixa As New ClassCaixa
    
On Error GoTo Erro_CaixaDe_Validate

    giCaixaInicial = 1

    If Len(Trim(CaixaDe.Text)) > 0 Then
        
        'preenche o obj c/ o código e a filial
        objCaixa.iCodigo = Codigo_Extrai(CaixaDe.Text)
        objCaixa.iFilialEmpresa = giFilialEmpresa
        
        'Tenta ler Caixa (Código ou nome_reduzido)
        lErro = CF("TP_Caixa_Le1", CaixaDe, objCaixa)
        If lErro <> SUCESSO And lErro <> 116175 And lErro <> 116177 Then gError 116162

        'código inexistente
        If lErro = 116175 Then gError 116168
        
        'nomereduzido inexistente
        If lErro = 116177 Then gError 116178

    End If
    
    Exit Sub

Erro_CaixaDe_Validate:

    Cancel = True
    
    Select Case gErr

        Case 116162

        Case 116168
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXA_INEXISTENTE", gErr, objCaixa.iCodigo)
            
        Case 116178
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXA_NOMERED_INEXISTENTE", gErr, objCaixa.sNomeReduzido)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171836)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaProduto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As ClassCategoriaProdutoItem
Dim colCategoria As New Collection

On Error GoTo Erro_CategoriaProduto_Validate

    If Len(Trim(CategoriaProduto.Text)) > 0 Then

        'Preenche o objeto com a Categoria
         objCategoriaProduto.sCategoria = CategoriaProduto.Text

        'Lê os dados de itens de categorias de produto
        lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colCategoria)
        If lErro <> SUCESSO And lErro <> 22541 Then gError 116012
        
        'se não encontrou
        If lErro = 22541 Then gError 116284
        
        'desmarca todasCategorias
        TodasCategorias.Value = vbUnchecked
        
        'Limpa Valor Inicial e Final
        ItemCatDe.Clear
        ItemCatAte.Clear
        ItemCatDe.Enabled = True
        ItemCatAte.Enabled = True
            
        'Preenche a Combo ValorInicial e Final
        For Each objCategoriaProdutoItem In colCategoria

            ItemCatDe.AddItem (objCategoriaProdutoItem.sItem)
            
            ItemCatAte.AddItem (objCategoriaProdutoItem.sItem)

        Next
                
    Else
    
        'se a Categoria não estiver preenchida ----> limpa e desabilita os Valores Inicial e Final
        ItemCatDe.Clear
        ItemCatAte.Clear
        ItemCatDe.Enabled = False
        ItemCatAte.Enabled = False
    
    End If

    Exit Sub

Erro_CategoriaProduto_Validate:

    Cancel = True

    Select Case gErr

        Case 116012

        Case 116284
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_NAO_CADASTRADA", gErr, objCategoriaProduto.sCategoria)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171837)

    End Select

    Exit Sub

End Sub

Private Function Formata_E_Critica_Parametros(tProdutosDevTroca As TypeProdutosDevTroca) As Long
'Formata os produtos retornando em TypeProdutosDevTroca
'Verifica se os parâmetros iniciais são maiores que os finais

Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim lErro As Long
 
On Error GoTo Erro_Formata_E_Critica_Parametros

    'data inicial não pode ser maior que a data final
    If Trim(DataDe.ClipText) <> "" And Trim(DataAte.ClipText) <> "" Then
    
         If StrParaDate(DataDe.Text) > StrParaDate(DataAte.Text) Then gError 116018
    
    End If

    With tProdutosDevTroca
        
        'critica Caixa Inicial e Final
        If Trim(CaixaDe.Text) <> "" Then
            .sCaixaI = Codigo_Extrai(CaixaDe.Text)
        Else
            .sCaixaI = ""
        End If
        
        If Trim(CaixaAte.Text) <> "" Then
            .sCaixaF = Codigo_Extrai(CaixaAte.Text)
        Else
            .sCaixaF = ""
        End If
          
        'caixa inicial não pode ser maior do que a final
        If .sCaixaI <> "" And .sCaixaF <> "" Then
            
            If StrParaInt(.sCaixaI) > StrParaInt(.sCaixaF) Then gError 116017
            
        End If
        
        'Vendedor inicial não pode ser maior que o Vendedor final
        If Trim(VendedorDe.Text) <> "" Then
            .sVendI = Codigo_Extrai(VendedorDe.Text)
        Else
            .sVendI = ""
        End If
        
        If Trim(VendedorAte.Text) <> "" Then
            .sVendF = Codigo_Extrai(VendedorAte.Text)
        Else
            .sVendF = ""
        End If
             
        If .sVendI <> "" And .sVendF <> "" Then
            If StrParaInt(.sVendI) > StrParaInt(.sVendF) Then gError 116016
        End If
        
        'formata o Produto Inicial
        lErro = CF("Produto_Formata", ProdutoDe.Text, .sProdI, iProdPreenchido_I)
        If lErro <> SUCESSO Then gError 116013
    
        If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then .sProdI = ""
    
        'formata o Produto Final
        lErro = CF("Produto_Formata", ProdutoAte.Text, .sProdF, iProdPreenchido_F)
        If lErro <> SUCESSO Then gError 116014
    
        If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then .sProdF = ""
    
        'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
        If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then
    
            If .sProdI > .sProdF Then gError 116015
    
        End If
                 
    End With
    
    If CategoriaProduto.Text <> "" Then
    
        'categoria ate não pode ser maior do que categoria de
        If ItemCatDe.Text <> "" And ItemCatAte.Text <> "" Then
           
            If ItemCatDe.Text > ItemCatAte.Text Then gError 116060
            
        End If
           
    End If
    
    'se TodasCategorias estiver desmarcada, tem que haver categoria selecionada
    If Trim(ItemCatDe.Text) = "" And Trim(ItemCatAte.Text) = "" And TodasCategorias.Value = vbUnchecked Then gError 116068
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 116013
            ProdutoDe.SetFocus

        Case 116014
            ProdutoAte.SetFocus

        Case 116015
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            ProdutoDe.SetFocus
            
        Case 116016
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_INICIAL_MAIOR", gErr)
            VendedorDe.SetFocus
            
        Case 116017
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXA_INICIAL_MAIOR", gErr)
            CaixaDe.SetFocus
        
         Case 116018
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataDe.SetFocus
            
        Case 116060
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_INICIAL_MAIOR", gErr)
            ItemCatDe.SetFocus
            
        Case 116068
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_NAO_INFORMADO", gErr)
            CategoriaProduto.SetFocus
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171838)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, tProdutosDevTroca As TypeProdutosDevTroca)
'monta a expressão de seleção de relatório

Dim lErro As Long
Dim sExpressao As String

On Error GoTo Erro_Monta_Expressao_Selecao

    With tProdutosDevTroca
    
        'monta expressão de produto
        If .sProdI <> "" Then sExpressao = "Produto >= " & Forprint_ConvTexto(.sProdI)
    
        If .sProdF <> "" Then
    
            If sExpressao <> "" Then sExpressao = sExpressao & " E "
            sExpressao = sExpressao & "Produto <= " & Forprint_ConvTexto(.sProdF)
    
        End If
        
        'monta expressão da caixa
        If .sCaixaI <> "" Then
            
            If sExpressao <> "" Then sExpressao = sExpressao & " E "
            sExpressao = sExpressao & "Caixa >= " & Forprint_ConvInt(StrParaInt(.sCaixaI))
            
        End If
    
        If .sCaixaF <> "" Then
    
            If sExpressao <> "" Then sExpressao = sExpressao & " E "
            sExpressao = sExpressao & "Caixa <= " & Forprint_ConvInt(StrParaInt(.sCaixaF))
            
        End If
        
        'monta a expressão do vendedor
        If .sVendI <> "" Then
            
            If sExpressao <> "" Then sExpressao = sExpressao & " E "
            sExpressao = sExpressao & "Vendedor >= " & Forprint_ConvInt(StrParaInt(.sVendI))
            
        End If
    
        If .sVendF <> "" Then
    
            If sExpressao <> "" Then sExpressao = sExpressao & " E "
            sExpressao = sExpressao & "Vendedor <= " & Forprint_ConvInt(StrParaInt(.sVendF))
    
        End If
        
    End With
    
    'monta a expressão da data
    If Trim(DataDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(StrParaDate(DataDe.Text))

    End If
    
    If Trim(DataAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(StrParaDate(DataAte.Text))

    End If
    
    'monta a expressão das categorias
    If TodasCategorias.Value = vbUnchecked Then
           
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CategoriaProduto = " & Forprint_ConvTexto(CategoriaProduto.Text)
            
        If Trim(ItemCatDe.Text) <> "" Then

            If sExpressao <> "" Then sExpressao = sExpressao & " E "
            sExpressao = sExpressao & "ItemCategoriaProduto >= " & Forprint_ConvTexto(ItemCatDe.Text)

        End If
        
        If Trim(ItemCatAte.Text) <> "" Then

            If sExpressao <> "" Then sExpressao = sExpressao & " E "
            sExpressao = sExpressao & "ItemCategoriaProduto <= " & Forprint_ConvTexto(ItemCatAte.Text)

        End If
        
    End If
    
    'passa a expressão completa para o obj
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171839)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim tProdutosDevTroca As TypeProdutosDevTroca

On Error GoTo Erro_PreencherRelOp
       
    lErro = Formata_E_Critica_Parametros(tProdutosDevTroca)
    If lErro <> SUCESSO Then gError 116019

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 116020
         
    'inclui os parametros produto, vendedor, caixa
    With tProdutosDevTroca
         
        lErro = objRelOpcoes.IncluirParametro("TPRODINIC", .sProdI)
        If lErro <> AD_BOOL_TRUE Then gError 116021
    
        lErro = objRelOpcoes.IncluirParametro("TPRODFIM", .sProdF)
        If lErro <> AD_BOOL_TRUE Then gError 116022
            
        lErro = objRelOpcoes.IncluirParametro("NVENDI", .sVendI)
        If lErro <> AD_BOOL_TRUE Then gError 166064
            
        lErro = objRelOpcoes.IncluirParametro("NVENDF", .sVendF)
        If lErro <> AD_BOOL_TRUE Then gError 166065
            
        lErro = objRelOpcoes.IncluirParametro("NCAIXAI", .sCaixaI)
        If lErro <> AD_BOOL_TRUE Then gError 116023
        
        lErro = objRelOpcoes.IncluirParametro("NCAIXAF", .sCaixaF)
        If lErro <> AD_BOOL_TRUE Then gError 116024
            
    End With
    
    'inclui os parametros vendedor, caixa (controle)
    lErro = objRelOpcoes.IncluirParametro("TCAIXAI", CaixaDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 116180

    lErro = objRelOpcoes.IncluirParametro("TCAIXAF", CaixaAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 116181

    lErro = objRelOpcoes.IncluirParametro("TVENDI", VendedorDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 116182

    lErro = objRelOpcoes.IncluirParametro("TVENDF", VendedorAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 116183

    'inclui o parametro Data
    If DataDe.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINIC", DataDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 116025

    If DataAte.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 116026
   
    'inclui o parametro todasCategorias
    lErro = objRelOpcoes.IncluirParametro("NTODASCAT", CStr(TodasCategorias.Value))
    If lErro <> AD_BOOL_TRUE Then gError 116027
    
    'inclui o parametro CategoriaProduto
    lErro = objRelOpcoes.IncluirParametro("TCATPROD", CategoriaProduto.Text)
    If lErro <> AD_BOOL_TRUE Then gError 116028
    
    lErro = objRelOpcoes.IncluirParametro("TITEMCATPRODINI", ItemCatDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 116029
    
    lErro = objRelOpcoes.IncluirParametro("TITEMCATPRODFIM", ItemCatAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 116030
       
    'Monta a expressão a ser passada ao gerador de relatórios
    lErro = Monta_Expressao_Selecao(objRelOpcoes, tProdutosDevTroca)
    If lErro <> SUCESSO Then gError 116031

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 116019 To 116031, 166064, 166065, 116180, 116181, 116182, 116183
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171840)

    End Select

    Exit Function

End Function

Private Sub objEventoVendedor_evSelecao(obj1 As Object)
'evento de inclusão de item selecionado no browser Vendedor

Dim objVendedor As ClassVendedor

On Error GoTo Erro_objEventoVendedor_evSelecao

    Set objVendedor = obj1
    
    'Preenche campo Vendedor
    If giVendedorInicial = 1 Then
        VendedorDe.Text = CStr(objVendedor.iCodigo)
        VendedorDe_Validate (bSGECancelDummy)
    Else
        VendedorAte.Text = CStr(objVendedor.iCodigo)
        VendedorAte_Validate (bSGECancelDummy)
    End If

    Me.Show

    Exit Sub

Erro_objEventoVendedor_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171841) '???Luiz: falta o parâmetro Error ok

    End Select
    
    Exit Sub

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long
'carrega o combo de opções

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 116172
    
    Set gobjRelOpcoes = objRelOpcoes
    Set gobjRelatorio = objRelatorio
    
    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 116032
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 116032
        
        Case 116172
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171842)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro Then gError 116033

   'pega parâmetro Caixa Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCAIXAI", sParam)
    If lErro <> SUCESSO Then gError 116034
    
    CaixaDe.Text = sParam
    Call CaixaDe_Validate(bSGECancelDummy)
    
    'pega parâmetro Caixa Final e exibe
    lErro = objRelOpcoes.ObterParametro("NCAIXAF", sParam)
    If lErro <> SUCESSO Then gError 116035
    
    CaixaAte.Text = sParam
    Call CaixaAte_Validate(bSGECancelDummy)
    
    'pega o parametro Vendedor inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NVENDI", sParam)
    If lErro <> SUCESSO Then gError 116066
    
    VendedorDe.Text = sParam
    Call VendedorDe_Validate(bSGECancelDummy)
    
    'pega o parametro Vendedor final e exibe
    lErro = objRelOpcoes.ObterParametro("NVENDF", sParam)
    If lErro <> SUCESSO Then gError 116067
    
    VendedorAte.Text = sParam
    Call VendedorAte_Validate(bSGECancelDummy)
    
    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro <> SUCESSO Then gError 116036

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoDe, ProdutoDescricaoDe)
    If lErro <> SUCESSO Then gError 116037

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 116038

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoAte, ProdutoDescricaoAte)
    If lErro <> SUCESSO Then gError 116039
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then gError 116040

    Call DateParaMasked(DataDe, StrParaDate(sParam))

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 116041

    Call DateParaMasked(DataAte, StrParaDate(sParam))
    
    'pega parâmetro TodasCategorias e exibe
    lErro = objRelOpcoes.ObterParametro("NTODASCAT", sParam)
    If lErro <> SUCESSO Then gError 116042

    TodasCategorias.Value = StrParaInt(sParam)

    'pega parâmetro categoria de produto e exibe
    lErro = objRelOpcoes.ObterParametro("TCATPROD", sParam)
    If lErro <> SUCESSO Then gError 116043
        
    If sParam <> "" Then
    
        CategoriaProduto.Text = sParam
        Call CategoriaProduto_Validate(bSGECancelDummy)
    
        'pega parâmetro valor inicial e exibe
        lErro = objRelOpcoes.ObterParametro("TITEMCATPRODINI", sParam)
        If lErro <> SUCESSO Then gError 116044
        
        ItemCatDe.Text = sParam
        
        'pega parâmetro Valor Final e exibe
        lErro = objRelOpcoes.ObterParametro("TITEMCATPRODFIM", sParam)
        If lErro <> SUCESSO Then gError 116045
    
        ItemCatAte.Text = sParam
    Else
    
        TodasCategorias.Value = vbChecked
    
    End If
         
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 116033 To 116045, 116066, 116067

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171843)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If Trim(ComboOpcoes.Text) = "" Then gError 116059

    'preenche o relatorio c/ as opções da tela
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 116046

    'carrega o obj com a opção da tela
    gobjRelOpcoes.sNome = ComboOpcoes.Text

    'grava a opção
    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 116047

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 116048
    
    'limpa a tela
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 116059
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 116046, 116047, 116048

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171844)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()
'exclui a opção de relatorio selecionada

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 116052

    'pergunta se deseja excluir
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPPRODUTO", ComboOpcoes.Text)

    'se a resposta for sim
    If vbMsgRes = vbYes Then

        'exclui a opção do gobjRelOpcoes
        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 116051

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa a tela
        Call BotaoLimpar_Click
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 116052
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 116051

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171845)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_DownClick()
'diminui a data

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 116053

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 116053
            DataDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171846)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()
'aumenta a data

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 116054

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 116054
            DataDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171847)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_DownClick()
'Dimunui a data

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 116055

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 116055
            DataAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171848)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()
'aumenta a data

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 116056

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 116056
            DataAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171849)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)
'libera os objs

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoProduto = Nothing
    Set objEventoVendedor = Nothing
    Set objEventoCaixa = Nothing
    
End Sub

Private Sub LabelCaixaDe_Click()
'sub chamadora do browser Caixa

Dim objCaixa As New ClassCaixa
Dim colSelecao As Collection

On Error GoTo Erro_LabelCaixaDe_Click

    giCaixaInicial = 1
    
    'se estiver preenchida
    If Len(Trim(CaixaDe.Text)) > 0 Then
        'Preenche com a caixa da tela
        objCaixa.iCodigo = Codigo_Extrai(CaixaDe.Text)
    End If
    
    'Chama Tela de caixa
    Call Chama_Tela("CaixaLista", colSelecao, objCaixa, objEventoCaixa)
    
    Exit Sub

Erro_LabelCaixaDe_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171850)

    End Select
    
    Exit Sub
    
End Sub

Private Sub LabelCaixaAte_Click()
'sub chamadora do browser Caixa

Dim objCaixa As New ClassCaixa
Dim colSelecao As Collection

On Error GoTo Erro_LabelCaixaAte_Click

    giCaixaInicial = 0
    
    'se estiver preenchida
    If Len(Trim(CaixaAte.Text)) > 0 Then
        'Preenche com a caixa da tela
        objCaixa.iCodigo = Codigo_Extrai(CaixaAte.Text)
    End If
    
    'Chama Tela Caixa
    Call Chama_Tela("CaixaLista", colSelecao, objCaixa, objEventoCaixa)

    Exit Sub

Erro_LabelCaixaAte_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171851)

    End Select
    
    Exit Sub
    
End Sub

Private Sub LabelVendedorDe_Click()
'sub chamadora do browser Vendedo

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection

On Error GoTo Erro_LabelVendedorDe_Click

    giVendedorInicial = 1
    
    'se estiver preenchido
    If Len(Trim(VendedorDe.Text)) > 0 Then
        'Preenche o obj com o cód do Vendedor da tela
        objVendedor.iCodigo = Codigo_Extrai(VendedorDe.Text)
    End If
    
    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)
    
    Exit Sub

Erro_LabelVendedorDe_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171852)

    End Select
    
    Exit Sub
    
End Sub

Private Sub LabelVendedorAte_Click()
'sub chamadora do browser Vendedor

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection

On Error GoTo Erro_LabelVendedorAte_Click

    giVendedorInicial = 0
    
    'se estiver preenchido
    If Len(Trim(VendedorAte.Text)) > 0 Then
        'Preenche o obj com o Vendedor da tela
        objVendedor.iCodigo = Codigo_Extrai(VendedorAte.Text)
    End If
    
    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

    Exit Sub

Erro_LabelVendedorAte_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171853)

    End Select
    
    Exit Sub
    
End Sub

Private Sub LabelProdutoDe_Click()
'sub chamadora do browser Produto

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProdutoDe_Click

    giProdInicial = 1

    'Verifica se o produto foi preenchido
    If Len(Trim(ProdutoDe.ClipText)) <> 0 Then

        'formata o produto
        lErro = CF("Produto_Formata", ProdutoDe.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 116057

        'Preenche o código de objProduto
        objProduto.sCodigo = sProdutoFormatado

    End If

    'chama a tela de produtos
    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 116057

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171854)

    End Select

    Exit Sub

End Sub

Private Sub LabelProdutoAte_Click()
'sub chamadora do browser Produtos

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProdutoAte_Click

    giProdInicial = 0

    'Verifica se o produto foi preenchido
    If Len(Trim(ProdutoAte.ClipText)) <> 0 Then

        'formata o produto
        lErro = CF("Produto_Formata", ProdutoAte.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 116058

        'Preenche o código de objProduto
        objProduto.sCodigo = sProdutoFormatado

    End If

    'chama a tela de produtos
    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 116058

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171855)

    End Select

    Exit Sub

End Sub

Private Sub TodasCategorias_Click()
'checkbox todascategorias

    'se estiver marcado
    If TodasCategorias.Value = vbChecked Then
    
        'desebilita e limpa as opções de categoria de produto
        ItemCatDe.Enabled = False
        ItemCatAte.Enabled = False
        ItemCatAte.Clear
        ItemCatDe.Clear
        CategoriaProduto.Text = ""
        
    'senão
    Else
        
        'habilita as opções
        ItemCatDe.Enabled = True
        ItemCatAte.Enabled = True
    
    End If

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is CaixaDe Then
            Call LabelCaixaDe_Click
        ElseIf Me.ActiveControl Is CaixaAte Then
            Call LabelCaixaAte_Click
        ElseIf Me.ActiveControl Is VendedorDe Then
            Call LabelVendedorDe_Click
        ElseIf Me.ActiveControl Is VendedorAte Then
            Call LabelVendedorAte_Click
        ElseIf Me.ActiveControl Is ProdutoDe Then
            Call LabelProdutoDe_Click
        ElseIf Me.ActiveControl Is ProdutoAte Then
            Call LabelProdutoAte_Click
        End If
    
    End If

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

'    Parent.HelpContextID = IDH_RELOP_PEDIDOS_NAO_ENTREGUES
    Set Form_Load_Ocx = Me
    Caption = "Produtos Devolvidos em Trocas"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpProdutosDevTroca"
    
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

Private Sub LabelVendedorDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelVendedorDe, Source, X, Y)
End Sub

Private Sub LabelVendedorDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelVendedorDe, Button, Shift, X, Y)
End Sub

Private Sub LabelVendedorAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelVendedorAte, Source, X, Y)
End Sub

Private Sub LabelVendedorAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelVendedorAte, Button, Shift, X, Y)
End Sub

Private Sub labelDataAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataAte, Source, X, Y)
End Sub

Private Sub LabelDataAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataAte, Button, Shift, X, Y)
End Sub

Private Sub LabelDataDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataDe, Source, X, Y)
End Sub

Private Sub LabelDataDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataDe, Button, Shift, X, Y)
End Sub

Private Sub LabelItemCatDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelItemCatDe, Source, X, Y)
End Sub

Private Sub LabelItemCatDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelItemCatDe, Button, Shift, X, Y)
End Sub

Private Sub LabelItemCatAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelItemCatAte, Source, X, Y)
End Sub

Private Sub LabelItemCatAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelItemCatAte, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LabelCaixaDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCaixaDe, Source, X, Y)
End Sub

Private Sub LabelCaixaDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCaixaDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCaixaAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCaixaAte, Source, X, Y)
End Sub

Private Sub LabelCaixaAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCaixaAte, Button, Shift, X, Y)
End Sub
