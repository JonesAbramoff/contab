VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpKardexDia_LOcx 
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8790
   KeyPreview      =   -1  'True
   ScaleHeight     =   5640
   ScaleWidth      =   8790
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   4185
      Index           =   1
      Left            =   165
      TabIndex        =   15
      Top             =   1200
      Width           =   8445
      Begin VB.ListBox Almoxarifados 
         Height          =   3570
         ItemData        =   "RelOpKardexDia_LOcx.ctx":0000
         Left            =   5940
         List            =   "RelOpKardexDia_LOcx.ctx":0002
         TabIndex        =   19
         Top             =   495
         Width           =   2385
      End
      Begin VB.Frame Frame1 
         Caption         =   "Produtos"
         Height          =   1332
         Left            =   120
         TabIndex        =   16
         Top             =   855
         Width           =   5655
         Begin MSMask.MaskEdBox ProdutoFinal 
            Height          =   315
            Left            =   750
            TabIndex        =   17
            Top             =   870
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
            TabIndex        =   18
            Top             =   360
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
            TabIndex        =   30
            Top             =   360
            Width           =   3135
         End
         Begin VB.Label DescProdFim 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2295
            TabIndex        =   31
            Top             =   870
            Width           =   3135
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
            Height          =   255
            Left            =   345
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   32
            Top             =   375
            Width           =   495
         End
         Begin VB.Label LabelProdutoAte 
            Caption         =   "At�:"
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
            Left            =   285
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   33
            Top             =   885
            Width           =   555
         End
      End
      Begin MSMask.MaskEdBox Almoxarifado 
         Height          =   315
         Left            =   1410
         TabIndex        =   20
         Top             =   315
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label Label9 
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
         Left            =   150
         TabIndex        =   34
         Top             =   345
         Width           =   1185
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
         Left            =   5880
         TabIndex        =   35
         Top             =   240
         Width           =   1185
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   4185
      Index           =   2
      Left            =   255
      TabIndex        =   0
      Top             =   1170
      Visible         =   0   'False
      Width           =   8250
      Begin VB.Frame Frame6 
         Caption         =   "Livro/P�ginas"
         Height          =   900
         Left            =   195
         TabIndex        =   11
         Top             =   1320
         Width           =   7995
         Begin VB.TextBox NumLivro 
            Height          =   315
            Left            =   1800
            TabIndex        =   14
            Top             =   375
            Width           =   585
         End
         Begin VB.TextBox PaginaInicial 
            Height          =   315
            Left            =   4080
            TabIndex        =   13
            Top             =   360
            Width           =   540
         End
         Begin VB.TextBox QtdePaginas 
            Height          =   315
            Left            =   7275
            TabIndex        =   12
            Top             =   360
            Width           =   570
         End
         Begin VB.Label LabelNumLivro 
            Caption         =   "N�mero do Livro:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   225
            TabIndex        =   36
            Top             =   405
            Width           =   1470
         End
         Begin VB.Label labelPagInic 
            Caption         =   "P�gina Inicial:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2820
            TabIndex        =   37
            Top             =   405
            Width           =   1245
         End
         Begin VB.Label LabelQtdePag 
            Caption         =   "Quantidade de  p�ginas:"
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
            Left            =   5145
            TabIndex        =   38
            Top             =   390
            Width           =   2145
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
         Height          =   435
         Left            =   210
         TabIndex        =   10
         Top             =   3525
         Width           =   3600
      End
      Begin VB.Frame FrameData 
         Caption         =   "Data"
         Height          =   900
         Left            =   195
         TabIndex        =   5
         Top             =   225
         Width           =   7995
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   315
            Left            =   2415
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   375
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataInicial 
            Height          =   300
            Left            =   1410
            TabIndex        =   7
            Top             =   390
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
            Left            =   6345
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   375
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataFinal 
            Height          =   300
            Left            =   5340
            TabIndex        =   9
            Top             =   390
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
            Caption         =   "At�:"
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
            Left            =   4935
            TabIndex        =   39
            Top             =   450
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
            Height          =   240
            Left            =   1020
            TabIndex        =   40
            Top             =   420
            Width           =   345
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Imprime"
         Height          =   900
         Left            =   195
         TabIndex        =   1
         Top             =   2430
         Width           =   7995
         Begin VB.OptionButton OpcaoLivro 
            Caption         =   "Somente o Livro"
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
            Left            =   525
            TabIndex        =   4
            Top             =   390
            Value           =   -1  'True
            Width           =   1995
         End
         Begin VB.OptionButton OpcaoTermos 
            Caption         =   "Somente os Termos"
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
            Left            =   3142
            TabIndex        =   3
            Top             =   390
            Width           =   2010
         End
         Begin VB.OptionButton OpcaoAmbos 
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
            Height          =   225
            Left            =   5760
            TabIndex        =   2
            Top             =   390
            Width           =   1905
         End
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpKardexDia_LOcx.ctx":0004
      Left            =   1395
      List            =   "RelOpKardexDia_LOcx.ctx":0006
      Sorted          =   -1  'True
      TabIndex        =   27
      Top             =   315
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
      Left            =   4575
      Picture         =   "RelOpKardexDia_LOcx.ctx":0008
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   240
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6480
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpKardexDia_LOcx.ctx":010A
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpKardexDia_LOcx.ctx":0288
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpKardexDia_LOcx.ctx":07BA
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpKardexDia_LOcx.ctx":0944
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4635
      Left            =   120
      TabIndex        =   28
      Top             =   810
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   8176
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
      Caption         =   "Op��o:"
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
      Left            =   690
      TabIndex        =   29
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "RelOpKardexDia_LOcx"
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
Dim iFrameAtual As Integer

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iFrameAtual = 1

    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento

    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd",ProdutoInicial)
    If lErro <> SUCESSO Then Error 64683

    lErro = CF("Inicializa_Mascara_Produto_MaskEd",ProdutoFinal)
    If lErro <> SUCESSO Then Error 64684

    'carrega a ListBox Almoxarifados
    lErro = Carrega_Lista_Almoxarifado()
    If lErro <> SUCESSO Then Error 64686
    
    Call Define_Padrao

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = Err

    Select Case Err

        Case 64683, 64684, 64686

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169499)

    End Select

    Exit Sub

End Sub

Sub Define_Padrao()

Dim lErro As Long

On Error GoTo Erro_Define_Padrao

    giProdInicial = 1
    
    CheckProdSemMovimento.Value = 0
    
    Exit Sub

Erro_Define_Padrao:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169500)

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

        'Preenche o c�digo de objProduto
        lErro = CF("Produto_Formata",ProdutoFinal.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 82363

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 82363

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169501)

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

        'Preenche o c�digo de objProduto
        lErro = CF("Produto_Formata",ProdutoInicial.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 82362

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 82362

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169502)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'L� o Produto
    lErro = CF("Produto_Le",objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82366

    'Se n�o achou o Produto --> erro
    If lErro = 28030 Then gError 82367

    lErro = CF("Traz_Produto_MaskEd",objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 82368

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 82366, 82368

        Case 82367
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169503)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoDe_evSelecao

    Set objProduto = obj1

    'L� o Produto
    lErro = CF("Produto_Le",objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82372

    'Se n�o achou o Produto --> erro
    If lErro = 28030 Then gError 82373

    lErro = CF("Traz_Produto_MaskEd",objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 82374

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 82372, 82374

        Case 82373
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169504)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169505)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169506)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'l� os par�metros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    Call Limpar_Tela

    lErro = objRelOpcoes.Carregar
    If lErro Then Error 64689

    'pega par�metro Almoxarifado e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("NALMOX", sParam)
    If lErro Then Error 64690
    
    Almoxarifado.Text = sParam
    Call Almoxarifado_Validate(bSGECancelDummy)
   
    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro Then Error 64691

    lErro = CF("Traz_Produto_MaskEd",sParam, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then Error 64692

    'pega par�metro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro Then Error 64693

    lErro = CF("Traz_Produto_MaskEd",sParam, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then Error 64694
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then Error 64695

    Call DateParaMasked(DataInicial, CDate(sParam))
 
    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then Error 64696

    Call DateParaMasked(DataFinal, CDate(sParam))
    
    'pega par�metro p/ listar produto sem movimento e exibe
    lErro = objRelOpcoes.ObterParametro("NPRODSEMMOV", sParam)
    If lErro Then Error 64697

    CheckProdSemMovimento.Value = CInt(sParam)
        
    'pega par�metro numero do livro e exibe
    lErro = objRelOpcoes.ObterParametro("NNUMLIVRO", sParam)
    If lErro Then Error 64698
    
    NumLivro.Text = sParam
    
    'pega par�metro de p�gina inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NPAGINIC", sParam)
    If lErro Then Error 64699
    
    PaginaInicial.Text = sParam
    
    'pega par�metro quantidade de p�ginas e exibe
    lErro = objRelOpcoes.ObterParametro("NQTDEPAG", sParam)
    If lErro Then Error 64700
    
    QtdePaginas.Text = sParam
   
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 64689 To 64700

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169507)

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

    If Not (gobjRelatorio Is Nothing) Then Error 64701
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 64702

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
    
        Case 64702
        
        Case 64701
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169508)

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

Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String) As Long
'Formata os produtos retornando em sProd_I e sProd_F
'Verifica se os par�metros iniciais s�o maiores que os finais

Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'formata o Produto Inicial
    lErro = CF("Produto_Formata",ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then Error 64703

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata",ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then Error 64704

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambos os produtos est�o preenchidos, o produto inicial n�o pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then Error 64705

    End If

   'data inicial n�o pode ser maior que a data final
    If Trim(DataInicial.ClipText) <> "" And Trim(DataFinal.ClipText) <> "" Then
    
         If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then Error 64706
    
    End If
    
    'O campo almoxarifado deve ser preenchido
    If Trim(Almoxarifado.Text) = "" Then Error 64707
          
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = Err

    Select Case Err

        Case 64703
            ProdutoInicial.SetFocus

        Case 64704
            ProdutoFinal.SetFocus

        Case 64705
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", Err)
            ProdutoInicial.SetFocus
             
        Case 64706
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)
            DataInicial.SetFocus
            
        Case 64707
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_PREENCHIDO1", Err)
      
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169509)

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

    'Se frame selecionado n�o for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame4(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame4(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index

    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169510)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usu�rio

Dim lErro As Long
Dim sProd_I As String
Dim sProd_F As String
Dim sAlmox As String
Dim iAlmox As Integer
Dim objEstoqueMes As New ClassEstoqueMes

On Error GoTo Erro_PreencherRelOp

    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)
    
    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F)
    If lErro <> SUCESSO Then Error 64712

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 64713
         
    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then Error 64714

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then Error 64715
        
    iAlmox = Codigo_Extrai(Almoxarifado.Text)

    lErro = objRelOpcoes.IncluirParametro("NALMOX", CStr(iAlmox))
    If lErro <> AD_BOOL_TRUE Then Error 64716
        
    lErro = objRelOpcoes.IncluirParametro("TALMOXARIFADO", Almoxarifado.Text)
    If lErro <> AD_BOOL_TRUE Then Error 64717

    If DataInicial.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINIC", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 64718

    If DataFinal.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 64719
    
    lErro = objRelOpcoes.IncluirParametro("NPRODSEMMOV", CStr(CheckProdSemMovimento.Value))
    If lErro <> AD_BOOL_TRUE Then Error 64720
    
    lErro = objRelOpcoes.IncluirParametro("NNUMLIVRO", NumLivro.Text)
    If lErro <> AD_BOOL_TRUE Then Error 64721
           
    lErro = objRelOpcoes.IncluirParametro("NPAGINIC", PaginaInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 64722
    
    lErro = objRelOpcoes.IncluirParametro("NQTDEPAG", QtdePaginas.Text)
    If lErro <> AD_BOOL_TRUE Then Error 64723
    
    objEstoqueMes.iFilialEmpresa = giFilialEmpresa
    
    lErro = CF("EstoqueMes_Le_Apurado",objEstoqueMes)
    If lErro <> SUCESSO And lErro <> 46225 Then Error 64724

    If lErro = 46225 Then
        objEstoqueMes.iAno = 0
        objEstoqueMes.iMes = 0
    End If
        
    lErro = objRelOpcoes.IncluirParametro("NANOAPURADO", objEstoqueMes.iAno)
    If lErro <> AD_BOOL_TRUE Then Error 64725
 
    lErro = objRelOpcoes.IncluirParametro("NMESAPURADO", objEstoqueMes.iMes)
    If lErro <> AD_BOOL_TRUE Then Error 64726
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sAlmox, sProd_I, sProd_F)
    If lErro <> SUCESSO Then Error 64727

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 64712 To 64727

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169511)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 64728

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui",gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 64729

        'retira nome das op��es do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as op��es da tela
        Limpar_Tela

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 64728
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 64729

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169512)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 64730

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 64730

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169513)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a op��o de relat�rio com os par�metros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da op��o de relat�rio n�o pode ser vazia
    If ComboOpcoes.Text = "" Then Error 64731

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then Error 64732

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava",gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 64733

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 64731
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 64732, 64733

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169514)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate

    giProdInicial = 0

    lErro = CF("Produto_Perde_Foco",ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO And lErro <> 27095 Then Error 64734
    
    If lErro <> SUCESSO Then Error 64735

    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True


    Select Case Err

        Case 64734

        Case 64735
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169515)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    giProdInicial = 1

    lErro = CF("Produto_Perde_Foco",ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then Error 64736
    
    If lErro <> SUCESSO Then Error 64737

    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True


    Select Case Err

        Case 64736

        Case 64737
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169516)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sAlmox As String, sProd_I As String, sProd_F As String) As Long
'monta a express�o de sele��o de relat�rio

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169517)

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
        If lErro <> SUCESSO Then Error 64738

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True


    Select Case Err

        Case 64738

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169518)

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
        If lErro <> SUCESSO Then Error 64739

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True


    Select Case Err

        Case 64739

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169519)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 64740

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 64740
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169520)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 64741

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 64741
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169521)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 64742

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case Err

        Case 64742
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169522)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 64743

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case Err

        Case 64743
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169523)

    End Select

    Exit Sub

End Sub

Private Function Carrega_Lista_Almoxarifado() As Long
'Carrega a ListBox Almoxarifados

Dim lErro As Long
Dim colAlmoxarifados As New Collection
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_Carrega_Lista_Almoxarifado
    
    'L� C�digos e NomesReduzidos da tabela Almoxarifado e devolve na cole��o
    lErro = CF("Almoxarifados_Le_FilialEmpresa",giFilialEmpresa, colAlmoxarifados)
    If lErro <> SUCESSO Then Error 64745

    'Preenche a ListBox AlmoxarifadoList com os objetos da cole��o
    For Each objAlmoxarifado In colAlmoxarifados
        Almoxarifados.AddItem objAlmoxarifado.iCodigo & SEPARADOR & objAlmoxarifado.sNomeReduzido
        Almoxarifados.ItemData(Almoxarifados.NewIndex) = objAlmoxarifado.iCodigo
    Next

    Carrega_Lista_Almoxarifado = SUCESSO

    Exit Function

Erro_Carrega_Lista_Almoxarifado:

    Carrega_Lista_Almoxarifado = Err

    Select Case Err

        Case 64745

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169524)

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
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 169525)

    End Select

    Exit Sub

End Sub

Private Sub Almoxarifado_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_Almoxarifado_Validate

    If Len(Trim(Almoxarifado.Text)) > 0 Then
       
        'Tenta ler o Almoxarifado (NomeReduzido ou C�digo)
        lErro = TP_Almoxarifado_Le_ComCodigo(Almoxarifado, objAlmoxarifado)
        If lErro <> SUCESSO Then Error 64746

    End If
        
    Exit Sub

Erro_Almoxarifado_Validate:

    Cancel = True


    Select Case Err

        Case 64746
           
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 169526)

    End Select

End Sub

Private Function Critica_Numero(sNumero As String) As Long

Dim lErro As Long

On Error GoTo Erro_Critica_Numero

    If sNumero <> "" Then
    
        lErro = Inteiro_Critica(sNumero)
        If lErro <> SUCESSO Then Error 64747
 
        If CInt(sNumero) < 0 Then Error 64748
 
    End If
    
    Critica_Numero = SUCESSO

    Exit Function

Erro_Critica_Numero:

    Critica_Numero = Err

    Select Case Err

        Case 64747
            
        Case 64748
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", Err, sNumero)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169527)

    End Select

    Exit Function

End Function

Private Sub NumLivro_Validate(Cancel As Boolean)
Dim lErro As Long

On Error GoTo Erro_NumLivro_Validate

    lErro = Critica_Numero(NumLivro.Text)
    If lErro <> SUCESSO Then Error 64749
    
    Exit Sub

Erro_NumLivro_Validate:

    Cancel = True


    Select Case Err

        Case 64749

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169528)

    End Select

    Exit Sub

End Sub

Private Sub PaginaInicial_Validate(Cancel As Boolean)
Dim lErro As Long

On Error GoTo Erro_PaginaInicial_Validate

    lErro = Critica_Numero(PaginaInicial.Text)
    If lErro <> SUCESSO Then Error 64750
    
    Exit Sub

Erro_PaginaInicial_Validate:

    Cancel = True


    Select Case Err

        Case 64750

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169529)

    End Select

    Exit Sub

End Sub

Private Sub QtdePaginas_Validate(Cancel As Boolean)
Dim lErro As Long

On Error GoTo Erro_QtdePaginas_Validate

    lErro = Critica_Numero(QtdePaginas.Text)
    If lErro <> SUCESSO Then Error 64751
    
    Exit Sub

Erro_QtdePaginas_Validate:

    Cancel = True


    Select Case Err

        Case 64751

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169530)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_KARDEX_DIA_L
    Set Form_Load_Ocx = Me
    Caption = "Kardex p/Dia"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpKardexDia_L"
    
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



Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
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

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub LabelAlmoxarifado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelAlmoxarifado, Source, X, Y)
End Sub

Private Sub LabelAlmoxarifado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelAlmoxarifado, Button, Shift, X, Y)
End Sub

Private Sub LabelNumLivro_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNumLivro, Source, X, Y)
End Sub

Private Sub LabelNumLivro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNumLivro, Button, Shift, X, Y)
End Sub

Private Sub labelPagInic_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(labelPagInic, Source, X, Y)
End Sub

Private Sub labelPagInic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(labelPagInic, Button, Shift, X, Y)
End Sub

Private Sub LabelQtdePag_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelQtdePag, Source, X, Y)
End Sub

Private Sub LabelQtdePag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelQtdePag, Button, Shift, X, Y)
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


Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub

