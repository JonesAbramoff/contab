VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl RelOPAcompEntPCOcx 
   ClientHeight    =   5100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8460
   KeyPreview      =   -1  'True
   ScaleHeight     =   5100
   ScaleWidth      =   8460
   Begin VB.Frame Frame10 
      Caption         =   "Item"
      Height          =   1170
      Left            =   4110
      TabIndex        =   45
      Top             =   3650
      Width           =   4035
      Begin VB.ListBox ItensCategoria 
         Height          =   735
         Left            =   840
         Style           =   1  'Checkbox
         TabIndex        =   46
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label labelItens 
         AutoSize        =   -1  'True
         Caption         =   "Itens:"
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
         TabIndex        =   47
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Categoria"
      Height          =   1440
      Left            =   36
      TabIndex        =   38
      Top             =   3492
      Width           =   8385
      Begin VB.ComboBox Categoria 
         Height          =   315
         Left            =   1116
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   450
         Width           =   2865
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   210
         TabIndex        =   39
         Top             =   495
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produtos"
      Height          =   1200
      Left            =   36
      TabIndex        =   37
      Top             =   2235
      Width           =   3090
      Begin MSMask.MaskEdBox ProdutoDe 
         Height          =   300
         Left            =   1020
         TabIndex        =   9
         Top             =   270
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ProdutoAte 
         Height          =   300
         Left            =   1020
         TabIndex        =   10
         Top             =   705
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
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
         Left            =   570
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   44
         Top             =   735
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
         Left            =   600
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   43
         Top             =   330
         Width           =   315
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Comprador"
      Height          =   1200
      Left            =   5685
      TabIndex        =   34
      Top             =   2232
      Width           =   2700
      Begin VB.ComboBox CompradorAte 
         Height          =   315
         Left            =   615
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   720
         Width           =   1824
      End
      Begin VB.ComboBox CompradorDe 
         Height          =   315
         Left            =   615
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   312
         Width           =   1824
      End
      Begin VB.Label Label4 
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
         Height          =   192
         Left            =   120
         TabIndex        =   36
         Top             =   781
         Width           =   360
      End
      Begin VB.Label Label5 
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
         Height          =   192
         Left            =   168
         TabIndex        =   35
         Top             =   373
         Width           =   312
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Fornecedores"
      Height          =   1200
      Left            =   3255
      TabIndex        =   31
      Top             =   2235
      Width           =   2370
      Begin MSMask.MaskEdBox FornecedorDe 
         Height          =   300
         Left            =   915
         TabIndex        =   11
         Top             =   330
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox FornecedorAte 
         Height          =   300
         Left            =   930
         TabIndex        =   12
         Top             =   750
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin VB.Label LabelFornecedorAte 
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
         Left            =   480
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   33
         Top             =   810
         Width           =   360
      End
      Begin VB.Label LabelFornecedorDe 
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
         Left            =   510
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   32
         Top             =   390
         Width           =   315
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Data de Emissão"
      Height          =   1416
      Left            =   6480
      TabIndex        =   26
      Top             =   756
      Width           =   1884
      Begin MSComCtl2.UpDown UpDownDataDe 
         Height          =   315
         Left            =   1485
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   390
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataDe 
         Height          =   315
         Left            =   450
         TabIndex        =   7
         Top             =   390
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataAte 
         Height          =   315
         Left            =   1500
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   840
         Width           =   225
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataAte 
         Height          =   315
         Left            =   450
         TabIndex        =   8
         Top             =   840
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
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
         Left            =   90
         TabIndex        =   30
         Top             =   900
         Width           =   360
      End
      Begin VB.Label Label2 
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
         TabIndex        =   29
         Top             =   450
         Width           =   315
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Pedido de Compra"
      Height          =   1416
      Left            =   36
      TabIndex        =   25
      Top             =   756
      Width           =   3144
      Begin VB.ComboBox Status 
         Height          =   315
         ItemData        =   "RelOPAcompEntPC.ctx":0000
         Left            =   864
         List            =   "RelOPAcompEntPC.ctx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   900
         Width           =   2028
      End
      Begin MSMask.MaskEdBox PCDe 
         Height          =   288
         Left            =   612
         TabIndex        =   2
         Top             =   396
         Width           =   804
         _ExtentX        =   1402
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox PCAte 
         Height          =   288
         Left            =   2052
         TabIndex        =   3
         Top             =   396
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin VB.Label LabelPCAte 
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
         Height          =   192
         Left            =   1656
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   42
         Top             =   444
         Width           =   360
      End
      Begin VB.Label LabelPCDe 
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
         Height          =   192
         Left            =   252
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   41
         Top             =   444
         Width           =   312
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
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
         Height          =   192
         Left            =   252
         TabIndex        =   40
         Top             =   948
         Width           =   576
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filial Empresa"
      Height          =   1416
      Left            =   3240
      TabIndex        =   22
      Top             =   756
      Width           =   3108
      Begin VB.ComboBox FilialEmpresaDe 
         Height          =   288
         Left            =   648
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   396
         Width           =   2316
      End
      Begin VB.ComboBox FilialEmpresaAte 
         Height          =   288
         Left            =   648
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   900
         Width           =   2316
      End
      Begin VB.Label LabelCodFilialAte 
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
         Height          =   192
         Left            =   240
         TabIndex        =   24
         Top             =   948
         Width           =   360
      End
      Begin VB.Label LabelCodFilialDe 
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
         Height          =   192
         Left            =   288
         TabIndex        =   23
         Top             =   444
         Width           =   312
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOPAcompEntPC.ctx":0028
      Left            =   792
      List            =   "RelOPAcompEntPC.ctx":002A
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   204
      Width           =   2580
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
      Height          =   528
      Left            =   3930
      Picture         =   "RelOPAcompEntPC.ctx":002C
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   84
      Width           =   1308
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6135
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   72
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOPAcompEntPC.ctx":012E
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOPAcompEntPC.ctx":02AC
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOPAcompEntPC.ctx":07DE
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOPAcompEntPC.ctx":0968
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
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
      Height          =   252
      Left            =   144
      TabIndex        =   21
      Top             =   228
      Width           =   612
   End
End
Attribute VB_Name = "RelOPAcompEntPCOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'??? ATENCAO: Quem for refazer o Relatório deverá prestar atenção as novas Macros S##
'Os valores foram trocados, pois alguns sairam e outros entraram ...
'Me pergunte ... (Daniel)

'##########

'Alteração dia 26/02/03 Feito por Sergio ...
'Para que os Itens da Categoria Possam ser selecionados não por Intervalos
'e sim quantos forem selecionados...
'a função monta expressão de seleção foi alterada, passando os Itens para a seleção ao Invés de Macros como 'Ex: "S01"'
'Qualquer dúvida perguntar a Shirley...

'###############

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoCodPCDe As AdmEvento
Attribute objEventoCodPCDe.VB_VarHelpID = -1
Private WithEvents objEventoCodPCAte As AdmEvento
Attribute objEventoCodPCAte.VB_VarHelpID = -1
Private WithEvents objEventoFornDe As AdmEvento
Attribute objEventoFornDe.VB_VarHelpID = -1
Private WithEvents objEventoFornAte As AdmEvento
Attribute objEventoFornAte.VB_VarHelpID = -1
Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1

Dim iAlterado As Integer
Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 73523

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 73524

    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 73523

        Case 73524
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166838)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub Limpa_Tela_Rel()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Rel

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 73525

    Status.ListIndex = 0
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus

    Exit Sub

Erro_Limpa_Tela_Rel:

    Select Case gErr

        Case 73525

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166839)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela_Rel

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCodPCDe = New AdmEvento
    Set objEventoCodPCAte = New AdmEvento
    Set objEventoFornDe = New AdmEvento
    Set objEventoFornAte = New AdmEvento
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    
    lErro = Carrega_FilialEmpresa()
    If lErro <> SUCESSO Then gError 108858
    
    lErro = Carrega_Compradores()
    If lErro <> SUCESSO Then gError 108859
    
    lErro = Carrega_Categorias()
    If lErro <> SUCESSO Then gError 108860
    
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoDe)
    If lErro <> SUCESSO Then gError 108861

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoAte)
    If lErro <> SUCESSO Then gError 108862
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 73526, 108858 To 108862
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166840)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

    Set objEventoCodPCDe = Nothing
    Set objEventoCodPCAte = Nothing
    Set objEventoFornDe = Nothing
    Set objEventoFornAte = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    
End Sub

Private Sub Categoria_Click()
'Preenche os itens da categoria selecionada

Dim lErro As Long
Dim colItensCategoria As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As ClassCategoriaProdutoItem

On Error GoTo Erro_Categoria_Click

    'Limpa a Combo de Itens
    ItensCategoria.Clear
    
    If Len(Trim(Categoria.Text)) > 0 Then

        'Preenche o Obj
        objCategoriaProduto.sCategoria = Categoria.List(Categoria.ListIndex)
        
        'Le as categorias do Produto
        lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colItensCategoria)
        If lErro <> SUCESSO And lErro <> 22541 Then gError 108885
                
        For Each objCategoriaProdutoItem In colItensCategoria
            ItensCategoria.AddItem (objCategoriaProdutoItem.sItem)
        Next
        
    End If
    
    Exit Sub

Erro_Categoria_Click:

    Select Case gErr

         Case 108885
         
         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166841)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataAte, iAlterado)
    
End Sub

Private Sub DataDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataDe, iAlterado)
    
End Sub

Private Sub FornecedorAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(FornecedorAte, iAlterado)
    
End Sub

Private Sub FornecedorDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(FornecedorDe, iAlterado)
    
End Sub

Private Sub LabelFornecedorDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelFornecedorDe_Click

    If Len(Trim(FornecedorDe.Text)) > 0 Then
    
        objFornecedor.lCodigo = StrParaLong(FornecedorDe.Text)
        
    End If
    
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornDe)
    
    Exit Sub

Erro_LabelFornecedorDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166842)

    End Select

    Exit Sub
    
End Sub

Private Sub LabelFornecedorAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelFornecedorAte_Click

    If Len(Trim(FornecedorAte.Text)) > 0 Then
    
        objFornecedor.lCodigo = StrParaLong(FornecedorAte.Text)
        
    End If
    
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornAte)
    
    Exit Sub

Erro_LabelFornecedorAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166843)

    End Select

    Exit Sub
    
End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataDe.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataDe.Text)
    If lErro <> SUCESSO Then gError 73527

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 73527
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166844)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataAte.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataAte.Text)
    If lErro <> SUCESSO Then gError 73528

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 73528
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166845)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCodPCAte_evSelecao(obj1 As Object)

Dim objPedCompra As New ClassPedidoCompras

    Set objPedCompra = obj1

    PCAte.Text = CStr(objPedCompra.lCodigo)

    Me.Show

End Sub
Private Sub objEventoCodPCDe_evSelecao(obj1 As Object)

Dim objPedCompra As New ClassPedidoCompras

    Set objPedCompra = obj1

    PCDe.Text = CStr(objPedCompra.lCodigo)

    Me.Show

End Sub


Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 73529

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 73529
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 166846)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 73530

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 73530
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 166847)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 73531

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 73531
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 166848)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 73532

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 73532
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 166849)

    End Select

    Exit Sub

End Sub

Private Sub objEventoFornDe_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor

    Set objFornecedor = obj1

    FornecedorDe.Text = CStr(objFornecedor.lCodigo)

    Me.Show

    Exit Sub

End Sub
Private Sub objEventoFornAte_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor

    Set objFornecedor = obj1

    FornecedorAte.Text = CStr(objFornecedor.lCodigo)

    Me.Show

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 73534

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 73535

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 73536

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 73537

    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 73534
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 73535 To 73537

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166850)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 73538

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPACOMPENTPC")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 73539

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call Limpa_Tela_Rel

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 73538
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 73539

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166851)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 73540

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 73540, 74949

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166852)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim iFilial_I As Integer, iFilial_F As Integer, sForn_I As String, sForn_F As String
Dim sComprador_I As String, sComprador_F As String, sProduto_I As String, sProduto_F As String
Dim sCategoria As String, sItem_I As String, sItem_F As String, iStatus As Integer
Dim sPedido_I As String, sPedido_F As String
Dim colItens As New Collection
Dim iCont As Integer
Dim iIndice As Integer

On Error GoTo Erro_PreencherRelOp

    lErro = Formata_E_Critica_Parametros(sPedido_I, sPedido_F, iFilial_I, iFilial_F, sForn_I, sForn_F, sComprador_I, sComprador_F, sProduto_I, sProduto_F, sItem_I, sItem_F, iStatus)
    If lErro <> SUCESSO Then gError 73541

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 73542

    lErro = objRelOpcoes.IncluirParametro("NCODFILIALINIC", CStr(iFilial_I))
    If lErro <> AD_BOOL_TRUE Then gError 73543

    lErro = objRelOpcoes.IncluirParametro("NCODFILIALFIM", CStr(iFilial_F))
    If lErro <> AD_BOOL_TRUE Then gError 73549
    
    lErro = objRelOpcoes.IncluirParametro("NCODFORNINIC", sForn_I)
    If lErro <> AD_BOOL_TRUE Then gError 73546
    
    lErro = objRelOpcoes.IncluirParametro("NCODFORNFIM", sForn_F)
    If lErro <> AD_BOOL_TRUE Then gError 73552
    
    lErro = objRelOpcoes.IncluirParametro("TPCINI", sPedido_I)
    If lErro <> AD_BOOL_TRUE Then gError 73546
    
    lErro = objRelOpcoes.IncluirParametro("TPCFIM", sPedido_F)
    If lErro <> AD_BOOL_TRUE Then gError 73552
    
    'Preenche data inicial
    If Trim(DataDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATAINIC", DataDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 73548

    'Preenche data final
    If Trim(DataAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATAFIM", DataAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 73554

    lErro = objRelOpcoes.IncluirParametro("TCOMPINI", sComprador_I)
    If lErro <> AD_BOOL_TRUE Then gError 108902
    
    lErro = objRelOpcoes.IncluirParametro("TCOMPFIM", sComprador_F)
    If lErro <> AD_BOOL_TRUE Then gError 108903
    
    lErro = objRelOpcoes.IncluirParametro("TPRODINI", sProduto_I)
    If lErro <> AD_BOOL_TRUE Then gError 108904
    
    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProduto_F)
    If lErro <> AD_BOOL_TRUE Then gError 108905
    
    lErro = objRelOpcoes.IncluirParametro("TCATEG", CStr(Categoria.List(Categoria.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then gError 108906
    
    'Inicia o Contador
    iCont = 0
    
    'Monta o Filtro
    For iIndice = 0 To ItensCategoria.ListCount - 1
        
        'Verifica se o Item da Categoria foi selecionado
        If ItensCategoria.Selected(iIndice) = True Then
            
            'Incrementa o Contador
            iCont = iCont + 1
            
            lErro = objRelOpcoes.IncluirParametro("TITEMDE" & iCont, CStr(ItensCategoria.List(iIndice)))
            If lErro <> AD_BOOL_TRUE Then gError 108907
                            
            colItens.Add CStr(ItensCategoria.List(iIndice))
                             
        End If
            
    Next
    
    'ListIndex -1 por causa do espaço em branco na combo como primeira opcao ...
    lErro = objRelOpcoes.IncluirParametro("NSTATUS", (Status.ListIndex - 1))
    If lErro <> AD_BOOL_TRUE Then gError 108908
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sPedido_I, sPedido_F, iFilial_I, iFilial_F, sForn_I, sForn_F, sComprador_I, sComprador_F, sProduto_I, sProduto_F, iStatus, colItens)
    If lErro <> SUCESSO Then gError 73559

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 73541 To 73559, 108900 To 108908

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166853)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sPedido_I As String, sPedido_F As String, iFilial_I As Integer, iFilial_F As Integer, sForn_I As String, sForn_F As String, sComprador_I As String, sComprador_F As String, sProduto_I As String, sProduto_F As String, sItem_I As String, sItem_F As String, iStatus As Integer)
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long
Dim iProdPreenchido_I As Integer, iProdPreenchido_F As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros

    If Status.ListIndex <> -1 Then
        iStatus = Status.ItemData(Status.ListIndex)
    Else
        iStatus = 0
    End If
    
    'Critica o PC inicial e o Final
    If PCDe.Text <> "" Then
        sPedido_I = CStr(PCDe.Text)
    Else
        sPedido_I = ""
    End If

    If PCAte.Text <> "" Then
        sPedido_F = CStr(PCAte.Text)
    Else
        sPedido_F = ""
    End If

    If sPedido_I <> "" And sPedido_F <> "" Then

        If sPedido_I > sPedido_F Then gError 108940

    End If
    
    'critica Codigo da Filial Inicial e Final
    If FilialEmpresaDe.List(FilialEmpresaDe.ListIndex) <> "" Then
        iFilial_I = Codigo_Extrai(FilialEmpresaDe.List(FilialEmpresaDe.ListIndex))
    Else
        iFilial_I = 0
    End If

    If FilialEmpresaAte.List(FilialEmpresaAte.ListIndex) <> "" Then
        iFilial_F = Codigo_Extrai(FilialEmpresaAte.List(FilialEmpresaAte.ListIndex))
    Else
        iFilial_F = 0
    End If

    If iFilial_I <> 0 And iFilial_F <> 0 Then

        If iFilial_I > iFilial_F Then gError 73560

    End If

    'data inicial não pode ser maior que a final
    If Trim(DataDe.ClipText) <> "" And Trim(DataAte.ClipText) <> "" Then

         If CDate(DataDe.Text) > CDate(DataAte.Text) Then gError 73563

    End If

    'critica Fornecedor Inicial e Final
    If FornecedorDe.Text <> "" Then
        sForn_I = CStr(FornecedorDe.Text)
    Else
        sForn_I = ""
    End If

    If FornecedorAte.Text <> "" Then
        sForn_F = CStr(FornecedorAte.Text)
    Else
        sForn_F = ""
    End If

    If sForn_I <> "" And sForn_F <> "" Then

        If CLng(sForn_I) > CLng(sForn_F) Then gError 73564

    End If
    
    'critica Comprador Inicial e Final
    If CompradorDe.List(CompradorDe.ListIndex) <> "" Then
        sComprador_I = CompradorDe.List(CompradorDe.ListIndex)
    Else
        sComprador_I = ""
    End If

    If CompradorAte.List(CompradorAte.ListIndex) <> "" Then
        sComprador_F = CompradorAte.List(CompradorAte.ListIndex)
    Else
        sComprador_F = ""
    End If

    If sComprador_I <> "" And sComprador_F <> "" Then

        If sComprador_I > sComprador_F Then gError 108890

    End If
    
    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoDe.Text, sProduto_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 108892

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProduto_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoAte.Text, sProduto_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 108893

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProduto_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProduto_I > sProduto_F Then gError 108894

    End If
    
'    'Critica os itens da Categoria
'    If ItemDe.List(ItemDe.ListIndex) <> "" Then sItem_I = ItemDe.List(ItemDe.ListIndex)
'
'    If ItemAte.List(ItemAte.ListIndex) <> "" Then sItem_F = ItemAte.List(ItemAte.ListIndex)
'
'    'Se estão preenchidos => Inicial não pode ser maior do que Final
'    If sItem_I <> "" And sItem_F <> "" Then
'
'        If sItem_I > sItem_F Then gError 108920
'
'    End If
'
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 73560
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            FilialEmpresaDe.SetFocus

        Case 73561
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            FilialEmpresaDe.SetFocus

        Case 73563
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataDe.SetFocus

        Case 73564
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INICIAL_MAIOR", gErr)
            FornecedorDe.SetFocus

        Case 108890
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPRADOR_INICIAL_MAIOR", gErr)
            CompradorDe.SetFocus
            
        Case 108892, 108893
        
        Case 108894
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            ProdutoDe.SetFocus
            
        'Case 108920
            'lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_ITEM_INICIAL_MAIOR", gErr)
            'ItemDe.SetFocus
            
        Case 108940
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PC_INICIAL_MAIOR", gErr)
            PCDe.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166854)

    End Select

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sPedido_I As String, sPedido_F As String, iFilial_I As Integer, iFilial_F As Integer, sForn_I As String, sForn_F As String, sComprador_I As String, sComprador_F As String, sProduto_I As String, sProduto_F As String, iStatus As Integer, colItens As Collection) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long
Dim iCont As Integer
Dim iIndice As Integer

On Error GoTo Erro_Monta_Expressao_Selecao
    
    'Inicia o Contador
    iCont = 0

   If sPedido_I <> "" Then sExpressao = "S01"

   If sPedido_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S02"

    End If

    If iFilial_I <> 0 Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S03"

    End If

    If iFilial_F <> 0 Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S04"

    End If

    If Trim(DataDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S05"

    End If

    If Trim(DataAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S06"

    End If

    If sForn_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S07"

    End If

    If sForn_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S08"

    End If

    If sComprador_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S09"

    End If

    If sComprador_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S10"

    End If

    If sProduto_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S11"

    End If

    If sProduto_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S12"

    End If

'    If sItem_I <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "S13"
'
'    End If
'
'    If sItem_F <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "S14"
'
'    End If
'
'    If sItem_I <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "S15"
'
'    End If
'
'    If sItem_F <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "S16"
'
'    End If
'
    For iIndice = 1 To colItens.Count
        
        iCont = iCont + 1
        
        If iCont = 1 Then
        
            If sExpressao <> "" Then sExpressao = sExpressao & " E "
            sExpressao = sExpressao & "(Item = " & Forprint_ConvTexto((colItens.Item(iIndice)))
        
        Else
            If sExpressao <> "" Then sExpressao = sExpressao & " OU "
            sExpressao = sExpressao & "Item = " & Forprint_ConvTexto((colItens.Item(iIndice)))
            
        End If
    
    Next
    
    If colItens.Count > 0 Then
            
        If sExpressao <> "" Then 'sExpressao = sExpressao & " E "
            sExpressao = sExpressao & ")"
        End If
    
    End If
    
    
    If iStatus = 1 Then 'PENDENTE

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S17"
        
    ElseIf iStatus = 2 Then 'ATENDIDO
    
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S18"

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166855)

    End Select

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim iIndice As Integer
Dim iCont As Integer

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 73566

    'pega Codigo inicial
    lErro = objRelOpcoes.ObterParametro("NCODFILIALINIC", sParam)
    If lErro <> SUCESSO Then gError 73567
    
    For iIndice = 0 To FilialEmpresaDe.ListCount - 1
        If Codigo_Extrai(FilialEmpresaDe.List(iIndice)) = StrParaInt(sParam) Then
            FilialEmpresaDe.ListIndex = iIndice
            Exit For
        End If
    Next
    
    'pega Codigo Filial
    lErro = objRelOpcoes.ObterParametro("NCODFILIALFIM", sParam)
    If lErro <> SUCESSO Then gError 73568

    For iIndice = 0 To FilialEmpresaAte.ListCount - 1
        If Codigo_Extrai(FilialEmpresaAte.List(iIndice)) = StrParaInt(sParam) Then
            FilialEmpresaAte.ListIndex = iIndice
            Exit For
        End If
    Next

    'pega Fornecedor Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFORNINIC", sParam)
    If lErro <> SUCESSO Then gError 73573

    FornecedorDe.Text = sParam

    'pega Fornecedor Final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFORNFIM", sParam)
    If lErro <> SUCESSO Then gError 73574

    FornecedorAte.Text = sParam

    'pega PC Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPCINI", sParam)
    If lErro <> SUCESSO Then gError 73575

    PCDe.Text = sParam
    Call PCDe_Validate(bSGECancelDummy)

    'pega PC Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPCFIM", sParam)
    If lErro <> SUCESSO Then gError 73576

    PCAte.Text = sParam
    Call PCAte_Validate(bSGECancelDummy)

    'pega data  inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAINIC", sParam)
    If lErro <> SUCESSO Then gError 73577

    Call DateParaMasked(DataDe, CDate(sParam))

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAFIM", sParam)
    If lErro <> SUCESSO Then gError 73578

    Call DateParaMasked(DataAte, CDate(sParam))

    'Pega o comprador Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TCOMPINI", sParam)
    If lErro <> SUCESSO Then gError 108912
        
    Call CF("SCombo_Seleciona2", CompradorDe, sParam)
    
    'Pega o comprador Final e exibe
    lErro = objRelOpcoes.ObterParametro("TCOMPFIM", sParam)
    If lErro <> SUCESSO Then gError 108913
    
    Call CF("SCombo_Seleciona2", CompradorAte, sParam)
    
    'pega  codigo do produto inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINI", sParam)
    If lErro <> SUCESSO Then gError 108914
    
    ProdutoDe.PromptInclude = False
    ProdutoDe.Text = sParam
    ProdutoDe.PromptInclude = True
    
    Call ProdutoDe_Validate(bSGECancelDummy)
    
    'pega  codigo do produto final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 108915
    
    ProdutoAte.PromptInclude = False
    ProdutoAte.Text = sParam
    ProdutoAte.PromptInclude = True
    
    Call ProdutoAte_Validate(bSGECancelDummy)
    
    'pega  a categoria e exibe
    lErro = objRelOpcoes.ObterParametro("TCATEG", sParam)
    If lErro <> SUCESSO Then gError 108916
    
    Call CF("SCombo_Seleciona2", Categoria, sParam)
    
    'Para Habilitar os Itens
    Call Categoria_Click

    iCont = 1
    sParam = ""
    
    'Traz o Itemde da Categoria
    lErro = objRelOpcoes.ObterParametro("TITEMDE1", sParam)
    If lErro <> SUCESSO Then gError 114376
    
    Do While sParam <> ""
        
       For iIndice = 0 To ItensCategoria.ListCount - 1
            If Trim(sParam) = Trim(ItensCategoria.List(iIndice)) Then
                ItensCategoria.Selected(iIndice) = True
                Exit For
            End If
        Next
        
        iCont = iCont + 1
        
        lErro = objRelOpcoes.ObterParametro("TITEMDE" & iCont, sParam)
        If lErro <> SUCESSO Then gError 108917

    Loop
        
'    'pega o item inicial da categoria e exibe
'    lErro = objRelOpcoes.ObterParametro("TITEMINI", sParam)
'    If lErro <> SUCESSO Then gError 108917
'
'    Call CF("SCombo_Seleciona2", ItemDe, sParam)
'
'    'pega o item final da categoria e exibe
'    lErro = objRelOpcoes.ObterParametro("TITEMFIM", sParam)
'    If lErro <> SUCESSO Then gError 108918
'
'    Call CF("SCombo_Seleciona2", ItemAte, sParam)
'
    
    lErro = objRelOpcoes.ObterParametro("NSTATUS", sParam)
    If lErro <> SUCESSO Then gError 73580

    'sParam + 1 por causa do espaco em branco na combo
    Status.ListIndex = StrParaInt(sParam) + 1

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 73566 To 73581, 108910 To 108917

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166856)

    End Select

    Exit Function

End Function

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

''    Parent.HelpContextID = IDH_RELOP_REQ
    Set Form_Load_Ocx = Me
    Caption = "Acompanhamento de Entrega por Pedido de Compras"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpAcompEntPC"

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

        If Me.ActiveControl Is FornecedorDe Then
            Call LabelFornecedorDe_Click

        ElseIf Me.ActiveControl Is FornecedorAte Then
            Call LabelFornecedorAte_Click
    
        ElseIf Me.ActiveControl Is ProdutoDe Then
            Call LabelProdutoDe_Click
            
        ElseIf Me.ActiveControl Is ProdutoAte Then
            Call LabelProdutoAte_Click
            
        ElseIf Me.ActiveControl Is PCDe Then
            Call LabelPCDe_Click
            
        ElseIf Me.ActiveControl Is PCAte Then
            Call LabelPCAte_Click

        End If

    End If

End Sub

Private Sub ProdutoDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_ProdutoDe_Validate

    If Len(Trim(ProdutoDe.ClipText)) > 0 Then
        
        lErro = CF("Produto_Formata", ProdutoDe.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 108862
        
        objProduto.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 108863
                
        If lErro = 28030 Then gError 108864
        
    End If
    
    Exit Sub
    
Erro_ProdutoDe_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 108862, 108863
        
        Case 108864
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166857)
            
    End Select
    
End Sub

Private Sub ProdutoAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_ProdutoAte_Validate

    If Len(Trim(ProdutoAte.ClipText)) > 0 Then
        
        lErro = CF("Produto_Formata", ProdutoAte.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 108865
        
        objProduto.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 108866
        
        If lErro = 28030 Then gError 108867
        
    End If
    
    Exit Sub
    
Erro_ProdutoAte_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 108865, 108866
        
        Case 108867
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166858)
            
    End Select
    
End Sub

Private Sub LabelProdutoDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String

On Error GoTo Erro_LabelProdutoDe_Click
    
    If Len(Trim(ProdutoDe.Text)) > 0 Then
        
        'Preenche com o Produto da tela
        lErro = CF("Produto_Formata", ProdutoDe.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 108868
        
        objProduto.sCodigo = sProdutoFormatado
        
    End If
    
    'Chama Tela ProdutoCompraLista
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoProdutoDe)

   Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 108868
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166859)

    End Select

End Sub

Private Sub LabelProdutoAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String

On Error GoTo Erro_LabelProdutoAte_Click
    
    If Len(Trim(ProdutoAte.Text)) > 0 Then
        
        'Preenche com o Produto da tela
        lErro = CF("Produto_Formata", ProdutoAte.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 108869
        
        objProduto.sCodigo = sProdutoFormatado
        
    End If
    
    'Chama Tela ProdutoCompraLista
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoProdutoAte)

   Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 108869
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166860)

    End Select

End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim sProdutoMascarado As String
Dim lErro As Long

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 108870
    
    ProdutoAte.Text = sProdutoMascarado

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr
    
        Case 108870
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166861)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoProdutoDe_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim sProdutoMascarado As String
Dim lErro As Long

On Error GoTo Erro_objEventoProdutoDe_evSelecao

    Set objProduto = obj1

    lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 108871
    
    ProdutoDe.Text = sProdutoMascarado

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr
    
        Case 108871
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166862)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub PCDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(PCDe, iAlterado)
    
End Sub

Private Sub PCAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(PCAte, iAlterado)
    
End Sub

Private Sub ProdutoDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(ProdutoDe, iAlterado)
    
End Sub

Private Sub ProdutoAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(ProdutoAte, iAlterado)
    
End Sub

Private Sub PCDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PCDe_Validate

    'Verifica se está preenchido
    If Len(Trim(PCDe.Text)) = 0 Then Exit Sub

    'Nao pode ser nulo
    If StrParaLong(PCDe.Text) = 0 Then gError 108872

    Exit Sub

Erro_PCDe_Validate:

    Cancel = True

    Select Case gErr

        Case 108872
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166863)

    End Select

End Sub

Private Sub PCAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PCAte_Validate

    'Verifica se está preenchido
    If Len(Trim(PCAte.Text)) = 0 Then Exit Sub

    'Nao pode ser nulo
    If StrParaLong(PCAte.Text) = 0 Then gError 108872

    Exit Sub

Erro_PCAte_Validate:

    Cancel = True

    Select Case gErr

        Case 108872
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166864)

    End Select

End Sub

Private Function Carrega_FilialEmpresa() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objFilialEmpresa As New AdmFiliais
Dim colFiliais As New Collection

On Error GoTo Erro_Carrega_FilialEmpresa

    'Faz a Leitura das Filiais
    lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, colFiliais)
    If lErro <> SUCESSO Then gError 108874
    
    FilialEmpresaDe.AddItem ("")
    FilialEmpresaAte.AddItem ("")
    
    'Carrega as combos
    For Each objFilialEmpresa In colFiliais
        
        'Se nao for a EMPRESA_TODA
        If objFilialEmpresa.iCodFilial <> EMPRESA_TODA Then
            
            FilialEmpresaDe.AddItem (objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome)
            FilialEmpresaAte.AddItem (objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome)
            
        End If
        
    Next

    Carrega_FilialEmpresa = SUCESSO
    
    Exit Function
    
Erro_Carrega_FilialEmpresa:

    Carrega_FilialEmpresa = gErr
    
    Select Case gErr
    
        Case 108874

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166865)
    
    End Select

End Function

Private Function Carrega_Compradores() As Long

Dim lErro As Long
Dim objComprador As New ClassComprador
Dim colComprador As New Collection

On Error GoTo Erro_Carrega_Compradores

    'Le os compradores
    lErro = CF("Comprador_Le_Todos", colComprador)
    If lErro <> SUCESSO And lErro <> 50126 Then gError 108875
    
    'Se nao encontrou => Erro
    If lErro = 50126 Then gError 108876
    
    CompradorDe.AddItem ("")
    CompradorAte.AddItem ("")
    
    'Carrega as combos de Compradores
    For Each objComprador In colComprador
    
        CompradorDe.AddItem objComprador.sCodUsuario
        CompradorAte.AddItem objComprador.sCodUsuario
    
    Next
    
    Carrega_Compradores = SUCESSO
    
    Exit Function
    
Erro_Carrega_Compradores:

    Carrega_Compradores = gErr
    
    Select Case gErr
    
        Case 108875
        
        Case 108876
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPRADOR_NAO_CADASTRADO2", gErr)
            '??? Não existe comprador cadastrado.

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166866)
    
    End Select

End Function

Private Function Carrega_Categorias() As Long

Dim lErro As Long
Dim objCategoria As New ClassCategoriaProduto
Dim colCategorias As New Collection

On Error GoTo Erro_Carrega_Categorias
    
    'Le a categoria
    lErro = CF("CategoriasProduto_Le_Todas", colCategorias)
    If lErro <> SUCESSO And lErro <> 22542 Then gError 108877
    
    'Se nao encontrou => Erro
    If lErro = 22542 Then gError 108878
    
    Categoria.AddItem ("")
    
    'Carrega as combos de Categorias
    For Each objCategoria In colCategorias
    
        Categoria.AddItem objCategoria.sCategoria
        
    Next
    
    Carrega_Categorias = SUCESSO
    
    Exit Function
    
Erro_Carrega_Categorias:

    Carrega_Categorias = gErr
    
    Select Case gErr
    
        Case 108877
        
        Case 108878
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_NAO_CADASTRADA", gErr)
            '??? Não existe categoria de produto cadastrada.

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166867)
    
    End Select

End Function

Private Sub LabelPCAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objPedCompra As New ClassPedidoCompras

On Error GoTo Erro_LabelPCAte_Click

    If Len(Trim(PCAte.Text)) > 0 Then
        'Preenche com o Pedido de Compra da tela
        objPedCompra.lCodigo = StrParaLong(PCAte.Text)
    End If

    'Chama Tela PedComprasTodosLista
    Call Chama_Tela("PedidoComprasLista", colSelecao, objPedCompra, objEventoCodPCAte)

   Exit Sub

Erro_LabelPCAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166868)

    End Select

End Sub

Private Sub LabelPCDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objPedCompra As New ClassPedidoCompras

On Error GoTo Erro_LabelPCDe_Click

    If Len(Trim(PCDe.Text)) > 0 Then
        'Preenche com o Pedido de Compra da tela
        objPedCompra.lCodigo = StrParaLong(PCDe.Text)
    End If

    'Chama Tela PedComprasTodosLista
    Call Chama_Tela("PedidoComprasLista", colSelecao, objPedCompra, objEventoCodPCDe)

   Exit Sub

Erro_LabelPCDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166869)

    End Select

End Sub
