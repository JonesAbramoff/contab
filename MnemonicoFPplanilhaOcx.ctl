VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl MnemonicoFPPlanilhaOcx 
   ClientHeight    =   6045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9420
   ScaleHeight     =   6045
   ScaleWidth      =   9420
   Begin VB.ComboBox TipoPlanilha 
      Height          =   315
      Left            =   1740
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   255
      Width           =   3570
   End
   Begin VB.CommandButton BotaoMnemonicoFPreco 
      Caption         =   "Mnem�nicos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Lista de F�rmulas Utilizadas na Forma��o de Pre�o"
      Top             =   5385
      Width           =   1380
   End
   Begin VB.TextBox MnemonicoDesc 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2325
      MaxLength       =   255
      TabIndex        =   8
      Top             =   3420
      Width           =   3240
   End
   Begin VB.Frame Frame1 
      Caption         =   "Escopo"
      Height          =   630
      Left            =   60
      TabIndex        =   31
      Top             =   765
      Width           =   5265
      Begin VB.OptionButton EscopoGeral 
         Caption         =   "Geral"
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
         Left            =   210
         TabIndex        =   0
         Top             =   210
         Width           =   900
      End
      Begin VB.OptionButton EscopoCategoria 
         Caption         =   "Categoria"
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
         Left            =   1165
         TabIndex        =   1
         Top             =   210
         Width           =   1185
      End
      Begin VB.OptionButton EscopoProduto 
         Caption         =   "Produto"
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
         Left            =   2405
         TabIndex        =   2
         Top             =   210
         Width           =   990
      End
      Begin VB.OptionButton EscopoTabela 
         Caption         =   "Tabela de Pre�o"
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
         Left            =   3450
         TabIndex        =   3
         Top             =   210
         Width           =   1740
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7155
      ScaleHeight     =   495
      ScaleWidth      =   2070
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   105
      Width           =   2130
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   75
         Picture         =   "MnemonicoFPplanilhaOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   585
         Picture         =   "MnemonicoFPplanilhaOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1095
         Picture         =   "MnemonicoFPplanilhaOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1590
         Picture         =   "MnemonicoFPplanilhaOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame FrameTabelaPreco 
      Caption         =   "Tabela de Pre�o"
      Height          =   660
      Left            =   75
      TabIndex        =   25
      Top             =   1515
      Visible         =   0   'False
      Width           =   9240
      Begin VB.ComboBox TabelaPreco 
         Height          =   315
         Left            =   1395
         TabIndex        =   4
         Text            =   "TabelaPreco"
         Top             =   240
         Width           =   1875
      End
      Begin MSMask.MaskEdBox Produto1 
         Height          =   315
         Left            =   4200
         TabIndex        =   5
         Top             =   240
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         PromptChar      =   " "
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Descri��o:"
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
         Left            =   5970
         TabIndex        =   29
         Top             =   300
         Width           =   930
      End
      Begin VB.Label LabelDescricao1 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   6960
         TabIndex        =   28
         Top             =   240
         Width           =   2145
      End
      Begin VB.Label LabelProduto1 
         AutoSize        =   -1  'True
         Caption         =   "Produto:"
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
         Left            =   3405
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   27
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tabela Pre�o:"
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
         Left            =   135
         TabIndex        =   26
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Frame FrameCategoria 
      Caption         =   "Categoria de Produto"
      Height          =   660
      Left            =   75
      TabIndex        =   20
      Top             =   1530
      Visible         =   0   'False
      Width           =   9240
      Begin VB.ComboBox ComboCategoriaProdutoItem 
         Height          =   315
         Left            =   5625
         TabIndex        =   21
         Text            =   "ComboCategoriaProdutoItem"
         Top             =   210
         Width           =   2610
      End
      Begin VB.Label LabelCategoria 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pre�o"
         Height          =   315
         Left            =   2985
         TabIndex        =   24
         Top             =   210
         Width           =   555
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Left            =   2040
         TabIndex        =   23
         Top             =   255
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Item:"
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
         Left            =   5130
         TabIndex        =   22
         Top             =   255
         Width           =   435
      End
   End
   Begin VB.Frame FrameProduto 
      Caption         =   "Produto"
      Height          =   660
      Left            =   75
      TabIndex        =   15
      Top             =   1500
      Visible         =   0   'False
      Width           =   9240
      Begin MSMask.MaskEdBox Produto 
         Height          =   315
         Left            =   1770
         TabIndex        =   16
         Top             =   225
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         PromptChar      =   " "
      End
      Begin VB.Label LabelProduto 
         AutoSize        =   -1  'True
         Caption         =   "Produto:"
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
         Left            =   945
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   19
         Top             =   285
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Descri��o:"
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
         Left            =   4365
         TabIndex        =   18
         Top             =   270
         Width           =   930
      End
      Begin VB.Label LabelDescricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5445
         TabIndex        =   17
         Top             =   240
         Width           =   3570
      End
   End
   Begin VB.TextBox Expressao 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   5415
      MaxLength       =   255
      TabIndex        =   9
      Top             =   3420
      Width           =   3360
   End
   Begin VB.TextBox Mnemonico 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   720
      MaxLength       =   20
      TabIndex        =   33
      Top             =   3330
      Width           =   1605
   End
   Begin MSFlexGridLib.MSFlexGrid GridItens 
      Height          =   2880
      Left            =   105
      TabIndex        =   7
      Top             =   2400
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   5080
      _Version        =   393216
      Rows            =   10
      Cols            =   4
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Planilha: "
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
      TabIndex        =   32
      Top             =   315
      Width           =   1515
   End
End
Attribute VB_Name = "MnemonicoFPPlanilhaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_objUserControl As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iFrameAtual As Integer
Dim iTipo As Integer

Dim objGrid As AdmGrid
Dim iGrid_Mnemonico_Col As Integer
Dim iGrid_Expressao_Col As Integer
Dim iGrid_MnemonicoDesc_Col As Integer

Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoProduto1 As AdmEvento
Attribute objEventoProduto1.VB_VarHelpID = -1
Private WithEvents objEventoMnemonicoFPTipo As AdmEvento
Attribute objEventoMnemonicoFPTipo.VB_VarHelpID = -1

Private Sub BotaoMnemonicoFPreco_Click()

Dim lErro As Long
Dim objMnemonicoFPTipo As New ClassMnemonicoFPTipo
Dim colSelecao As Collection

On Error GoTo Erro_BotaoMnemonicoFPreco_Click

    lErro = Move_Tela_Memoria(objMnemonicoFPTipo)
    If lErro <> SUCESSO Then gError 92389

    'Chama a Tela ProdutoVendaLista
    Call Chama_Tela("MnemonicoFPlanilhaLista", colSelecao, objMnemonicoFPTipo, objEventoMnemonicoFPTipo)

    Exit Sub
    
Erro_BotaoMnemonicoFPreco_Click:

    Select Case gErr
    
        Case 92389
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162724)
            
    End Select

    Exit Sub
    
End Sub

Private Sub objEventoMnemonicoFPTipo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objMnemonicoFPTipo As ClassMnemonicoFPTipo
Dim colMnemonicoFPTipo As New Collection

On Error GoTo Erro_objEventoMnemonicoFPTipo_evSelecao

    Set objMnemonicoFPTipo = obj1

    'L� o Mnemonico
    lErro = CF("MnemonicoFPTipo_Le_Todos", objMnemonicoFPTipo, colMnemonicoFPTipo)
    If lErro <> SUCESSO Then gError 92390

    lErro = Traz_MnemonicoFPreco_Tela(colMnemonicoFPTipo)
    If lErro <> SUCESSO Then gError 92391

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoMnemonicoFPTipo_evSelecao:

    Select Case gErr

        Case 92390, 92391

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162725)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Private Sub ComboCategoriaProdutoItem_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub EscopoGeral_Click()
    
Dim lErro As Long
    
On Error GoTo Erro_EscopoGeral_Click

    If iFrameAtual = MNEMONICOFPRECO_ESCOPO_GERAL Then gError 116817
    
    'verifica se existe a necessidade de salvar o escopo antigo
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 92333
    
    iFrameAtual = MNEMONICOFPRECO_ESCOPO_GERAL

    Call Retorna_Frame_Anterior
    
    Call TipoPlanilha_Click

    iAlterado = 0

    Exit Sub
    
Erro_EscopoGeral_Click:

    Select Case gErr

        Case 92333
            Call Retorna_Frame_Anterior

        Case 116817
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162726)
            
    End Select
        
    Exit Sub
    
End Sub

Private Sub EscopoCategoria_Click()

Dim lErro As Long
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As ClassCategoriaProdutoItem
Dim colCategoria As New Collection
Dim sCategoriaItem As String
Dim iIndice As Integer
    
On Error GoTo Erro_EscopoCategoria_Click
    
    If iFrameAtual = MNEMONICOFPRECO_ESCOPO_CATEGORIA Then gError 116819
    
    'verifica se existe a necessidade de salvar o escopo antigo
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 92334
    
    iFrameAtual = MNEMONICOFPRECO_ESCOPO_CATEGORIA

    Call Retorna_Frame_Anterior

    sCategoriaItem = ComboCategoriaProdutoItem.Text

    ComboCategoriaProdutoItem.Clear
    
    'Preenche o objeto com a Categoria
     objCategoriaProduto.sCategoria = LabelCategoria.Caption

     'L� Categoria De Produto no BD
     lErro = CF("CategoriaProduto_Le", objCategoriaProduto)
     If lErro <> SUCESSO And lErro <> 22540 Then gError 92335
    
    'Categoria n�o est� cadastrada
     If lErro <> SUCESSO Then gError 92336

    'L� os dados de itens de categorias de produto
    lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colCategoria)
    If lErro <> SUCESSO Then gError 92337

    'Preenche Valor Inicial e final
    For Each objCategoriaProdutoItem In colCategoria

        ComboCategoriaProdutoItem.AddItem (objCategoriaProdutoItem.sItem)

    Next

    For iIndice = 0 To ComboCategoriaProdutoItem.ListCount - 1
        If ComboCategoriaProdutoItem.List(iIndice) = sCategoriaItem Then
            ComboCategoriaProdutoItem.ListIndex = iIndice
            Exit For
        End If
    Next
    
    'Atualiza Grid
    Call TipoPlanilha_Click
    
    iAlterado = 0

    Exit Sub
    
Erro_EscopoCategoria_Click:

    Select Case gErr

        Case 92334
            Call Retorna_Frame_Anterior

        Case 92335, 92337, 116819

        Case 92336
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_INEXISTENTE", gErr)
            Call Grid_Limpa(objGrid)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162727)
            
    End Select
        
    Exit Sub

End Sub

Private Sub EscopoProduto_Click()

Dim lErro As Long
    
On Error GoTo Erro_EscopoProduto_Click
    
    If iFrameAtual = MNEMONICOFPRECO_ESCOPO_PRODUTO Then gError 116816
    
    'verifica se existe a necessidade de salvar o escopo antigo
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 92338
    
    iFrameAtual = MNEMONICOFPRECO_ESCOPO_PRODUTO

    Call Retorna_Frame_Anterior
    
    'Atualiza Grid
    Call TipoPlanilha_Click

    iAlterado = 0

    Exit Sub
    
Erro_EscopoProduto_Click:

    Select Case gErr

        Case 92338
            Call Retorna_Frame_Anterior
            
        Case 116816

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162728)
            
    End Select
        
    Exit Sub

End Sub

Private Sub EscopoTabela_Click()

Dim lErro As Long
Dim sTabela As String
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome
Dim iIndice As Integer
    
On Error GoTo Erro_EscopoTabela_Click

    If iFrameAtual = MNEMONICOFPRECO_ESCOPO_TABPRECO Then gError 116818
    
    'verifica se existe a necessidade de salvar o escopo antigo
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 92339
    
    iFrameAtual = MNEMONICOFPRECO_ESCOPO_TABPRECO

    Call Retorna_Frame_Anterior

    sTabela = Tabelapreco.Text

    Tabelapreco.Clear

    'L� cada codigo e descricao da tabela TabelasDePreco
    lErro = CF("Cod_Nomes_Le", "TabelasDePrecoVenda", "Codigo", "Descricao", STRING_TABELA_PRECO_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 92340

    'Preenche a ComboBox TabelaPreco com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao
        Tabelapreco.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        Tabelapreco.ItemData(Tabelapreco.NewIndex) = objCodigoDescricao.iCodigo
    Next

    For iIndice = 0 To Tabelapreco.ListCount - 1
        If Tabelapreco.List(iIndice) = sTabela Then
            Tabelapreco.ListIndex = iIndice
            Exit For
        End If
    Next
    
    'Atualiza Grid
    Call TipoPlanilha_Click

    iAlterado = 0

    Exit Sub
    
Erro_EscopoTabela_Click:

    Select Case gErr

        Case 92339
            Call Retorna_Frame_Anterior

        Case 92340, 116818

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162729)
            
    End Select
        
    Exit Sub
    
End Sub

Public Sub TabelaPreco_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Retorna_Frame_Anterior()

    Select Case iFrameAtual
    
        Case MNEMONICOFPRECO_ESCOPO_GERAL
            FrameCategoria.Visible = False
            FrameProduto.Visible = False
            FrameTabelaPreco.Visible = False
            EscopoGeral.Value = True

        Case MNEMONICOFPRECO_ESCOPO_CATEGORIA
            FrameCategoria.Visible = True
            FrameProduto.Visible = False
            FrameTabelaPreco.Visible = False
            EscopoCategoria.Value = True
        
        Case MNEMONICOFPRECO_ESCOPO_PRODUTO
            FrameCategoria.Visible = False
            FrameProduto.Visible = True
            FrameTabelaPreco.Visible = False
            EscopoProduto.Value = True
        
        Case MNEMONICOFPRECO_ESCOPO_TABPRECO
            FrameCategoria.Visible = False
            FrameProduto.Visible = False
            FrameTabelaPreco.Visible = True
            EscopoTabela.Value = True
        
    End Select
        
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    iFrameAtual = MNEMONICOFPRECO_ESCOPO_GERAL
    
    EscopoGeral.Value = True
    
    Set objEventoProduto = New AdmEvento
    Set objEventoProduto1 = New AdmEvento
    Set objEventoMnemonicoFPTipo = New AdmEvento
    
    'Carrega o Combo TipoPlanilha
    lErro = Carga_Combobox_TipoPlanilha
    If lErro <> SUCESSO Then gError 116811
    
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 92341
    
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto1)
    If lErro <> SUCESSO Then gError 92342
    
    'Inicializa o Grid
    Set objGrid = New AdmGrid
    
    lErro = Inicializa_Grid_Itens(objGrid)
    If lErro <> SUCESSO Then gError 92343
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 92341, 92342, 92343, 116811
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162730)
    
    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Carga_Combobox_TipoPlanilha() As Long
'carrega a combobox que contem os tipos de planilhas disponiveis

Dim colTipoPlanilha As New Collection
Dim objTipoPlanilha As ClassTipoPlanilha
Dim lErro As Long
    
On Error GoTo Erro_Carga_Combobox_TipoPlanilha
        
    'leitura dos tipos de planilha no BD
    lErro = CF("TipoPlanilha_Le_Todos", colTipoPlanilha)
    If lErro <> SUCESSO Then gError 116807
    
    For Each objTipoPlanilha In colTipoPlanilha
        
        If objTipoPlanilha.iGrupo = 0 Or objTipoPlanilha.iGrupo = 1 Then
        
            TipoPlanilha.AddItem objTipoPlanilha.iCodigo & SEPARADOR & objTipoPlanilha.sDescricao
                
        End If
        
    Next
    
    Carga_Combobox_TipoPlanilha = SUCESSO

    Exit Function

Erro_Carga_Combobox_TipoPlanilha:

    Carga_Combobox_TipoPlanilha = gErr

    Select Case gErr

        Case 116807
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162731)

    End Select
    
    Exit Function

End Function

Private Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long

    'tela em quest�o
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Mnem�nico")
    objGridInt.colColuna.Add ("Descri��o")
    objGridInt.colColuna.Add ("Valor")

   'campos de edi��o do grid
    objGridInt.colCampo.Add (Mnemonico.Name)
    objGridInt.colCampo.Add (MnemonicoDesc.Name)
    objGridInt.colCampo.Add (Expressao.Name)

    'Indica onde est�o situadas as colunas do grid
    iGrid_Mnemonico_Col = 1
    iGrid_MnemonicoDesc_Col = 2
    iGrid_Expressao_Col = 3

    objGridInt.objGrid = GridItens

    'todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_MNEMONICOFPRECO + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 10

    GridItens.ColWidth(0) = 400

    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Itens = SUCESSO

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a cr�tica da c�lula do grid que est� deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        Select Case GridItens.Col

            Case iGrid_Mnemonico_Col

                lErro = Saida_Celula_Mnemonico(objGridInt)
                If lErro <> SUCESSO Then gError 92344

            Case iGrid_MnemonicoDesc_Col

                lErro = Saida_Celula_MnemonicoDesc(objGridInt)
                If lErro <> SUCESSO Then gError 92345

            Case iGrid_Expressao_Col

                lErro = Saida_Celula_Expressao(objGridInt)
                If lErro <> SUCESSO Then gError 92346

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 92347

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 92347
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 92344, 92345, 92346

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162732)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Mnemonico(objGridInt As AdmGrid) As Long
'faz a critica da celula Titulo do grid que est� deixando de ser a corrente

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Saida_Celula_Mnemonico

    Set objGridInt.objControle = Mnemonico

    If Len(Trim(Mnemonico.Text)) > 0 Then

        For iLinha = 1 To objGridInt.iLinhasExistentes
            If iLinha <> objGridInt.objGrid.Row Then
                If objGridInt.objGrid.TextMatrix(iLinha, iGrid_Mnemonico_Col) = Mnemonico.Text Then gError 92369
            End If
        Next
    
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
        
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 92348
    
    Saida_Celula_Mnemonico = SUCESSO

    Exit Function

Erro_Saida_Celula_Mnemonico:

    Saida_Celula_Mnemonico = gErr

    Select Case gErr

        Case 92348
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 92369
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MNEMONICO_JA_DEFINIDO", gErr, Mnemonico.Text, iLinha)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162733)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MnemonicoDesc(objGridInt As AdmGrid) As Long
'faz a critica da celula Titulo do grid que est� deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MnemonicoDesc

    Set objGridInt.objControle = MnemonicoDesc

    If Len(Trim(MnemonicoDesc.Text)) > 0 Then

        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
        
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 92349
    
    Saida_Celula_MnemonicoDesc = SUCESSO

    Exit Function

Erro_Saida_Celula_MnemonicoDesc:

    Saida_Celula_MnemonicoDesc = gErr

    Select Case gErr

        Case 92349
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162734)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Expressao(objGridInt As AdmGrid) As Long
'Faz a cr�tica da c�lula Item do grid que est� deixando de ser a corrente

Dim lErro As Long
Dim iInicio As Integer
Dim iTamanho As Integer
Dim sExpressao As String

On Error GoTo Erro_Saida_Celula_Expressao

    Set objGridInt.objControle = Expressao

    'Verifica o preenchimento de Quantidade
    If Len(Trim(Expressao.Text)) > 0 Then
        
        'VAlida a quantidae informada
        lErro = Valor_NaoNegativo_Critica(Expressao.Text)
        If lErro <> SUCESSO Then gError 92350
        
'        Expressao.Text = Format(CDbl(Expressao.Text), "Standard")

        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
        
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 92351

    Saida_Celula_Expressao = SUCESSO

    Exit Function
    
Erro_Saida_Celula_Expressao:

    Saida_Celula_Expressao = gErr

    Select Case gErr

        Case 92350, 92351
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162735)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Libera as vari�veis globais da tela
    Set objEventoProduto = Nothing
    Set objEventoProduto1 = Nothing
    
    Set objGrid = Nothing
    
    'Fecha o Comando de Setas
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub LabelProduto_Click()

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As Collection

On Error GoTo Erro_LabelProduto_Click

    lErro = CF("Produto_Formata", Produto.Text, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 92352
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    objProduto.sCodigo = sProduto

    'Chama a Tela ProdutoVendaLista
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoProduto)

    Exit Sub
    
Erro_LabelProduto_Click:

    Select Case gErr
    
        Case 92352
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162736)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'L� o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 92353

    'Se n�o achou o Produto --> erro
    If lErro = 28030 Then gError 92354

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, Produto, LabelDescricao)
    If lErro <> SUCESSO Then gError 92355

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 92353, 92355

        Case 92354
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162737)

    End Select

    Exit Sub

End Sub

Private Sub LabelProduto1_Click()

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As Collection

On Error GoTo Erro_LabelProduto1_Click

    lErro = CF("Produto_Formata", Produto1.Text, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 92356
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    objProduto.sCodigo = sProduto

    'Chama a Tela ProdutoVendaLista
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoProduto1)

    Exit Sub
    
Erro_LabelProduto1_Click:

    Select Case gErr
    
        Case 92356
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162738)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoProduto1_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProduto1_evSelecao

    Set objProduto = obj1

    'L� o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 92357

    'Se n�o achou o Produto --> erro
    If lErro = 28030 Then gError 92358

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, Produto1, LabelDescricao1)
    If lErro <> SUCESSO Then gError 92359

    Me.Show

    Exit Sub

Erro_objEventoProduto1_evSelecao:

    Select Case gErr

        Case 92357, 92359

        Case 92358
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162739)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objMnemonicoFPTipo As ClassMnemonicoFPTipo) As Long

Dim lErro As Long
Dim colMnemonicoFPTipo As New Collection

On Error GoTo Erro_Trata_Parametros
    
    'Se h� uma formula de Forma��o de Pre�o selecionada
    If Not (objMnemonicoFPTipo Is Nothing) Then

        'Verifica se a formula existe no BD
        lErro = CF("MnemonicoFPTipo_Le_Todos", objMnemonicoFPTipo, colMnemonicoFPTipo)
        If lErro <> SUCESSO Then gError 92360

        'Se a formula existe
        If lErro = SUCESSO Then

            lErro = Traz_MnemonicoFPreco_Tela(colMnemonicoFPTipo)
            If lErro <> SUCESSO Then gError 92361

        End If

    End If

    iAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
    
        Case 92360, 92361
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162740)
    
    End Select
    
    iAlterado = 0
    
    Exit Function

End Function

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 92362
    
    Call Limpa_Tela_MnemonicoFPreco

    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case gErr
    
        Case 92362
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162741)
            
    End Select
    
    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long
'grava os dados da tela

Dim lErro As Long
Dim objMnemonicoFPTipo As New ClassMnemonicoFPTipo
Dim objMnemonicoFPTipo1 As ClassMnemonicoFPTipo
Dim sExpressao As String
Dim iInicio As Integer
Dim iTamanho As Integer
Dim sProduto As String
Dim iPreenchido As Integer
Dim iLinha As Integer
Dim colMnemonicoFPTipo As New Collection
    
On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    objMnemonicoFPTipo.iFilialEmpresa = giFilialEmpresa
    objMnemonicoFPTipo.iEscopo = iFrameAtual
            
    'verifica se tipo de planilha foi preenchido
    If TipoPlanilha.ListIndex = -1 Then gError 116808
    
    If objMnemonicoFPTipo.iEscopo = MNEMONICOFPRECO_ESCOPO_CATEGORIA Then
        
        If Len(ComboCategoriaProdutoItem.Text) = 0 Then gError 92360
        
        objMnemonicoFPTipo.sItemCategoria = ComboCategoriaProdutoItem.Text
        
    ElseIf objMnemonicoFPTipo.iEscopo = MNEMONICOFPRECO_ESCOPO_PRODUTO Then
    
        If Len(Trim(Produto.Text)) = 0 Then gError 92361
        
        lErro = CF("Produto_Formata", Produto.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 92362
        
        If iPreenchido = PRODUTO_PREENCHIDO Then objMnemonicoFPTipo.sProduto = sProduto
        
    ElseIf objMnemonicoFPTipo.iEscopo = MNEMONICOFPRECO_ESCOPO_TABPRECO Then
    
        If Len(Tabelapreco.Text) = 0 Then gError 92363
        If Len(Trim(Produto1.Text)) = 0 Then gError 92364
    
        objMnemonicoFPTipo.iTabelaPreco = Codigo_Extrai(Tabelapreco.Text)
        
        lErro = CF("Produto_Formata", Produto1.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 92365
        
        If iPreenchido = PRODUTO_PREENCHIDO Then objMnemonicoFPTipo.sProduto = sProduto
        
    End If
    
    'se n�o houver nenhuma linha preenchida no grid ==> erro
    If objGrid.iLinhasExistentes = 0 Then gError 92366
    
    For iLinha = 1 To objGrid.iLinhasExistentes
    
        If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_Mnemonico_Col))) = 0 Then gError 92367
        
'        If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_Expressao_Col))) = 0 Then gError 92368
    
        Set objMnemonicoFPTipo1 = New ClassMnemonicoFPTipo
        
        objMnemonicoFPTipo1.iFilialEmpresa = objMnemonicoFPTipo.iFilialEmpresa
        objMnemonicoFPTipo1.iEscopo = objMnemonicoFPTipo.iEscopo
        objMnemonicoFPTipo1.sItemCategoria = objMnemonicoFPTipo.sItemCategoria
        objMnemonicoFPTipo1.sProduto = objMnemonicoFPTipo.sProduto
        objMnemonicoFPTipo1.iTabelaPreco = objMnemonicoFPTipo.iTabelaPreco
        objMnemonicoFPTipo1.sMnemonico = GridItens.TextMatrix(iLinha, iGrid_Mnemonico_Col)
        objMnemonicoFPTipo1.sExpressao = GridItens.TextMatrix(iLinha, iGrid_Expressao_Col)
        objMnemonicoFPTipo1.sMnemonicoDesc = GridItens.TextMatrix(iLinha, iGrid_MnemonicoDesc_Col)
        objMnemonicoFPTipo1.iTipoPlanilha = Codigo_Extrai(TipoPlanilha.Text)
            
        colMnemonicoFPTipo.Add objMnemonicoFPTipo1
            
    Next
    
    'Grava o modelo padr�o de contabiliza��o em quest�o
    lErro = CF("MnemonicoFPTipo_Grava", colMnemonicoFPTipo)
    If lErro <> SUCESSO Then gError 92370
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
    
        Case 92360
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_NAO_INFORMADO1", gErr)
    
        Case 92361, 92364
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_INFORMADO", gErr)
    
        Case 92362, 92365, 92370
    
        Case 92363
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_NAO_PREENCHIDA", gErr)
        
        Case 92367
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MNEMONICO_NAO_PREENCHIDO", gErr, iLinha)

        Case 92368
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXPRESSAO_NAO_PREENCHIDA_GRID", gErr, iLinha)
        
        Case 92366
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRID_NAO_PREENCHIDO1", gErr)
        
         Case 116808
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPLANILHA_NAO_PREENCHIDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162742)
            
    End Select
    
    Exit Function
    
End Function

Public Sub BotaoExcluir_Click()
    
Dim lErro As Long
Dim objMnemonicoFPTipo As New ClassMnemonicoFPTipo
Dim vbMsgRes As VbMsgBoxResult
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
    
On Error GoTo Erro_BotaoExcluir_Click
     
    GL_objMDIForm.MousePointer = vbHourglass
    
    objMnemonicoFPTipo.iFilialEmpresa = giFilialEmpresa
    objMnemonicoFPTipo.iEscopo = iFrameAtual
    
    If objMnemonicoFPTipo.iEscopo = MNEMONICOFPRECO_ESCOPO_CATEGORIA Then
        
        If Len(ComboCategoriaProdutoItem.Text) = 0 Then gError 92371
        
        objMnemonicoFPTipo.sItemCategoria = ComboCategoriaProdutoItem.Text
        
    ElseIf objMnemonicoFPTipo.iEscopo = MNEMONICOFPRECO_ESCOPO_PRODUTO Then
    
        If Len(Trim(Produto.Text)) = 0 Then gError 92372
        
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 92269
        
        objMnemonicoFPTipo.sProduto = sProdutoFormatado
        
    ElseIf objMnemonicoFPTipo.iEscopo = MNEMONICOFPRECO_ESCOPO_TABPRECO Then
    
        If Len(Tabelapreco.Text) = 0 Then gError 92373
        If Len(Trim(Produto1.Text)) = 0 Then gError 92374
    
        objMnemonicoFPTipo.iTabelaPreco = Codigo_Extrai(Tabelapreco.Text)
        
        lErro = CF("Produto_Formata", Produto1.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 92375
        
        objMnemonicoFPTipo.sProduto = sProdutoFormatado
    
    End If
    
    objMnemonicoFPTipo.iTipoPlanilha = Codigo_Extrai(TipoPlanilha.Text)
     
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_MNEMONICOFPRECO")
    
    If vbMsgRes = vbYes Then
    
        'exclui o modelo padr�o de contabiliza��o em quest�o
        lErro = CF("MnemonicoFPTipo_Exclui", objMnemonicoFPTipo)
        If lErro <> SUCESSO Then gError 92376
    
        Call Limpa_Tela_MnemonicoFPreco
        
        iAlterado = 0
        
    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
        
        Case 92371
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_NAO_INFORMADO1", gErr)
    
        Case 92372, 92374
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_INFORMADO", gErr)
    
        Case 92373
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_NAO_PREENCHIDA", gErr)
        
        Case 92375, 92376
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162743)
        
    End Select

    Exit Sub
    
End Sub

Function Limpa_Tela_MnemonicoFPreco() As Long

    Call Limpa_Tela(Me)

    Tabelapreco.ListIndex = -1
    ComboCategoriaProdutoItem.ListIndex = -1
    
    'Limpa a combo tipo Planilha
    'TipoPlanilha.ListIndex = -1
    
    LabelDescricao.Caption = ""
    LabelDescricao1.Caption = ""
    
    Call Grid_Limpa(objGrid)

    objGrid.iLinhasExistentes = 0
    
    Limpa_Tela_MnemonicoFPreco = SUCESSO
    
End Function

Public Sub BotaoLimpar_Click()

Dim dtData As Date
Dim objPeriodo As New ClassPeriodo
Dim lDoc As Long
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 92377

    Call Limpa_Tela_MnemonicoFPreco

    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 92377
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162744)
        
    End Select
    
End Sub

Public Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub Produto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Produto_Validate
    
    lErro = CF("Produto_Perde_Foco", Produto, LabelDescricao)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 92378
    
    If lErro <> SUCESSO Then gError 92379
    
    'Atualiza Grid
    'Call TipoPlanilha_Click

    Exit Sub

Erro_Produto_Validate:

    Cancel = True

    Select Case gErr

        Case 92378

        Case 92379
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
          
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162745)

    End Select

    Exit Sub

End Sub

Private Sub Produto1_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Produto1_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Produto1_Validate
    
    lErro = CF("Produto_Perde_Foco", Produto1, LabelDescricao1)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 92380
    
    If lErro <> SUCESSO Then gError 92381
    
    'Atualiza Grid
    'Call TipoPlanilha_Click

    Exit Sub

Erro_Produto1_Validate:

    Cancel = True

    Select Case gErr

        Case 92380

        Case 92381
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
          
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162746)

    End Select

    Exit Sub

End Sub

Private Function Traz_MnemonicoFPreco_Tela(colMnemonicoFPTipo As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objMnemonicoFPTipo As ClassMnemonicoFPTipo
Dim objTipoPlanilha As New ClassTipoPlanilha

On Error GoTo Erro_Traz_MnemonicoFPreco_Tela

    'limpa o grid de express�es
    Call Grid_Limpa(objGrid)

    If colMnemonicoFPTipo.Count <> 0 Then
    
        Set objMnemonicoFPTipo = colMnemonicoFPTipo.Item(1)
        
        Select Case objMnemonicoFPTipo.iEscopo
        
            Case MNEMONICOFPRECO_ESCOPO_GERAL
                EscopoGeral.Value = True
            
            Case MNEMONICOFPRECO_ESCOPO_CATEGORIA
                EscopoCategoria.Value = True
                ComboCategoriaProdutoItem.Text = objMnemonicoFPTipo.sItemCategoria
                
            Case MNEMONICOFPRECO_ESCOPO_PRODUTO
                
                EscopoProduto.Value = True
                
                lErro = CF("Traz_Produto_MaskEd", objMnemonicoFPTipo.sProduto, Produto, LabelDescricao)
                If lErro <> SUCESSO Then gError 92382
    
            Case MNEMONICOFPRECO_ESCOPO_TABPRECO
    
                EscopoTabela.Value = True
    
                For iIndice = 0 To Tabelapreco.ListCount - 1
                    If Tabelapreco.ItemData(iIndice) = objMnemonicoFPTipo.iTabelaPreco Then
                        Tabelapreco.ListIndex = iIndice
                        Exit For
                    End If
                Next
                
                lErro = CF("Traz_Produto_MaskEd", objMnemonicoFPTipo.sProduto, Produto1, LabelDescricao1)
                If lErro <> SUCESSO Then gError 92383
    
        End Select
    
         'preenche objTipoPlanilha com o Codigo do tipo de planilha
        objTipoPlanilha.iCodigo = objMnemonicoFPTipo.iTipoPlanilha
        
        lErro = CF("TipoPlanilha_Le", objTipoPlanilha)
        If lErro <> SUCESSO Then gError 116809
        
        If Codigo_Extrai(TipoPlanilha.Text) <> objTipoPlanilha.iCodigo Then
        
            TipoPlanilha.Text = objTipoPlanilha.iCodigo & SEPARADOR & objTipoPlanilha.sDescricao
        
        End If
        
        iIndice = 0
    
        For Each objMnemonicoFPTipo In colMnemonicoFPTipo
        
            If objMnemonicoFPTipo.iFuncao = 0 Then
            
                iIndice = iIndice + 1
                
                GridItens.TextMatrix(iIndice, iGrid_Mnemonico_Col) = objMnemonicoFPTipo.sMnemonico
                GridItens.TextMatrix(iIndice, iGrid_MnemonicoDesc_Col) = objMnemonicoFPTipo.sMnemonicoDesc
                GridItens.TextMatrix(iIndice, iGrid_Expressao_Col) = objMnemonicoFPTipo.sExpressao
        
            End If
        
        Next

    End If
    
    objGrid.iLinhasExistentes = iIndice

    Traz_MnemonicoFPreco_Tela = SUCESSO
    
    Exit Function

Erro_Traz_MnemonicoFPreco_Tela:

    Traz_MnemonicoFPreco_Tela = gErr

    Select Case gErr

        Case 92382, 92383, 116809
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162747)

    End Select
    
    Exit Function

End Function

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim objCampoValor As AdmCampoValor
Dim objMnemonicoFPTipo As New ClassMnemonicoFPTipo
Dim lErro As Long

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada � Tela
    sTabela = "MnemonicoFPTipo"

    lErro = Move_Tela_Memoria(objMnemonicoFPTipo)
    If lErro <> SUCESSO Then gError 92384

    'Preenche a cole��o colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "FilialEmpresa", giFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "Escopo", objMnemonicoFPTipo.iEscopo, 0, "Escopo"
    colCampoValor.Add "ItemCategoria", objMnemonicoFPTipo.sItemCategoria, STRING_CATEGORIAPRODUTOITEM_ITEM, "ItemCategoria"
    colCampoValor.Add "Produto", objMnemonicoFPTipo.sProduto, STRING_PRODUTO, "Produto"
    colCampoValor.Add "TabelaPreco", objMnemonicoFPTipo.iTabelaPreco, 0, "TabelaPreco"
    colCampoValor.Add "TipoPlanilha", objMnemonicoFPTipo.iTipoPlanilha, 0, "TipoPlanilha"
    
    Exit Sub
    
Erro_Tela_Extrai:

    Select Case gErr
    
        Case 92384
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162748)

    End Select

    Exit Sub
    
End Sub

Private Function Move_Tela_Memoria(objMnemonicoFPTipo As ClassMnemonicoFPTipo) As Long

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer

On Error GoTo Erro_Move_Tela_Memoria

    If EscopoGeral.Value = True Then
        objMnemonicoFPTipo.iEscopo = MNEMONICOFPRECO_ESCOPO_GERAL
    ElseIf EscopoCategoria.Value = True Then
        objMnemonicoFPTipo.iEscopo = MNEMONICOFPRECO_ESCOPO_CATEGORIA
        objMnemonicoFPTipo.sItemCategoria = ComboCategoriaProdutoItem.Text
    ElseIf EscopoProduto.Value = True Then
        objMnemonicoFPTipo.iEscopo = MNEMONICOFPRECO_ESCOPO_PRODUTO
        
        lErro = CF("Produto_Formata", Produto.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 92384
        
        If iPreenchido = PRODUTO_PREENCHIDO Then objMnemonicoFPTipo.sProduto = sProduto
        
    ElseIf EscopoTabela.Value = True Then
        
        objMnemonicoFPTipo.iEscopo = MNEMONICOFPRECO_ESCOPO_TABPRECO
        
        If Tabelapreco.ListIndex <> -1 Then objMnemonicoFPTipo.iTabelaPreco = Tabelapreco.ItemData(Tabelapreco.ListIndex)
        
        lErro = CF("Produto_Formata", Produto1.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 92386
        
        If iPreenchido = PRODUTO_PREENCHIDO Then objMnemonicoFPTipo.sProduto = sProduto
        
    End If

    objMnemonicoFPTipo.iTipoPlanilha = Codigo_Extrai(TipoPlanilha.Text)
    
    'Preenche o codigo da Filial Empresa
    objMnemonicoFPTipo.iFilialEmpresa = giFilialEmpresa

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 92385, 92386

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162749)

    End Select

    Exit Function

End Function

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objMnemonicoFPTipo As New ClassMnemonicoFPTipo
Dim colMnemonicoFPTipo As New Collection

On Error GoTo Erro_Tela_Preenche

    objMnemonicoFPTipo.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
    objMnemonicoFPTipo.iEscopo = colCampoValor.Item("Escopo").vValor
    objMnemonicoFPTipo.sItemCategoria = colCampoValor.Item("ItemCategoria").vValor
    objMnemonicoFPTipo.sProduto = colCampoValor.Item("Produto").vValor
    objMnemonicoFPTipo.iTabelaPreco = colCampoValor.Item("TabelaPreco").vValor
    objMnemonicoFPTipo.iTipoPlanilha = colCampoValor.Item("TipoPlanilha").vValor

    'L� o Produto
    lErro = CF("MnemonicoFPTipo_Le_Todos", objMnemonicoFPTipo, colMnemonicoFPTipo)
    If lErro <> SUCESSO And lErro <> 116812 Then gError 92387

    lErro = Traz_MnemonicoFPreco_Tela(colMnemonicoFPTipo)
    If lErro <> SUCESSO Then gError 92388
        
    iAlterado = 0
    
    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr
    
        Case 92387, 92388, 116812

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162750)

    End Select

    Exit Sub

End Sub

Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Private Sub GridItens_GotFocus()

    Call Grid_Recebe_Foco(objGrid)

End Sub

Private Sub GridItens_EnterCell()

    Call Grid_Entrada_Celula(objGrid, iAlterado)

End Sub

Private Sub GridItens_LeaveCell()

    Call Saida_Celula(objGrid)

End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGrid)

End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Private Sub GridItens_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGrid)
    
End Sub

Private Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGrid)

End Sub

Private Sub GridItens_Scroll()

    Call Grid_Scroll(objGrid)

End Sub

Private Sub Mnemonico_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Mnemonico_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Mnemonico_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Mnemonico_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Mnemonico
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MnemonicoDesc_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MnemonicoDesc_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub MnemonicoDesc_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub MnemonicoDesc_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = MnemonicoDesc
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Expressao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Expressao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Expressao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Expressao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Expressao
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub TipoPlanilha_Click()

Dim objMnemonicoFPlanilha As ClassMnemonicoFPTipo
Dim lErro As Long

On Error GoTo Erro_TipoPlanilha_Click
    
    Call Grid_Limpa(objGrid)
        
    Set objMnemonicoFPlanilha = New ClassMnemonicoFPTipo

    'preenche Objeto com dados que est�o na tela
    lErro = Move_Tela_Memoria(objMnemonicoFPlanilha)
    If lErro <> SUCESSO Then gError 116814
        
    'chama fun��o de Browser, que vai buscar Mnemonicos no BD e preenche-los na tela
    Call objEventoMnemonicoFPTipo_evSelecao(objMnemonicoFPlanilha)
       
    Exit Sub
    
Erro_TipoPlanilha_Click:

    Select Case gErr
    
        Case 116814

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162751)

    End Select

    Exit Sub
     
End Sub

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
        
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Produto Then
            Call LabelProduto_Click
        ElseIf Me.ActiveControl Is Produto1 Then
            Call LabelProduto1_Click
        End If
    
    End If
    
End Sub

Public Function Form_Load_Ocx() As Object

'    Parent.HelpContextID = IDH_PLANO_CONTAS
    Set Form_Load_Ocx = Me
    Caption = "Mnem�nicos para Planilhas de Forma��o de Pre�os"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "MnemonicoFPPlanilha"
    
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

Private Sub Unload(objme As Object)
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

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label8(Index), Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8(Index), Button, Shift, X, Y)
End Sub

Private Sub LabelCategoria_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCategoria, Source, X, Y)
End Sub

Private Sub LabelCategoria_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCategoria, Button, Shift, X, Y)
End Sub

Private Sub LabelDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDescricao, Source, X, Y)
End Sub

Private Sub LabelDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDescricao, Button, Shift, X, Y)
End Sub

Private Sub LabelDescricao1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDescricao1, Source, X, Y)
End Sub

Private Sub LabelDescricao1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDescricao1, Button, Shift, X, Y)
End Sub

Private Sub LabelProduto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProduto, Source, X, Y)
End Sub

Private Sub LabelProduto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProduto, Button, Shift, X, Y)
End Sub

Private Sub LabelProduto1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProduto1, Source, X, Y)
End Sub

Private Sub LabelProduto1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProduto1, Button, Shift, X, Y)
End Sub


