VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl RastroEstoqueInicial 
   ClientHeight    =   5805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8025
   KeyPreview      =   -1  'True
   ScaleHeight     =   5805
   ScaleWidth      =   8025
   Begin VB.CommandButton BotaoSerie 
      Caption         =   "Séries"
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
      Left            =   240
      TabIndex        =   24
      Top             =   5280
      Width           =   1665
   End
   Begin VB.ComboBox Escaninho 
      Height          =   315
      ItemData        =   "RastroEstoqueInicial.ctx":0000
      Left            =   1440
      List            =   "RastroEstoqueInicial.ctx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1350
      Width           =   6330
   End
   Begin VB.Frame Frame7 
      Caption         =   "Rastreamento do Produto"
      Height          =   2760
      Left            =   225
      TabIndex        =   5
      Top             =   2445
      Width           =   7560
      Begin MSMask.MaskEdBox Lote 
         Height          =   255
         Left            =   750
         TabIndex        =   6
         Top             =   285
         Width           =   2000
         _ExtentX        =   3519
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox LoteData 
         Height          =   255
         Left            =   3540
         TabIndex        =   7
         Top             =   225
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox FilialOP 
         Height          =   225
         Left            =   1905
         TabIndex        =   8
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox QuantLote 
         Height          =   225
         Left            =   4710
         TabIndex        =   9
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridRastro 
         Height          =   1935
         Left            =   90
         TabIndex        =   10
         Top             =   315
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   3413
         _Version        =   393216
         Rows            =   51
         Cols            =   7
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin VB.Label LabelTotal 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
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
         Height          =   195
         Left            =   3495
         TabIndex        =   23
         Top             =   2355
         Width           =   510
      End
      Begin VB.Label QuantTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4080
         TabIndex        =   22
         Top             =   2310
         Width           =   1470
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6150
      ScaleHeight     =   495
      ScaleWidth      =   1560
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   105
      Width           =   1620
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   75
         Picture         =   "RastroEstoqueInicial.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   570
         Picture         =   "RastroEstoqueInicial.ctx":015E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1065
         Picture         =   "RastroEstoqueInicial.ctx":0690
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoLotes 
      Caption         =   "Lotes"
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
      Left            =   6105
      TabIndex        =   0
      Top             =   5325
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Escaninho:"
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
      Left            =   405
      TabIndex        =   21
      Top             =   1395
      Width           =   960
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   630
      TabIndex        =   20
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Produto 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   19
      Top             =   810
      Width           =   1410
   End
   Begin VB.Label Descricao 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2925
      TabIndex        =   18
      Top             =   810
      Width           =   4860
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Unidade:"
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
      Left            =   4035
      TabIndex        =   17
      Top             =   1965
      Width           =   780
   End
   Begin VB.Label UnidadeMedida 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4875
      TabIndex        =   16
      Top             =   1920
      Width           =   1440
   End
   Begin VB.Label Almoxarifado 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1470
      TabIndex        =   15
      Top             =   330
      Width           =   1425
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   225
      TabIndex        =   14
      Top             =   345
      Width           =   1155
   End
   Begin VB.Label Quantidade 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   13
      Top             =   1920
      Width           =   1410
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Quantidade:"
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
      Height          =   195
      Left            =   315
      TabIndex        =   12
      Top             =   1950
      Width           =   1050
   End
End
Attribute VB_Name = "RastroEstoqueInicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'#####################################
'Inserido por Wagner 03/04/2006
'Acrescentado botão Série no Layout e alterado o tabIndex
'#####################################

'Variáveis globais
Dim iAlterado As Integer
Dim gobjTelaEstoqueInicial As Object
Dim gobjGenerico As AdmGenerico

'GridRastro
Dim objGridRastro As AdmGrid
Dim iGrid_Lote_Col As Integer
Dim iGrid_LoteData_Col As Integer
Dim iGrid_QuantLote_Col As Integer
Dim iGrid_FilialOP_Col As Integer
Dim giEscaninho As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

'Browses
Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

'm
Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'testa se houva alguma alteração
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 71886

    'Limpa a Tela
    Call Limpa_Tela_Rastreamento

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 71886

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165920)

    End Select

    Exit Sub

End Sub

'm
Public Sub Form_Load()

Dim lErro As Long
Dim colEscaninhos As New Collection
Dim objEscaninho As ClassEscaninho

On Error GoTo Erro_Form_Load

    QuantLote.Format = FORMATO_ESTOQUE

    Set objGridRastro = New AdmGrid
    Set objEventoLote = New AdmEvento
    
    'Le todos os escaninhos que podem ter rastreamento do estoque inicial
    lErro = Escaninhos_Le_EstoqueInicial(colEscaninhos)
    If lErro <> SUCESSO Then gError 71835

    For Each objEscaninho In colEscaninhos
    
        Escaninho.AddItem objEscaninho.sNome
        Escaninho.ItemData(Escaninho.NewIndex) = objEscaninho.iCodigo
    
    Next

    giEscaninho = -1

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 71835, 71850

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165921)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

'm
Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, UnloadMode, Cancel, iTelaCorrenteAtiva)

End Sub

'm
Public Sub Form_UnLoad(Cancel As Integer)
    
    'Libera as variaveis globais
    Set objGridRastro = Nothing
    
    Set objEventoLote = Nothing
    
    'Só desabilita a tela se foi passado o obj como parâmetro
    If Not gobjGenerico Is Nothing Then
        gobjGenerico.vVariavel = HABILITA_TELA
    End If
    
End Sub

'm
Private Sub BotaoLotes_Click()

Dim lErro As Long
Dim objRastroLote As New ClassRastreamentoLote
Dim colSelecao As New Collection
Dim sSelecao As String
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoLotes_Click
    
    If Escaninho.ListIndex = -1 Then gError 71958
    
    'Se o produto não foi preenchido, erro
    If Len(Trim(Produto.Caption)) = 0 Then gError 71851
    
    'Verifica se tem alguma linha selecionada no Grid
    If GridRastro.Row = 0 Then gError 71852
        
    'Formata o produto
    lErro = CF("Produto_Formata", Produto.Caption, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 71853
    
    'Lê o produto
    objProduto.sCodigo = sProdutoFormatado
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 71854
    
    'Produto não cadastrado
    If lErro = 28030 Then gError 71855
        
    'Verifica o tipo de rastreamento do produto
    If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
        sSelecao = " FilialOP = ? AND Produto = ?"
    ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
        sSelecao = " FilialOP <> ? AND Produto = ?"
    ElseIf objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
        sSelecao = " FilialOP = ? AND Produto = ?"
    End If
    
    'Adiciona filtros
    colSelecao.Add 0
    colSelecao.Add sProdutoFormatado
    
    'Chama a tela de browse RastroLoteLista passando como parâmetro a seleção do Filtro (sSelecao)
    Call Chama_Tela("RastroLoteLista", colSelecao, objRastroLote, objEventoLote, sSelecao)
                    
    Exit Sub

Erro_BotaoLotes_Click:

    Select Case gErr
        
        Case 71851
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case 71852
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 71853, 71854
        
        Case 71855
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case 71958
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESCANINHO_NAO_SELECIONADO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165922)
    
    End Select
    
    Exit Sub

End Sub

Private Sub FilialOP_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialOP_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRastro)

End Sub

Private Sub FilialOP_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRastro)

End Sub

Private Sub FilialOP_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRastro.objControle = FilialOP
    lErro = Grid_Campo_Libera_Foco(objGridRastro)
    If lErro <> SUCESSO Then Cancel = True

End Sub

'm
Private Sub objEventoLote_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRastroLote As New ClassRastreamentoLote

On Error GoTo Erro_objEventoLote_evSelecao

    Set objRastroLote = obj1

    'Se a Linha corrente for diferente da Linha fixa
    If GridRastro.Row <> 0 Then

        'Coloca o Lote na tela
        GridRastro.TextMatrix(GridRastro.Row, iGrid_Lote_Col) = objRastroLote.sCodigo
        Lote.Text = objRastroLote.sCodigo
        
        'Lê lote e preenche dados
        lErro = Lote_Saida_Celula(objRastroLote)
        If lErro <> SUCESSO Then gError 71856
        
    End If

    Me.Show

    Exit Sub

Erro_objEventoLote_evSelecao:

    Select Case gErr

        Case 71856
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165923)

    End Select

    Exit Sub

End Sub

'm
Private Sub Escaninho_Click()

Dim lErro As Long
Dim objItemNF As ClassItemNF
Dim objRastroEstIni As ClassRastroEstIni
Dim colRastreamento As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Escaninho_Click

    'testa se houva alguma alteração
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 71984

    'Verifica se não tem um item selecionado ==> não faz nada
    If Escaninho.ListIndex = -1 Then Exit Sub

    If giEscaninho = Escaninho.ItemData(Escaninho.ListIndex) Then Exit Sub

    'Limpa o Grid
    Call Grid_Limpa(objGridRastro)

    'seleciona os rastreamentos do escaninho escolhido
    For Each objRastroEstIni In gobjTelaEstoqueInicial.gcolRastreamento
      
        If objRastroEstIni.iEscaninho = Escaninho.ItemData(Escaninho.ListIndex) Then
            colRastreamento.Add objRastroEstIni
        End If
    Next

    giEscaninho = Escaninho.ItemData(Escaninho.ListIndex)

    'Preenche Grid de Rastreamento
    lErro = Preenche_Tela(colRastreamento)
    If lErro <> SUCESSO Then gError 71840
    
    iAlterado = 0

    Exit Sub

Erro_Escaninho_Click:

    Select Case gErr
        
        Case 71840
        
        Case 71984
            For iIndice = 0 To Escaninho.ListCount - 1
                If Escaninho.ItemData(iIndice) = giEscaninho Then
                    Escaninho.ListIndex = iIndice
                    Exit For
                End If
            Next
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165924)

    End Select

    Exit Sub

End Sub

'm
Private Function Inicializa_Grid_Rastreamento(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Alocação

Dim iIndice As Integer

    Set objGridRastro.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Lote")
    objGridInt.colColuna.Add ("FilialOP do Lote")
    objGridInt.colColuna.Add ("Data do Lote")
    objGridInt.colColuna.Add ("Qtd. Alocada Lote")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Lote.Name)
    objGridInt.colCampo.Add (FilialOP.Name)
    objGridInt.colCampo.Add (LoteData.Name)
    objGridInt.colCampo.Add (QuantLote.Name)

    'Colunas da Grid
    iGrid_Lote_Col = 1
    iGrid_FilialOP_Col = 2
    iGrid_LoteData_Col = 3
    iGrid_QuantLote_Col = 4

    'Grid do GridInterno
    objGridInt.objGrid = GridRastro

    'Largura da primeira coluna
    GridRastro.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    'Linhas visíveis
    objGridInt.iLinhasVisiveis = 6

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridRastro)

    'Posiciona os painéis totalizadores
    QuantTotal.Top = objGridInt.objGrid.Top + objGridInt.objGrid.Height
    QuantTotal.Left = objGridInt.objGrid.Left
    For iIndice = 0 To iGrid_QuantLote_Col - 1
        QuantTotal.Left = QuantTotal.Left + objGridInt.objGrid.ColWidth(iIndice) + objGridInt.objGrid.GridLineWidth + 20
    Next
    
    QuantTotal.Width = objGridInt.objGrid.ColWidth(iGrid_QuantLote_Col)
    
    LabelTotal.Top = QuantTotal.Top + (QuantTotal.Height - LabelTotal.Height) / 2
    LabelTotal.Left = QuantTotal.Left - LabelTotal.Width

    Inicializa_Grid_Rastreamento = SUCESSO

End Function

'm
Public Function Trata_Parametros(objTelaEstoqueInicial As Object, Optional objGenerico As AdmGenerico) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Trata_Parametros

    Set gobjGenerico = objGenerico

    'Formata o Produto para o BD
    lErro = CF("Produto_Formata", objTelaEstoqueInicial.Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 71966
        
    objProduto.sCodigo = sProdutoFormatado
            
    'Lê os demais atributos do Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 71967
        
    'Se o produto não está cadastrado, erro
    If lErro = 28030 Then gError 71968
                
    If objProduto.iRastro = PRODUTO_RASTRO_LOTE Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
        FilialOP.Enabled = False
    Else
        FilialOP.Enabled = True
    End If
    
    'Inicializa o grid de Rastreamento
    lErro = Inicializa_Grid_Rastreamento(objGridRastro)
    If lErro <> SUCESSO Then gError 71969
                
    Produto.Caption = objTelaEstoqueInicial.Produto.Text
    Descricao.Caption = objTelaEstoqueInicial.DescricaoProduto.Caption
    UnidadeMedida.Caption = objTelaEstoqueInicial.UnidMed.Caption
    Almoxarifado.Caption = objTelaEstoqueInicial.Almoxarifado.Text
    
    Set gobjTelaEstoqueInicial = objTelaEstoqueInicial
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 71966, 71967, 71969

        Case 71968
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, Produto.Caption)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165925)

    End Select

    Exit Function

End Function

'm
Private Function Preenche_Tela(colRastreamento As Collection) As Long
'preenche o grid com os rastreamentos passados como parametro

Dim objRastroEstIni As ClassRastroEstIni
Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim dQuantTotal As Double

On Error GoTo Erro_Preenche_Tela

    Call Preenche_Campo_Quantidade
    
    'Colocar cada rastreamento no grid
    For Each objRastroEstIni In colRastreamento

        GridRastro.TextMatrix(objGridRastro.iLinhasExistentes + 1, iGrid_Lote_Col) = objRastroEstIni.sLote

        If objRastroEstIni.dtDataEntrada <> DATA_NULA Then
            GridRastro.TextMatrix(objGridRastro.iLinhasExistentes + 1, iGrid_LoteData_Col) = Format(objRastroEstIni.dtDataEntrada, "dd/mm/yyyy")
        End If

        If objRastroEstIni.iFilialOP <> 0 Then

            'Lê FilialEmpresa
            objFilialEmpresa.iCodFilial = objRastroEstIni.iFilialOP
            lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
            If lErro <> SUCESSO And lErro <> 27378 Then gError 71856

            'Se não encontrou a FilialEmpresa
            If lErro = 27378 Then gError 71857

            GridRastro.TextMatrix(objGridRastro.iLinhasExistentes + 1, iGrid_FilialOP_Col) = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome

        End If

        GridRastro.TextMatrix(objGridRastro.iLinhasExistentes + 1, iGrid_QuantLote_Col) = Formata_Estoque(objRastroEstIni.dQuantidade)

        dQuantTotal = dQuantTotal + objRastroEstIni.dQuantidade

        'Incrementa o número de linhas existentes no Grid
        objGridRastro.iLinhasExistentes = objGridRastro.iLinhasExistentes + 1

    Next

    QuantTotal.Caption = Formata_Estoque(dQuantTotal)

    Preenche_Tela = SUCESSO

    Exit Function

Erro_Preenche_Tela:

    Preenche_Tela = Err

    Select Case gErr

        Case 71856

        Case 71857
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165926)

    End Select

    Exit Function

End Function

'm
Private Sub Preenche_Campo_Quantidade()
'preenche o campo quantidade com o valor contido no escaninho selecionado na tela de estoque inicial

Dim sProdutoMascarado As String
Dim objRastroItemNF As ClassRastroItemNF
Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Preenche_Campo_Quantidade

    Select Case giEscaninho
    
        Case ESCANINHO_DISPONIVEL
            Quantidade.Caption = Format(gobjTelaEstoqueInicial.Quantidade, FORMATO_ESTOQUE)
            
        Case ESCANINHO_CONSERTO_NOSSO
            Quantidade.Caption = Format(gobjTelaEstoqueInicial.QuantConserto, FORMATO_ESTOQUE)
        
        Case ESCANINHO_CONSIG_NOSSO
            Quantidade.Caption = Format(gobjTelaEstoqueInicial.QuantConsig, FORMATO_ESTOQUE)

        Case ESCANINHO_DEMO_NOSSO
            Quantidade.Caption = Format(gobjTelaEstoqueInicial.QuantDemo, FORMATO_ESTOQUE)

        Case ESCANINHO_OUTROS_NOSSO
            Quantidade.Caption = Format(gobjTelaEstoqueInicial.QuantOutras, FORMATO_ESTOQUE)

        Case ESCANINHO_BENEF_NOSSO
            Quantidade.Caption = Format(gobjTelaEstoqueInicial.QuantBenef, FORMATO_ESTOQUE)

        Case ESCANINHO_CONSERTO_3
            Quantidade.Caption = Format(gobjTelaEstoqueInicial.QuantConserto3, FORMATO_ESTOQUE)
        
        Case ESCANINHO_CONSIG_3
            Quantidade.Caption = Format(gobjTelaEstoqueInicial.QuantConsig3, FORMATO_ESTOQUE)

        Case ESCANINHO_DEMO_3
            Quantidade.Caption = Format(gobjTelaEstoqueInicial.QuantDemo3, FORMATO_ESTOQUE)

        Case ESCANINHO_OUTROS_3
            Quantidade.Caption = Format(gobjTelaEstoqueInicial.QuantOutras3, FORMATO_ESTOQUE)

        Case ESCANINHO_BENEF_3
            Quantidade.Caption = Format(gobjTelaEstoqueInicial.QuantBenef3, FORMATO_ESTOQUE)

    End Select

    Exit Sub

Erro_Preenche_Campo_Quantidade:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165927)

    End Select

    Exit Sub

End Sub

'm
Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    If lErro = SUCESSO Then

        'Verifica qual a coluna do Grid em questão
        Select Case objGridInt.objGrid.Col

            'Lote
            Case iGrid_Lote_Col
                lErro = Saida_Celula_Lote(objGridInt)
                If lErro <> SUCESSO Then gError 71858
            
            'FilialOP
            Case iGrid_FilialOP_Col
                lErro = Saida_Celula_FilialOP(objGridInt)
                If lErro <> SUCESSO Then gError 71859
            
            'Quantidade
            Case iGrid_QuantLote_Col
                lErro = Saida_Celula_QuantLote(objGridInt)
                If lErro <> SUCESSO Then gError 71860
        
        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 71861

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 71858, 71859, 71860

        Case 71861
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165928)

    End Select

    Exit Function

End Function

'm
Private Function Saida_Celula_QuantLote(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_QuantLote

    Set objGridInt.objControle = QuantLote

    'Se a quantidade alocada do lote foi preenchida
    If Len(Trim(QuantLote.ClipText)) > 0 Then

        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(QuantLote.Text)
        If lErro <> SUCESSO Then gError 71862

        'Se a quantidade alocada do lote for maior que a quantidade disponível, erro
        If StrParaDbl(QuantLote.Text) > StrParaDbl(Quantidade.Caption) Then gError 71863

        '############################################################
        'Inserido por Wagner 06/04/2006
        lErro = Valida_Serie(GridRastro.Row, StrParaDbl(QuantLote.Text), GridRastro.TextMatrix(GridRastro.Row, iGrid_Lote_Col))
        If lErro <> SUCESSO Then gError 177315
        '############################################################
       
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 71864

    'totaliza as quantidades dos lotes e mostra no campo QuantTotal
    QuantTotal.Caption = Format(GridQuantLote_Soma() + StrParaDbl(QuantLote.Text), FORMATO_ESTOQUE)

    Saida_Celula_QuantLote = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantLote:

    Saida_Celula_QuantLote = gErr

    Select Case gErr

        Case 71862, 71864, 177315
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 71863
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTLOTE_MAIOR_QUANTALM", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165929)

    End Select

    Exit Function

End Function

'm
Private Function Saida_Celula_Lote(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objRastroLote As New ClassRastreamentoLote
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim iLinha As Integer

On Error GoTo Erro_Saida_Celula_Lote

    Set objGridInt.objControle = Lote
        
    'Se o lote foi preenchido
    If Len(Trim(Lote.Text)) > 0 Then
        
        'Se o produto não está preenchido, sai da rotina
        If Len(Trim(Produto.Caption)) = 0 Then gError 71865
        
        If Escaninho.ListIndex = -1 Then gError 71961
        
        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", Produto.Caption, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 71866
            
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 71867
            
        'Se o produto não está cadastrado, erro
        If lErro = 28030 Then gError 71868
                
        'Se o Produto foi preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
            'Se o produto possuir rastro por lote
            If objProduto.iRastro = PRODUTO_RASTRO_LOTE Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
                
                For iLinha = 1 To objGridInt.iLinhasExistentes
                    If iLinha <> objGridInt.objGrid.Row Then
                        If objGridInt.objGrid.TextMatrix(iLinha, iGrid_Lote_Col) = Lote.Text Then gError 71983
                    End If
                Next
        
                objRastroLote.sCodigo = Lote.Text
                objRastroLote.sProduto = sProdutoFormatado
                
                'Lê o Rastreamento do Lote vinculado ao produto
                lErro = CF("RastreamentoLote_Le", objRastroLote)
                If lErro <> SUCESSO And lErro <> 75710 Then gError 71869
                
                'Se não encontrou --> Erro
                If lErro = 75710 Then gError 71870
                
            'Se o produto possuir rastro por OP
            ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
                                                
                objRastroLote.sCodigo = Lote.Text
                objRastroLote.sProduto = sProdutoFormatado
                If Len(GridRastro.TextMatrix(objGridInt.objGrid.Row, iGrid_FilialOP_Col)) = 0 Then
                    objRastroLote.iFilialOP = giFilialEmpresa
                Else
                    objRastroLote.iFilialOP = Codigo_Extrai(GridRastro.TextMatrix(objGridInt.objGrid.Row, iGrid_FilialOP_Col))
                End If
                
                For iLinha = 1 To objGridInt.iLinhasExistentes
                    If iLinha <> objGridInt.objGrid.Row Then
                        If objGridInt.objGrid.TextMatrix(iLinha, iGrid_Lote_Col) = Lote.Text And Codigo_Extrai(objGridInt.objGrid.TextMatrix(iLinha, iGrid_FilialOP_Col)) = objRastroLote.iFilialOP Then
                            If Len(GridRastro.TextMatrix(objGridInt.objGrid.Row, iGrid_FilialOP_Col)) = 0 Then
                                objRastroLote.iFilialOP = 0
                                objRastroLote.dtDataEntrada = DATA_NULA
                            Else
                                gError 83016
                            End If
                            Exit For
                        End If
                    End If
                Next
                
                If objRastroLote.iFilialOP <> 0 Then
                    'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                    lErro = CF("RastreamentoLote_Le", objRastroLote)
                    If lErro <> SUCESSO And lErro <> 75710 Then gError 71871
                
                    'Se não encontrou --> Erro
                    If lErro = 75710 Then gError 71872
                
                End If
                
            End If
        
        End If
    
    
        'Preenche campos do lote
        lErro = Lote_Saida_Celula(objRastroLote)
        If lErro <> SUCESSO Then gError 71873
        
    Else
    
        GridRastro.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteData_Col) = ""
        
    End If
                                    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 71874

    Saida_Celula_Lote = SUCESSO

    Exit Function

Erro_Saida_Celula_Lote:

    Saida_Celula_Lote = gErr

    Select Case gErr

        Case 71865, 71866, 71867, 71869, 71871, 71873, 71874
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 71868
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 71870
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("RastreamentoLote", objRastroLote)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
        
        Case 71872
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RASTREAMENTOOP_NAO_CADASTRADO", gErr, objRastroLote.sProduto, objRastroLote.sCodigo, objRastroLote.iFilialOP)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 71961
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESCANINHO_NAO_SELECIONADO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 71983
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_JA_UTILIZADO_GRID", gErr, Lote.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 83016
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_FILIALOP_JA_UTILIZADO_GRID", gErr, Lote.Text, objRastroLote.iFilialOP)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165930)

    End Select

    Exit Function

End Function

'm
Private Function Saida_Celula_FilialOP(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objRastroLote As New ClassRastreamentoLote
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim objFilialEmpresa As New AdmFiliais
Dim iLinha As Integer

On Error GoTo Erro_Saida_Celula_FilialOP

    Set objGridInt.objControle = FilialOP
        
    'Se a Filial foi preenchida
    If Len(Trim(FilialOP.Text)) > 0 Then
        
        If Escaninho.ListIndex = -1 Then gError 71979
        
        'Valida a Filial
        lErro = TP_FilialEmpresa_Le(FilialOP.Text, objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 71971 And lErro <> 71972 Then gError 71976

        'Se não for encontrado --> Erro
        If lErro = 71971 Then gError 71977
        If lErro = 71972 Then gError 71978
        
        If Len(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Lote_Col)) <> 0 Then
        
            For iLinha = 1 To objGridInt.iLinhasExistentes
                If iLinha <> objGridInt.objGrid.Row Then
                    If objGridInt.objGrid.TextMatrix(iLinha, iGrid_Lote_Col) = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Lote_Col) And Codigo_Extrai(objGridInt.objGrid.TextMatrix(iLinha, iGrid_FilialOP_Col)) = objFilialEmpresa.iCodFilial Then gError 83017
                End If
            Next
        
            'Formata o produto
            lErro = CF("Produto_Formata", Produto.Caption, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 83015
        
            objRastroLote.sCodigo = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Lote_Col)
            objRastroLote.sProduto = sProdutoFormatado
            objRastroLote.iFilialOP = objFilialEmpresa.iCodFilial
                
            'Lê o Rastreamento do Lote vinculado ao produto
            lErro = CF("RastreamentoLote_Le", objRastroLote)
            If lErro <> SUCESSO And lErro <> 75710 Then gError 71980
                
            'Se não encontrou --> Erro
            If lErro = 75710 Then gError 71981
        
            If objRastroLote.dtDataEntrada <> DATA_NULA And lErro = SUCESSO Then
                GridRastro.TextMatrix(GridRastro.Row, iGrid_LoteData_Col) = Format(objRastroLote.dtDataEntrada, "dd/mm/yyyy")
            Else
                GridRastro.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteData_Col) = ""
            End If
            
            FilialOP.Text = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome
            
        Else
            
            FilialOP.Text = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome
            
        End If
        
    Else
    
        GridRastro.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteData_Col) = ""
        
    End If
                                    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 83036

    Saida_Celula_FilialOP = SUCESSO

    Exit Function

Erro_Saida_Celula_FilialOP:

    Saida_Celula_FilialOP = gErr

    Select Case gErr

        Case 71976, 71980, 71982, 83015, 83036
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 71977
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA1", gErr, FilialOP.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 71978
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, FilialOP.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 71979
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESCANINHO_NAO_SELECIONADO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 71981
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RASTREAMENTOOP_NAO_CADASTRADO", gErr, objRastroLote.sProduto, objRastroLote.sCodigo, objRastroLote.iFilialOP)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 83017
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_FILIALOP_JA_UTILIZADO_GRID", gErr, objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Lote_Col), objFilialEmpresa.iCodFilial)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165931)

    End Select

    Exit Function

End Function

'm
Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Se nenhum item foi selecionado, sai da rotina
    If Escaninho.ListIndex = -1 Or giEscaninho = -1 Then gError 71957
    
    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 71878

    iAlterado = 0

    'Limpa a Tela
    Call Limpa_Tela_Rastreamento

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 71878

        Case 71957
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESCANINHO_NAO_SELECIONADO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165932)

    End Select

    Exit Sub

End Sub

'm
Function Lote_Saida_Celula(objRastroLote As ClassRastreamentoLote) As Long
'Executa a saida de celula do campo lote, o tratamento dos erros do Grid é feita na rotina chamadora

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Lote_Saida_Celula

    If objRastroLote.dtDataEntrada <> DATA_NULA Then
        GridRastro.TextMatrix(GridRastro.Row, iGrid_LoteData_Col) = Format(objRastroLote.dtDataEntrada, "dd/mm/yyyy")
    End If
    
    'Se a filial empresa foi preenchida
    If objRastroLote.iFilialOP <> 0 Then
        
        objFilialEmpresa.iCodFilial = objRastroLote.iFilialOP
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 71879

        'Se não encontrou a FilialEmpresa
        If lErro = 27378 Then gError 71880

        GridRastro.TextMatrix(GridRastro.Row, iGrid_FilialOP_Col) = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome
    
    End If
    
    '############################################################
    'Inserido por Wagner 06/04/2006
    lErro = Valida_Serie(GridRastro.Row, StrParaDbl(GridRastro.TextMatrix(GridRastro.Row, iGrid_QuantLote_Col)), Lote.Text)
    If lErro <> SUCESSO Then gError 177316
    '############################################################
    
    If objGridRastro.objGrid.Row - objGridRastro.objGrid.FixedRows = objGridRastro.iLinhasExistentes Then
        objGridRastro.iLinhasExistentes = objGridRastro.iLinhasExistentes + 1
    End If
    
    Lote_Saida_Celula = SUCESSO
    
    Exit Function
        
Erro_Lote_Saida_Celula:

    Lote_Saida_Celula = gErr
    
    Select Case gErr
        
        Case 71879, 177316
        
        Case 71880
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165933)
    
    End Select
    
    Exit Function
    
End Function

'm
Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objItemNF As ClassItemNF
Dim iIndice As Integer
Dim dQuantLoteTotal As Double
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Gravar_Registro

    If giEscaninho <> -1 Then
        
        GL_objMDIForm.MousePointer = vbHourglass
    
        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", Produto.Caption, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 71962
            
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 71963
            
        'Se o produto não está cadastrado, erro
        If lErro = 28030 Then gError 71964
                    
        'Para cada linha do Grid
        For iIndice = 1 To objGridRastro.iLinhasExistentes
        
            'Se o lote não foi preenchido, erro
            If Len(Trim(GridRastro.TextMatrix(iIndice, iGrid_Lote_Col))) = 0 Then gError 71881
            
            'Se o produto possuir rastro por lote
            If objProduto.iRastro = PRODUTO_RASTRO_OP And Len(Trim(GridRastro.TextMatrix(iIndice, iGrid_FilialOP_Col))) = 0 Then gError 71965
    
            'Se a quantidade não foi preenchida,erro
            If Len(Trim(GridRastro.TextMatrix(iIndice, iGrid_QuantLote_Col))) = 0 Then gError 71959
            
            'Se a quantidade está zerada, erro
            If StrParaDbl(GridRastro.TextMatrix(iIndice, iGrid_QuantLote_Col)) = 0 Then gError 71960
            
            dQuantLoteTotal = dQuantLoteTotal + StrParaDbl(GridRastro.TextMatrix(iIndice, iGrid_QuantLote_Col))
    
        Next
    
        'Se a quantidade alocada do Lote foi maior que a quantidade alocada no almoxarifado
        If dQuantLoteTotal > StrParaDbl(Quantidade.Caption) Then gError 71883
    
        'Move dados da tela para a memória
        lErro = Move_Tela_Memoria()
        If lErro <> SUCESSO Then gError 71884
    
        gobjTelaEstoqueInicial.iAlterado = REGISTRO_ALTERADO
        GL_objMDIForm.MousePointer = vbDefault
    
        giRetornoTela = vbOK
    
    End If
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 71881
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRID_LOTE_NAO_PREENCHIDO", gErr, iIndice)
        
        Case 71883
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTTOTAL_LOTE_MAIOR_ESCANINHO", gErr, dQuantLoteTotal, StrParaDbl(Quantidade.Caption))

        Case 71884, 71962, 71963
        
        Case 71959
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRID_QUANTLOTE_NAO_PREENCHIDA", gErr, iIndice)
        
        Case 71960
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRID_QUANTLOTE_ZERADA", gErr, iIndice)
        
        Case 71964
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 71965
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_OP_NAO_PREENCHIDA", gErr, iIndice)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165934)

    End Select

    Exit Function

End Function

'm
Private Function Move_Tela_Memoria() As Long

Dim lErro As Long
Dim iLinha As Integer
Dim iIndice As Integer
Dim objRastroEstIni As ClassRastroEstIni
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_Move_Tela_Memoria
    
    'seleciona os rastreamentos do escaninho escolhido
    For iIndice = gobjTelaEstoqueInicial.gcolRastreamento.Count To 1 Step -1
      
        If gobjTelaEstoqueInicial.gcolRastreamento.Item(iIndice).iEscaninho = giEscaninho Then
            gobjTelaEstoqueInicial.gcolRastreamento.Remove (iIndice)
        End If
    Next

    'Verifica se Produto está preenchido
    If Len(Trim(Produto.Caption)) > 0 Then

        'Critica o formato do Produto
        lErro = CF("Produto_Formata", Produto.Caption, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 71884
        
    End If

    'Verifica se Almoxarifado está preenchido
    If Almoxarifado.Caption <> "" Then
    
        'preenche o objAlmoxarifado
        objAlmoxarifado.sNomeReduzido = Almoxarifado.Caption

        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO Then gError 71885

    End If

    'Para cada linha do Grid de Rastreamento
    For iLinha = 1 To objGridRastro.iLinhasExistentes
        
        'Cria novo item de rastreamento
        Set objRastroEstIni = New ClassRastroEstIni
        objRastroEstIni.dQuantidade = StrParaDbl(GridRastro.TextMatrix(iLinha, iGrid_QuantLote_Col))
        objRastroEstIni.dtDataEntrada = StrParaDate(GridRastro.TextMatrix(iLinha, iGrid_LoteData_Col))
        objRastroEstIni.sLote = GridRastro.TextMatrix(iLinha, iGrid_Lote_Col)
        objRastroEstIni.iFilialOP = Codigo_Extrai(GridRastro.TextMatrix(iLinha, iGrid_FilialOP_Col))
        objRastroEstIni.iAlmoxarifado = objAlmoxarifado.iCodigo
        objRastroEstIni.sProduto = sProdutoFormatado
        objRastroEstIni.iEscaninho = giEscaninho
        
        gobjTelaEstoqueInicial.gcolRastreamento.Add objRastroEstIni
        
    Next
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 71884, 71885

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165935)

    End Select

    Exit Function

End Function

'm
Function GridQuantLote_Soma() As Double
'soma o conteudo da coluna QuantLote e retorna o total

Dim dAcumulador As Double
Dim iLinha As Integer

    dAcumulador = 0

    For iLinha = 1 To objGridRastro.iLinhasExistentes
        If Len(objGridRastro.objGrid.TextMatrix(iLinha, iGrid_QuantLote_Col)) > 0 And iLinha <> objGridRastro.objGrid.Row Then
            dAcumulador = dAcumulador + CDbl(objGridRastro.objGrid.TextMatrix(iLinha, iGrid_QuantLote_Col))
        End If
    Next

    GridQuantLote_Soma = dAcumulador

End Function


'm
Private Sub Limpa_Tela_Rastreamento()

    giEscaninho = -1
    Escaninho.ListIndex = -1

    'Limpa o Grid de Rastreamento
    Call Grid_Limpa(objGridRastro)
    
    QuantTotal.Caption = ""
    
End Sub

'Tratamento do Grid
Private Sub GridRastro_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridRastro, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRastro, iAlterado)
    End If

End Sub

Private Sub GridRastro_EnterCell()

    Call Grid_Entrada_Celula(objGridRastro, iAlterado)

End Sub

Private Sub GridRastro_GotFocus()

    Call Grid_Recebe_Foco(objGridRastro)

End Sub

Private Sub GridRastro_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridRastro, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRastro, iAlterado)
    End If

End Sub

Private Sub GridRastro_LeaveCell()

    Call Saida_Celula(objGridRastro)

End Sub

Private Sub GridRastro_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridRastro)

End Sub

Private Sub GridRastro_RowColChange()

    Call Grid_RowColChange(objGridRastro)

End Sub

Private Sub GridRastro_Scroll()

    Call Grid_Scroll(objGridRastro)

End Sub

Private Sub GridRastro_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long

On Error GoTo Erro_GridRastro_KeyDown

    Call Grid_Trata_Tecla1(KeyCode, objGridRastro)

    Exit Sub

Erro_GridRastro_KeyDown:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165936)

    End Select
    
    Exit Sub

End Sub

Private Sub QuantLote_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantLote_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRastro)

End Sub

Private Sub QuantLote_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRastro)

End Sub

Private Sub QuantLote_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRastro.objControle = QuantLote
    lErro = Grid_Campo_Libera_Foco(objGridRastro)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Lote_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Lote_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRastro)

End Sub

Private Sub Lote_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRastro)

End Sub

Private Sub Lote_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRastro.objControle = Lote
    lErro = Grid_Campo_Libera_Foco(objGridRastro)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub LoteData_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LoteData_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRastro)

End Sub

Private Sub LoteData_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRastro)

End Sub

Private Sub LoteData_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRastro.objControle = LoteData
    lErro = Grid_Campo_Libera_Foco(objGridRastro)
    If lErro <> SUCESSO Then Cancel = True

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RASTROESTOQUEINICIAL
    Set Form_Load_Ocx = Me
    Caption = "Rastreamento de Produto"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RastroProdNFEST"

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
        
        If Me.ActiveControl Is Lote Then
            Call BotaoLotes_Click
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

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub LabelTotal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTotal, Source, X, Y)
End Sub

Private Sub LabelTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTotal, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Produto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Produto, Source, X, Y)
End Sub

Private Sub Produto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Produto, Button, Shift, X, Y)
End Sub

Private Sub Descricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Descricao, Source, X, Y)
End Sub

Private Sub Descricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Descricao, Button, Shift, X, Y)
End Sub

Private Sub QuantTotal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantTotal, Source, X, Y)
End Sub

Private Sub QuantTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantTotal, Button, Shift, X, Y)
End Sub

Private Sub Almoxarifado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Almoxarifado, Source, X, Y)
End Sub

Private Sub Almoxarifado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Almoxarifado, Button, Shift, X, Y)
End Sub

Private Sub Quantidade_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Quantidade, Source, X, Y)
End Sub

Private Sub Quantidade_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Quantidade, Button, Shift, X, Y)
End Sub

'?????******* Copiada de outras telas *******************************************
'Copiada de RastreamentoLote
Function RastreamentoLote_Le(objRastroLote As ClassRastreamentoLote) As Long
'Lê rastreamento do lote a partir do produto, filialOP e código do lote passados

Dim lErro As Long
Dim lComando As Long
Dim tRastroLote As typeRastreamentoLote

On Error GoTo Erro_RastreamentoLote_Le

    'Abertura dos comandos
    lComando = Comando_Abrir()
    If lErro <> SUCESSO Then gError 75707

    tRastroLote.sObservacao = String(STRING_RASTRO_OBSERVACAO, 0)

    'Lê dados de RastrementoLote a partir de Produto, FilialOP e Lote
    lErro = Comando_Executar(lComando, "SELECT DataValidade, DataEntrada, DataFabricacao, Observacao FROM RastreamentoLote WHERE Produto = ? AND Lote = ? AND FilialOP = ?", tRastroLote.dtDataValidade, tRastroLote.dtDataEntrada, tRastroLote.dtDataFabricacao, tRastroLote.sObservacao, objRastroLote.sProduto, objRastroLote.sCodigo, objRastroLote.iFilialOP)
    If lErro <> AD_SQL_SUCESSO Then gError 75708

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 75709

    objRastroLote.dtDataEntrada = tRastroLote.dtDataEntrada
    objRastroLote.dtDataFabricacao = tRastroLote.dtDataFabricacao
    objRastroLote.dtDataValidade = tRastroLote.dtDataValidade
    objRastroLote.sObservacao = tRastroLote.sObservacao
    
    'Se não encontrou, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 75710

    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    RastreamentoLote_Le = SUCESSO

    Exit Function

Erro_RastreamentoLote_Le:

    RastreamentoLote_Le = gErr

    Select Case gErr

        Case 75707
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 75708, 75709
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOLOTE", gErr)

        Case 75710 'RastreamentoLote não cadastrado

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165937)

    End Select

    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'***************************************************************************

'????? Transferir para ClassMATSelect
Function Escaninhos_Le_EstoqueInicial(colEscaninhos As Collection) As Long
'Le todos os escaninhos que podem ter rastreamento do estoque inicial

Dim lErro As Long
Dim lComando As Long
Dim tEscaninho As typeEscaninho
Dim objEscaninho As ClassEscaninho

On Error GoTo Erro_Escaninhos_Le_EstoqueInicial

    'Abertura dos comandos
    lComando = Comando_Abrir()
    If lErro <> SUCESSO Then gError 71836

    tEscaninho.sNome = String(STRING_NOME_ESCANINHO, 0)

    'Lê os dados dos Escaninhos que podem ter rastreamento do estoque inicial
    lErro = Comando_Executar(lComando, "SELECT Codigo, Nome FROM Escaninhos WHERE RastroEstoqueInicial = ?", tEscaninho.iCodigo, tEscaninho.sNome, ESCANINHO_RASTRO_ESTOQUE_INICIAL)
    If lErro <> AD_SQL_SUCESSO Then gError 71837

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71838

    Do While lErro = AD_SQL_SUCESSO
    
        Set objEscaninho = New ClassEscaninho
    
        objEscaninho.iCodigo = tEscaninho.iCodigo
        objEscaninho.sNome = tEscaninho.sNome
        
        colEscaninhos.Add objEscaninho
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71839
    
    Loop
    
    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    Escaninhos_Le_EstoqueInicial = SUCESSO

    Exit Function

Erro_Escaninhos_Le_EstoqueInicial:

    Escaninhos_Le_EstoqueInicial = gErr

    Select Case gErr

        Case 71836
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 71837, 71838, 71839
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESCANINHOS", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165938)

    End Select

    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'###########################################################
'Inserido por Wagner 06/04/2006
Private Function Valida_Serie(ByVal iLinha As Integer, ByVal dQuantidade As Double, ByVal sSerie As String) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim objItemMovEst As New ClassItemMovEstoque
Dim objEstProd As New ClassEstoqueProduto
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_Valida_Serie

    lErro = CF("Produto_Formata", Produto.Caption, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 177317
        
    objProduto.sCodigo = sProdutoFormatado
            
    'Lê os demais atributos do Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 177318
    
    If objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
        
        objAlmoxarifado.sNomeReduzido = Almoxarifado.Caption

        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO Then gError 177319
        
        objItemMovEst.dQuantidade = dQuantidade
        objItemMovEst.dQuantidadeEst = dQuantidade
        objItemMovEst.iAlmoxarifado = objAlmoxarifado.iCodigo
        objItemMovEst.sProduto = objProduto.sCodigo
        objItemMovEst.sSiglaUM = objProduto.sSiglaUMEstoque
        objItemMovEst.sSiglaUMEst = objProduto.sSiglaUMEstoque
    
        Select Case giEscaninho
            
            Case ESCANINHO_DISPONIVEL
                objEstProd.dQuantDispNossa = dQuantidade
            Case ESCANINHO_CONSERTO_NOSSO
                objEstProd.dQuantConserto = dQuantidade
            Case ESCANINHO_CONSIG_NOSSO
                objEstProd.dQuantConsig = dQuantidade
            Case ESCANINHO_DEMO_NOSSO
                objEstProd.dQuantDemo = dQuantidade
            Case ESCANINHO_OUTROS_NOSSO
                objEstProd.dQuantOutras = dQuantidade
            Case ESCANINHO_BENEF_NOSSO
                objEstProd.dQuantBenef = dQuantidade
            Case ESCANINHO_CONSERTO_3
                objEstProd.dQuantConserto3 = dQuantidade
            Case ESCANINHO_CONSIG_3
                objEstProd.dQuantConsig3 = dQuantidade
            Case ESCANINHO_DEMO_3
                objEstProd.dQuantDemo3 = dQuantidade
            Case ESCANINHO_OUTROS_3
                objEstProd.dQuantOutras3 = dQuantidade
            Case ESCANINHO_BENEF_3
                objEstProd.dQuantBenef3 = dQuantidade
                
        End Select
            
        lErro = CF("RastreamentoSerie_Valida_Serie", objItemMovEst, objEstProd, sSerie)
        If lErro <> SUCESSO Then gError 177320
    
    End If
    
    Valida_Serie = SUCESSO
        
    Exit Function

Erro_Valida_Serie:

    Valida_Serie = gErr

    Select Case gErr
    
        Case 177317 To 177320
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177321)

    End Select

    Exit Function

End Function
'###########################################################

