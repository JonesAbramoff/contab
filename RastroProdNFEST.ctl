VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl RastroProdNFEST 
   ClientHeight    =   5775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7785
   KeyPreview      =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   7785
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
      Left            =   5940
      TabIndex        =   21
      Top             =   5340
      Width           =   1665
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5940
      ScaleHeight     =   495
      ScaleWidth      =   1560
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   135
      Width           =   1620
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1065
         Picture         =   "RastroProdNFEST.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   570
         Picture         =   "RastroProdNFEST.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   75
         Picture         =   "RastroProdNFEST.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Rastreamento do Produto"
      Height          =   2760
      Left            =   150
      TabIndex        =   1
      Top             =   2460
      Width           =   7455
      Begin MSMask.MaskEdBox Lote 
         Height          =   255
         Left            =   1305
         TabIndex        =   2
         Top             =   285
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox LoteData 
         Height          =   255
         Left            =   4065
         TabIndex        =   4
         Top             =   300
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
         Left            =   2460
         TabIndex        =   3
         Top             =   300
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
         Left            =   5250
         TabIndex        =   5
         Top             =   315
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
         Height          =   1860
         Left            =   195
         TabIndex        =   6
         Top             =   315
         Width           =   7020
         _ExtentX        =   12383
         _ExtentY        =   3281
         _Version        =   393216
         Rows            =   51
         Cols            =   7
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin VB.Label QuantTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4260
         TabIndex        =   23
         Top             =   2340
         Width           =   1470
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
         Left            =   3675
         TabIndex        =   22
         Top             =   2370
         Width           =   510
      End
   End
   Begin VB.ComboBox Item 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   675
      Width           =   750
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
      Left            =   180
      TabIndex        =   20
      Top             =   2025
      Width           =   1050
   End
   Begin VB.Label Quantidade 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1305
      TabIndex        =   19
      Top             =   1995
      Width           =   1410
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
      Left            =   75
      TabIndex        =   18
      Top             =   1560
      Width           =   1155
   End
   Begin VB.Label Almoxarifado 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1305
      TabIndex        =   17
      Top             =   1545
      Width           =   1410
   End
   Begin VB.Label UnidadeMedida 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5070
      TabIndex        =   16
      Top             =   1995
      Width           =   1440
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
      Left            =   4230
      TabIndex        =   15
      Top             =   2025
      Width           =   780
   End
   Begin VB.Label Descricao 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2790
      TabIndex        =   14
      Top             =   1110
      Width           =   4335
   End
   Begin VB.Label Produto 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1305
      TabIndex        =   13
      Top             =   1110
      Width           =   1410
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
      Left            =   510
      TabIndex        =   12
      Top             =   1125
      Width           =   735
   End
   Begin VB.Label Label1 
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
      Left            =   795
      TabIndex        =   11
      Top             =   735
      Width           =   435
   End
End
Attribute VB_Name = "RastroProdNFEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit




'IDHs para Help
Const IDH_RASTROPRODNFEST = 0

'Variáveis globais
Dim iAlterado As Integer
Dim gcolItemNF As Collection
Dim gobjGenerico As AdmGenerico
Dim giItemNF As Integer

'GridRastro
Dim objGridRastro As AdmGrid
Dim iGrid_Lote_Col As Integer
Dim iGrid_LoteData_Col As Integer
Dim iGrid_QuantLote_Col As Integer
Dim iGrid_FilialOP_Col As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

'Browses
Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'testa se houva alguma alteração
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 83038

    'Limpa a Tela
    Call Limpa_Tela_Rastreamento

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 83038

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166033)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    QuantLote.Format = FORMATO_ESTOQUE

    Set objGridRastro = New AdmGrid
    Set objEventoLote = New AdmEvento
    
    'Inicializa o grid de Rastreamento
    lErro = Inicializa_Grid_Rastreamento(objGridRastro)
    If lErro <> SUCESSO Then gError 75788

    giItemNF = -1

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 75788

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166034)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, UnloadMode, Cancel, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    'Libera as variaveis globais
    Set gcolItemNF = Nothing
    Set objGridRastro = Nothing
    
    Set objEventoLote = Nothing
    
    'Só desabilita a tela se foi passado o obj como parâmetro
    If Not gobjGenerico Is Nothing Then
        gobjGenerico.vVariavel = HABILITA_TELA
    End If
    
End Sub

Private Sub BotaoLotes_Click()

Dim lErro As Long
Dim objRastroLote As New ClassRastreamentoLote
Dim colSelecao As New Collection
Dim sSelecao As String
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoLotes_Click
    
    'Se o produto não foi preenchido, erro
    If Len(Trim(Produto.Caption)) = 0 Then gError 75890
    
    'Verifica se tem alguma linha selecionada no Grid
    If GridRastro.Row = 0 Then gError 75886
        
    'Formata o produto
    lErro = CF("Produto_Formata",Produto.Caption, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 75887
    
    'Lê o produto
    objProduto.sCodigo = sProdutoFormatado
    lErro = CF("Produto_Le",objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 75888
    
    'Produto não cadastrado
    If lErro = 28030 Then gError 75889
        
    'Verifica o tipo de rastreamento do produto
    If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
        sSelecao = " FilialOP = ? AND Produto = ?"
    ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
        sSelecao = " FilialOP <> ? AND Produto = ?"
    End If
    
    'Adiciona filtros
    colSelecao.Add 0
    colSelecao.Add sProdutoFormatado
    
    'Chama a tela de browse RastroLoteLista passando como parâmetro a seleção do Filtro (sSelecao)
    Call Chama_Tela("RastroLoteLista", colSelecao, objRastroLote, objEventoLote, sSelecao)
                    
    Exit Sub

Erro_BotaoLotes_Click:

    Select Case gErr
        
        Case 75886
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 75887, 75888
        
        Case 75889
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
                    
        Case 75890
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166035)
    
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

Private Sub objEventoLote_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRastroLote As New ClassRastreamentoLote
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iLinha As Integer

On Error GoTo Erro_objEventoLote_evSelecao

    Set objRastroLote = obj1

    'Se a Linha corrente for diferente da Linha fixa
    If GridRastro.Row <> 0 Then

        'Formata o Produto para o BD
        lErro = CF("Produto_Formata",Produto.Caption, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 83054
            
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le",objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 83055
            
        'Se o produto não está cadastrado, erro
        If lErro = 28030 Then gError 83056
                
        'Se o Produto foi preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
            'Se o produto possuir rastro por lote
            If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
                
                For iLinha = 1 To objGridRastro.iLinhasExistentes
                    If iLinha <> objGridRastro.objGrid.Row Then
                        If objGridRastro.objGrid.TextMatrix(iLinha, iGrid_Lote_Col) = objRastroLote.sCodigo Then gError 83057
                    End If
                Next
        
            'Se o produto possuir rastro por OP
            ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
                                                
                For iLinha = 1 To objGridRastro.iLinhasExistentes
                    If iLinha <> objGridRastro.objGrid.Row Then
                        If objGridRastro.objGrid.TextMatrix(iLinha, iGrid_Lote_Col) = objRastroLote.sCodigo And Codigo_Extrai(objGridRastro.objGrid.TextMatrix(iLinha, iGrid_FilialOP_Col)) = objRastroLote.iFilialOP Then gError 83058
                    End If
                Next
                
            End If

        End If

        'Coloca o Lote na tela
        GridRastro.TextMatrix(GridRastro.Row, iGrid_Lote_Col) = objRastroLote.sCodigo
        Lote.Text = objRastroLote.sCodigo
        
        'Lê lote e preenche dados
        lErro = Lote_Saida_Celula(objRastroLote)
        If lErro <> SUCESSO Then gError 75907
        
    End If

    Me.Show

    Exit Sub

Erro_objEventoLote_evSelecao:

    Select Case gErr

        Case 75907, 83054, 83055
        
        Case 83056
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 83057
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_JA_UTILIZADO_GRID", gErr, objRastroLote.sCodigo)
            
        Case 83058
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_FILIALOP_JA_UTILIZADO_GRID", gErr, Lote.Text, objRastroLote.iFilialOP)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 166036)

    End Select

    Exit Sub

End Sub

Private Sub Item_Click()

Dim lErro As Long
Dim objItemNF As ClassItemNF
Dim iIndice As Integer

On Error GoTo Erro_Item_Click

    'testa se houva alguma alteração
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 71984

    'Verifica se tem um item selecionado
    If Item.ListIndex = -1 Then Exit Sub

    If CInt(Item.Text) = giItemNF Then Exit Sub

    'Guarda o ItemNF escolhido
    Set objItemNF = gcolItemNF.Item(CInt(Item.Text))

    'Limpa o Grid
    Call Grid_Limpa(objGridRastro)

    'Preenche Grid de Rastreamento
    lErro = Preenche_Tela(objItemNF)
    If lErro <> SUCESSO Then gError 75882
    
    giItemNF = CInt(Item.Text)
    
    iAlterado = 0

    Exit Sub

Erro_Item_Click:

    Select Case gErr
        
        Case 71984
            For iIndice = 0 To Item.ListCount - 1
                If CInt(Item.List(iIndice)) = giItemNF Then
                    Item.ListIndex = iIndice
                    Exit For
                End If
            Next
        
        Case 75882
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166037)

    End Select

    Exit Sub

End Sub

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
    
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

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

    Exit Function

End Function

Public Function Trata_Parametros(colItemNF As Collection, Optional objGenerico As AdmGenerico) As Long

Dim objItemNF As ClassItemNF, lErro As Long
Dim objProduto As New ClassProduto

On Error GoTo Erro_Trata_Parametros

    Set gcolItemNF = colItemNF
    Set gobjGenerico = objGenerico
    
    'Carrega combo de Itens
    For Each objItemNF In colItemNF

        'Lê o produto
        objProduto.sCodigo = objItemNF.sProduto
        lErro = CF("Produto_Le",objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 75789

        'Se o produto não está cadastrado, erro
        If lErro = 28030 Then gError 75790

        'Se o produto possui Rastreamento por lote
        If objProduto.iRastro <> PRODUTO_RASTRO_NENHUM Then
            Item.AddItem objItemNF.iItem
        End If

    Next

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 75789

        Case 75790
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166038)

    End Select

    Exit Function

End Function

Private Function Preenche_Tela(objItemNF As ClassItemNF) As Long

Dim sProdutoMascarado As String
Dim objRastroItemNF As ClassRastroItemNF
Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim objProduto As New ClassProduto
Dim dQuantTotal As Double

On Error GoTo Erro_Preenche_Tela

    'Mascara o produto
    lErro = Mascara_MascararProduto(objItemNF.sProduto, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 75791

    Produto.Caption = sProdutoMascarado
    Descricao.Caption = objItemNF.sDescricaoItem
    UnidadeMedida.Caption = objItemNF.sUnidadeMed
    Almoxarifado.Caption = objItemNF.sAlmoxarifadoNomeRed
    Quantidade.Caption = Format(objItemNF.dQuantidade, FORMATO_ESTOQUE)
    
    objProduto.sCodigo = objItemNF.sProduto
            
    'Lê os demais atributos do Produto
    lErro = CF("Produto_Le",objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 83027
        
    'Se o produto não está cadastrado, erro
    If lErro = 28030 Then gError 83028
                
    If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
        FilialOP.Enabled = False
    Else
        FilialOP.Enabled = True
    End If
    
'    Set objGridRastro = New AdmGrid
'
'    'Inicializa o grid de Rastreamento
'    lErro = Inicializa_Grid_Rastreamento(objGridRastro)
'    If lErro <> SUCESSO Then gError 83029
    
    'Para cada Produto,Almoxarifado,Lote
    For Each objRastroItemNF In objItemNF.colRastreamento

        GridRastro.TextMatrix(objGridRastro.iLinhasExistentes + 1, iGrid_Lote_Col) = objRastroItemNF.sLote

        If objRastroItemNF.dtLoteData <> DATA_NULA Then
            GridRastro.TextMatrix(objGridRastro.iLinhasExistentes + 1, iGrid_LoteData_Col) = Format(objRastroItemNF.dtLoteData, "dd/mm/yyyy")
        End If

        If objRastroItemNF.iLoteFilialOP <> 0 Then

            'Lê FilialEmpresa
            objFilialEmpresa.iCodFilial = objRastroItemNF.iLoteFilialOP
            lErro = CF("FilialEmpresa_Le",objFilialEmpresa)
            If lErro <> SUCESSO And lErro <> 27378 Then gError 75696

            'Se não encontrou a FilialEmpresa
            If lErro = 27378 Then gError 75697

            GridRastro.TextMatrix(objGridRastro.iLinhasExistentes + 1, iGrid_FilialOP_Col) = CStr(objFilialEmpresa.iCodFilial) & SEPARADOR & objFilialEmpresa.sNome

        End If

        dQuantTotal = dQuantTotal + objRastroItemNF.dLoteQdtAlocada

        GridRastro.TextMatrix(objGridRastro.iLinhasExistentes + 1, iGrid_QuantLote_Col) = Formata_Estoque(objRastroItemNF.dLoteQdtAlocada)

        'Incrementa o número de linhas existentes no Grid
        objGridRastro.iLinhasExistentes = objGridRastro.iLinhasExistentes + 1

    Next

    QuantTotal.Caption = Formata_Estoque(dQuantTotal)

    Preenche_Tela = SUCESSO

    Exit Function

Erro_Preenche_Tela:

    Preenche_Tela = gErr
    
    Select Case gErr

        Case 75791, 75696, 83027, 83029

        Case 75697
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)
        
        Case 83028
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, Produto.Caption)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166039)

    End Select

    Exit Function

End Function

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
                If lErro <> SUCESSO Then gError 75858

            'FilialOP
            Case iGrid_FilialOP_Col
                lErro = Saida_Celula_FilialOP(objGridInt)
                If lErro <> SUCESSO Then gError 75859
            
            'Quantidade
            Case iGrid_QuantLote_Col
                lErro = Saida_Celula_QuantLote(objGridInt)
                If lErro <> SUCESSO Then gError 75860
        
        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 75861

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 75858, 75859, 75860

        Case 75861
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166040)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantLote(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_QuantLote

    Set objGridInt.objControle = QuantLote

    'Se a quantidade alocada do lote foi preenchida
    If Len(Trim(QuantLote.ClipText)) > 0 Then

        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(QuantLote.Text)
        If lErro <> SUCESSO Then gError 75862

        'Se a quantidade alocada do lote for maior que a quantidade disponível, erro
        If StrParaDbl(QuantLote.Text) > StrParaDbl(Quantidade.Caption) Then gError 75863

    End If

    'totaliza as quantidades dos lotes e mostra no campo QuantTotal
    QuantTotal.Caption = Format(GridQuantLote_Soma() + StrParaDbl(QuantLote.Text), FORMATO_ESTOQUE)

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 75864

    Saida_Celula_QuantLote = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantLote:

    Saida_Celula_QuantLote = gErr

    Select Case gErr

        Case 75862, 75864
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 75863
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTLOTE_MAIOR_QUANTALM", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166041)

    End Select

    Exit Function

End Function

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
        If Len(Trim(Produto.Caption)) = 0 Then gError 75885
        
        If Item.ListIndex = -1 Then gError 83040
        
        'Formata o Produto para o BD
        lErro = CF("Produto_Formata",Produto.Caption, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 75865
            
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        objProduto.sCodigo = sProdutoFormatado
        lErro = CF("Produto_Le",objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 75866
            
        'Se o produto não está cadastrado, erro
        If lErro = 28030 Then gError 75867
                
        'Se o Produto foi preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
            objRastroLote.dtDataEntrada = DATA_NULA
            
            'Se o produto possuir rastro por lote
            If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
                
                For iLinha = 1 To objGridInt.iLinhasExistentes
                    If iLinha <> objGridInt.objGrid.Row Then
                        If objGridInt.objGrid.TextMatrix(iLinha, iGrid_Lote_Col) = Lote.Text Then gError 83048
                    End If
                Next
                
                objRastroLote.sCodigo = Lote.Text
                objRastroLote.sProduto = sProdutoFormatado
                
                'Lê o Rastreamento do Lote vinculado ao produto
                lErro = CF("RastreamentoLote_Le",objRastroLote)
                If lErro <> SUCESSO And lErro <> 75710 Then gError 75868
                
                'Se não encontrou --> Erro
                If lErro = 75710 Then gError 75869
                
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
                                gError 83039
                            End If
                            Exit For
                        End If
                    End If
                Next
                
                If objRastroLote.iFilialOP <> 0 Then
                
                    'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                    lErro = CF("RastreamentoLote_Le",objRastroLote)
                    If lErro <> SUCESSO And lErro <> 75710 Then gError 75870
                
                    'Se não encontrou --> Erro
                    If lErro = 75710 Then gError 75871
                
                End If
                
            End If
        
        End If
    
        'Preenche campos do lote
        lErro = Lote_Saida_Celula(objRastroLote)
        If lErro <> SUCESSO Then gError 75891
    
    Else
    
        GridRastro.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteData_Col) = ""
        
    End If
                                    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 71875

    Saida_Celula_Lote = SUCESSO

    Exit Function

Erro_Saida_Celula_Lote:

    Saida_Celula_Lote = gErr

    Select Case gErr

        Case 71875, 75865, 75866, 75868, 75870, 75885, 75891
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 75867
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 75869
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("RastreamentoLote", objRastroLote)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
        
        Case 75871
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OP_RASTREAMENTO_INEXISTENTE", gErr, objRastroLote.sCodigo, objRastroLote.iFilialOP, objRastroLote.sProduto)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83039
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_FILIALOP_JA_UTILIZADO_GRID", gErr, Lote.Text, objRastroLote.iFilialOP)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83040
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMNF_NAO_SELECIONADO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83048
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_JA_UTILIZADO_GRID", gErr, Lote.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166042)

    End Select

    Exit Function

End Function

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
        
    'Se a filial foi preenchida
    If Len(Trim(FilialOP.Text)) > 0 Then
        
        If Item.ListIndex = -1 Then gError 83041
        
        'Valida a Filial
        lErro = TP_FilialEmpresa_Le(FilialOP.Text, objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 71971 And lErro <> 71972 Then gError 83030

        'Se não for encontrado --> Erro
        If lErro = 71971 Then gError 83031
        If lErro = 71972 Then gError 83032
        
        If Len(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Lote_Col)) <> 0 Then
        
            For iLinha = 1 To objGridInt.iLinhasExistentes
                If iLinha <> objGridInt.objGrid.Row Then
                    If objGridInt.objGrid.TextMatrix(iLinha, iGrid_Lote_Col) = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Lote_Col) And Codigo_Extrai(objGridInt.objGrid.TextMatrix(iLinha, iGrid_FilialOP_Col)) = objFilialEmpresa.iCodFilial Then gError 83033
                End If
            Next
        
            'Formata o produto
            lErro = CF("Produto_Formata",Produto.Caption, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 83034
        
            objRastroLote.sCodigo = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Lote_Col)
            objRastroLote.sProduto = sProdutoFormatado
            objRastroLote.iFilialOP = objFilialEmpresa.iCodFilial
                
            'Lê o Rastreamento do Lote vinculado ao produto
            lErro = CF("RastreamentoLote_Le",objRastroLote)
            If lErro <> SUCESSO And lErro <> 75710 Then gError 83035
                
            If objRastroLote.dtDataEntrada <> DATA_NULA And lErro = SUCESSO Then
                GridRastro.TextMatrix(GridRastro.Row, iGrid_LoteData_Col) = Format(objRastroLote.dtDataEntrada, "dd/mm/yyyy")
            End If
            
            FilialOP.Text = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome
            
        Else
            
            FilialOP.Text = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome
            
        End If
        
    Else
    
        GridRastro.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteData_Col) = ""
        
    End If
                                    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 83037

    Saida_Celula_FilialOP = SUCESSO

    Exit Function

Erro_Saida_Celula_FilialOP:

    Saida_Celula_FilialOP = gErr

    Select Case gErr

        Case 83030, 83034, 83035, 83037
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83031, 83032
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, FilialOP.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 83033
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_FILIALOP_JA_UTILIZADO_GRID", gErr, objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Lote_Col), objFilialEmpresa.iCodFilial)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 83041
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMNF_NAO_SELECIONADO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166043)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Se nenhum item foi selecionado ==> erro
    If giItemNF = -1 Then gError 83039

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 75874

    iAlterado = 0

    'Limpa a Tela
    Call Limpa_Tela_Rastreamento

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 75874

        Case 83039
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMNF_NAO_SELECIONADO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 166044)

    End Select

    Exit Sub

End Sub

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
        lErro = CF("FilialEmpresa_Le",objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 75892

        'Se não encontrou a FilialEmpresa
        If lErro = 27378 Then gError 75893

        GridRastro.TextMatrix(GridRastro.Row, iGrid_FilialOP_Col) = CStr(objFilialEmpresa.iCodFilial) & SEPARADOR & objFilialEmpresa.sNome
    
    End If
    
    If objGridRastro.objGrid.Row - objGridRastro.objGrid.FixedRows = objGridRastro.iLinhasExistentes Then
        objGridRastro.iLinhasExistentes = objGridRastro.iLinhasExistentes + 1
    End If
    
    Lote_Saida_Celula = SUCESSO
    
    Exit Function
        
Erro_Lote_Saida_Celula:

    Lote_Saida_Celula = gErr
    
    Select Case gErr
        
        Case 75892
        
        Case 75893
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 166045)
    
    End Select
    
    Exit Function
    
End Function

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

    If giItemNF <> -1 Then

        GL_objMDIForm.MousePointer = vbHourglass
    
        'Formata o Produto para o BD
        lErro = CF("Produto_Formata",Produto.Caption, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 83042
            
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le",objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 83043
            
        'Se o produto não está cadastrado, erro
        If lErro = 28030 Then gError 83044
    
        'Para cada linha do Grid
        For iIndice = 1 To objGridRastro.iLinhasExistentes
        
            'Se o lote não foi preenchido, erro
            If Len(Trim(GridRastro.TextMatrix(iIndice, iGrid_Lote_Col))) = 0 Then gError 75879
            
            'Se o produto possuir rastro por lote
            If objProduto.iRastro = PRODUTO_RASTRO_OP And Len(Trim(GridRastro.TextMatrix(iIndice, iGrid_FilialOP_Col))) = 0 Then gError 83045
    
            'Se a quantidade não foi preenchida,erro
            If Len(Trim(GridRastro.TextMatrix(iIndice, iGrid_QuantLote_Col))) = 0 Then gError 83047
            
            'Se a quantidade está zerada, erro
            If StrParaDbl(GridRastro.TextMatrix(iIndice, iGrid_QuantLote_Col)) = 0 Then gError 83048
            
            dQuantLoteTotal = dQuantLoteTotal + StrParaDbl(GridRastro.TextMatrix(iIndice, iGrid_QuantLote_Col))
    
        Next
    
        'Se a quantidade alocada do Lote foi maior que a quantidade alocada no almoxarifado
        If dQuantLoteTotal > StrParaDbl(Quantidade.Caption) Then gError 75875
    
        Set objItemNF = gcolItemNF.Item(giItemNF)
    
        'Move dados da tela para a memória
        lErro = Move_Tela_Memoria(objItemNF)
        If lErro <> SUCESSO Then gError 75876
    
        GL_objMDIForm.MousePointer = vbDefault
    
        giRetornoTela = vbOK
    
    End If
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 75875
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTTOTAL_LOTE_MAIOR_ALMOXARIFADO", gErr, dQuantLoteTotal, StrParaDbl(Quantidade.Caption))

        Case 75876, 83042, 83043
        
        Case 75879
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRID_LOTE_NAO_PREENCHIDO", gErr, iIndice)
        
        Case 83044
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 83045
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_OP_NAO_PREENCHIDA", gErr, iIndice)
        
        Case 83047
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRID_QUANTLOTE_NAO_PREENCHIDA", gErr, iIndice)
        
        Case 83048
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRID_QUANTLOTE_ZERADA", gErr, iIndice)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166046)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(objItemNF As ClassItemNF) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim iIndice As Integer
Dim objRastroItemNF As ClassRastroItemNF
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_Move_Tela_Memoria
    
    Set objItemNF.colRastreamento = New Collection
    
    'Para cada linha do Grid de Rastreamento
    For iLinha = 1 To objGridRastro.iLinhasExistentes
        
        'Cria novo item de rastreamento
        Set objRastroItemNF = New ClassRastroItemNF
        objRastroItemNF.dAlmoxQtdAlocada = StrParaDbl(Quantidade.Caption)
        objRastroItemNF.dLoteQdtAlocada = StrParaDbl(GridRastro.TextMatrix(iLinha, iGrid_QuantLote_Col))
        objRastroItemNF.dtLoteData = StrParaDate(GridRastro.TextMatrix(iLinha, iGrid_LoteData_Col))
        objRastroItemNF.sAlmoxNomeRed = Almoxarifado.Caption
        objRastroItemNF.sLote = GridRastro.TextMatrix(iLinha, iGrid_Lote_Col)
        
        'Lê almoxarifado
        objAlmoxarifado.sNomeReduzido = Almoxarifado.Caption
        lErro = CF("Almoxarifado_Le_NomeReduzido",objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25060 Then gError 75878
        
        'Se não encontrou o almoxarifado, erro
        If lErro = 25060 Then gError 75879
        
        objRastroItemNF.iAlmoxCodigo = objAlmoxarifado.iCodigo
        objRastroItemNF.iLoteFilialOP = Codigo_Extrai(GridRastro.TextMatrix(iLinha, iGrid_FilialOP_Col))
                    
        'Adiciona na coleção
        objItemNF.colRastreamento.Add objRastroItemNF
        
    Next
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 75878, 83046

        Case 75879
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166047)

    End Select

    Exit Function

End Function

Private Function GridQuantLote_Soma() As Double
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

Private Sub Limpa_Tela_Rastreamento()

    giItemNF = -1
    Item.ListIndex = -1

    'Limpa o Grid de Rastreamento
    Call Grid_Limpa(objGridRastro)
    
    Produto.Caption = ""
    Descricao.Caption = ""
    UnidadeMedida.Caption = ""
    Almoxarifado.Caption = ""
    Quantidade.Caption = ""
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166048)

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


'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RASTROPRODNFEST
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Lote Then
            Call BotaoLotes_Click
        End If

    End If

End Sub


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

Private Sub QuantTotal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantTotal, Source, X, Y)
End Sub

Private Sub QuantTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantTotal, Button, Shift, X, Y)
End Sub

Private Sub LabelTotal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTotal, Source, X, Y)
End Sub

Private Sub LabelTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTotal, Button, Shift, X, Y)
End Sub

'******* Copiada de outras telas *******************************************
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 166049)

    End Select

    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'***************************************************************************

