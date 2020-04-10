VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RastroProdNFFATOcx 
   ClientHeight    =   5775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7785
   ScaleHeight     =   5775
   ScaleMode       =   0  'User
   ScaleWidth      =   7785
   Begin VB.ComboBox ComboAlmox 
      Height          =   315
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   975
      Width           =   2655
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
      Left            =   5970
      TabIndex        =   17
      Top             =   5280
      Width           =   1665
   End
   Begin VB.ComboBox Item 
      Height          =   315
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   540
      Width           =   750
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5955
      ScaleHeight     =   495
      ScaleWidth      =   1560
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   165
      Width           =   1620
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "RastroProdNFFATOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   570
         Picture         =   "RastroProdNFFATOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1065
         Picture         =   "RastroProdNFFATOcx.ctx":068C
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Rastreamento do Produto"
      Height          =   2835
      Left            =   165
      TabIndex        =   5
      Top             =   2250
      Width           =   7455
      Begin MSMask.MaskEdBox LoteData 
         Height          =   255
         Left            =   4260
         TabIndex        =   7
         Top             =   465
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.TextBox FilialOP 
         BorderStyle     =   0  'None
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
         Left            =   2430
         TabIndex        =   9
         Top             =   450
         Width           =   1785
      End
      Begin VB.TextBox Lote 
         BorderStyle     =   0  'None
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
         Left            =   810
         TabIndex        =   6
         Top             =   450
         Width           =   1560
      End
      Begin MSMask.MaskEdBox QuantLote 
         Height          =   225
         Left            =   5445
         TabIndex        =   8
         Top             =   450
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
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   7020
         _ExtentX        =   12383
         _ExtentY        =   3281
         _Version        =   393216
         Rows            =   7
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
         Left            =   3960
         TabIndex        =   21
         Top             =   2370
         Width           =   510
      End
      Begin VB.Label QuantTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4545
         TabIndex        =   20
         Top             =   2340
         Width           =   1470
      End
   End
   Begin VB.Label Quantidade 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1380
      TabIndex        =   23
      Top             =   1845
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
      Left            =   270
      TabIndex        =   22
      Top             =   1875
      Width           =   1050
   End
   Begin VB.Label LabelAlmoxarifado 
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
      Height          =   195
      Left            =   165
      TabIndex        =   19
      Top             =   1050
      Width           =   1155
   End
   Begin VB.Label LabelItem 
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
      Left            =   885
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   16
      Top             =   585
      Width           =   435
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
      Left            =   585
      TabIndex        =   15
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Produto 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1380
      TabIndex        =   14
      Top             =   1425
      Width           =   1410
   End
   Begin VB.Label Descricao 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2910
      TabIndex        =   13
      Top             =   1425
      Width           =   4710
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
      Left            =   4455
      TabIndex        =   12
      Top             =   1890
      Width           =   780
   End
   Begin VB.Label UnidadeMedida 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5295
      TabIndex        =   11
      Top             =   1860
      Width           =   1350
   End
End
Attribute VB_Name = "RastroProdNFFATOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit




'HElp
Const IDH_RASTROPRODNFFAT = 0

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim gcolItemNF As Collection
Dim gobjGenerico As AdmGenerico
Dim giItemNF As Integer
Dim giAlmox As Integer

Dim objGridRastro As AdmGrid
Dim iGrid_Almoxarifado_Col As Integer
Dim iGrid_QuantAlocadaAlm_Col As Integer
Dim iGrid_Lote_Col As Integer
Dim iGrid_LoteData_Col As Integer
Dim iGrid_LoteQuantDisp_Col As Integer
Dim iGrid_QuantLote_Col As Integer
Dim iGrid_FilialOP_Col As Integer

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
    If lErro <> SUCESSO Then gError 83102

    'Limpa a Tela
    Call Limpa_Tela_Rastreamento

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 83102

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166050)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim lNumAlmoxarifados As Long

On Error GoTo Erro_Form_Load

    QuantLote.Format = FORMATO_ESTOQUE

    'Lê quantos Almoxarifados tem na filialempresa
    lErro = CF("AlmoxarifadosFilial_Le_Quantidade",giFilialEmpresa, lNumAlmoxarifados)
    If lErro <> SUCESSO Then gError 75616

    'Se a FilialEmpresa não possui Almoxarifados, erro
    If lNumAlmoxarifados = 0 Then gError 75617

    Set objGridRastro = New AdmGrid
    Set objEventoLote = New AdmEvento

    'Inicializa o grid de Rastreamento
    lErro = Inicializa_Grid_Rastreamento(objGridRastro, lNumAlmoxarifados)
    If lErro <> SUCESSO Then gError 75618

    giItemNF = -1
    giAlmox = -1
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 75616, 75618

        Case 75617
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ALMOXARIFADO_FILIAL", gErr, giFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166051)

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
    If Len(Trim(Produto.Caption)) = 0 Then gError 83103
    
    'Verifica se tem alguma linha selecionada no Grid
    If GridRastro.Row = 0 Then gError 83104
        
    'Formata o produto
    lErro = CF("Produto_Formata",Produto.Caption, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 83105
    
    'Lê o produto
    objProduto.sCodigo = sProdutoFormatado
    lErro = CF("Produto_Le",objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 83106
    
    'Produto não cadastrado
    If lErro = 28030 Then gError 83107
        
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
        
        Case 83103
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case 83104
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 83105, 83106
        
        Case 83107
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166052)
    
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
        If lErro <> SUCESSO Then gError 83106
            
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le",objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 83107
            
        'Se o produto não está cadastrado, erro
        If lErro = 28030 Then gError 83108
                
        'Se o Produto foi preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
            'Se o produto possuir rastro por lote
            If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
                
                For iLinha = 1 To objGridRastro.iLinhasExistentes
                    If iLinha <> objGridRastro.objGrid.Row Then
                        If objGridRastro.objGrid.TextMatrix(iLinha, iGrid_Lote_Col) = objRastroLote.sCodigo Then gError 83109
                    End If
                Next
        
            'Se o produto possuir rastro por OP
            ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
                                                
                For iLinha = 1 To objGridRastro.iLinhasExistentes
                    If iLinha <> objGridRastro.objGrid.Row Then
                        If objGridRastro.objGrid.TextMatrix(iLinha, iGrid_Lote_Col) = objRastroLote.sCodigo And Codigo_Extrai(objGridRastro.objGrid.TextMatrix(iLinha, iGrid_FilialOP_Col)) = objRastroLote.iFilialOP Then gError 83110
                    End If
                Next
                
            End If

        End If

        'Coloca o Lote na tela
        GridRastro.TextMatrix(GridRastro.Row, iGrid_Lote_Col) = objRastroLote.sCodigo
        Lote.Text = objRastroLote.sCodigo
        
        'Lê lote e preenche dados
        lErro = Lote_Saida_Celula(objRastroLote)
        If lErro <> SUCESSO Then gError 83111
        
    End If

    Me.Show

    Exit Sub

Erro_objEventoLote_evSelecao:

    Select Case gErr

        Case 83106, 83107, 83111
        
        Case 83108
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 83109
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_JA_UTILIZADO_GRID", gErr, objRastroLote.sCodigo)
            
        Case 83110
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_FILIALOP_JA_UTILIZADO_GRID", gErr, Lote.Text, objRastroLote.iFilialOP)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 166053)

    End Select

    Exit Sub

End Sub

Private Sub Item_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objItemNF As ClassItemNF
'Dim colEstoqueProduto As New colEstoqueProduto
'Dim objEstoqueProduto As ClassEstoqueProduto
Dim objProduto As New ClassProduto
'Dim objItemPV As New ClassItemPedido
'Dim colReserva As New colReservaItem
'Dim objReservaItem As ClassReservaItem
Dim objItemNFAloc As ClassItemNFAlocacao
Dim sProdutoMascarado As String

On Error GoTo Erro_Item_Click

    'testa se houva alguma alteração
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 83109

    'Verifica se tem um item selecionado
    If Item.ListIndex = -1 Then Exit Sub

    If CInt(Item.Text) = giItemNF Then Exit Sub

    'Guarda o ItemNF escolhido
    Set objItemNF = gcolItemNF.Item(CInt(Item.Text))

'    'Lê os estoques do produto
'    lErro = CF("EstoquesProduto_Le_Filial",objItemNF.sProduto, colEstoqueProduto)
'    If lErro <> SUCESSO Then gError 75619

    '??? Jones: o que reserva tem a ver com isso ?
''    'Se o itemNF foi gerado por um itemPV
''    If objItemNF.lNumIntItemPedVenda > 0 Then
''
''        objItemPV.lNumIntDoc = objItemNF.lNumIntItemPedVenda
''        objItemPV.sProduto = objItemNF.sProduto
''
''        'Lê as reservas do item do pedido
''        lErro = CF("ReservasItemPV_Le_NumIntOrigem",objItemPV, colReserva)
''        If lErro <> SUCESSO And lErro <> 51601 Then gError 75620
''
''    End If
''
''    iIndice = 0

    '??? Jones: o que estoque disponivel tem a ver com isso ?
''
''    'Exclui almoxarifados sem estoque disponível
''    For Each objEstoqueProduto In colEstoqueProduto
''
''        iIndice = iIndice + 1
''
''        '??? Jones: o que reserva tem a ver com isso ?
''''        'Verifica se há uma reserva para esse item nesse almoxarifado
''''        'e inclui a qtd reservada como disponivel p\ esse pedido
''''        Call Procura_Almoxarifado(objEstoqueProduto, colReserva)
''
''        'SE não tiver quantidade disponível e nem reserva nesse almoxarifado
''        If objEstoqueProduto.dQuantDisponivel = 0 Then
''
''            'Retira o estoque produto da coleção
''            colEstoqueProduto.Remove (iIndice)
''
''            iIndice = iIndice - 1
''
''        End If
''
''    Next

    'Lê o produto
    objProduto.sCodigo = objItemNF.sProduto
    lErro = CF("Produto_Le",objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 75621

    'Se não encontrou o produto, erro
    If lErro = 28030 Then gError 75622

    'Mascara o produto
    lErro = Mascara_MascararProduto(objItemNF.sProduto, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 75629

    Produto.Caption = sProdutoMascarado
    Descricao.Caption = objItemNF.sDescricaoItem
    UnidadeMedida.Caption = objItemNF.sUMEstoque

    objItemNF.iClasseUM = objProduto.iClasseUM
    objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque

    ComboAlmox.Clear
    
    For Each objItemNFAloc In objItemNF.colAlocacoes
        ComboAlmox.AddItem CStr(objItemNFAloc.iAlmoxarifado) & SEPARADOR & objItemNFAloc.sAlmoxarifado
        ComboAlmox.ItemData(ComboAlmox.NewIndex) = objItemNFAloc.iAlmoxarifado
    Next

    'seleciona o primeiro almoxarifado
    If ComboAlmox.ListCount > 0 Then ComboAlmox.ListIndex = 0

    'Limpa o Grid
    Call Grid_Limpa(objGridRastro)

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


'    'Lê Rastreamento de Lotes associados as alocações do ItemNF
'    lErro = RastroLoteSaldo_Le_ItemNF2(objItemNF, colEstoqueProduto)
'    If lErro <> SUCESSO Then gError 75623

'    'Lê Rastreamento de Lotes abertos
'    lErro = RastroLoteSaldo_Le_ItemNF(objItemNF, colEstoqueProduto)
'    If lErro <> SUCESSO Then gError 75623

'    'Prenche a tela com os dados do estoque do produto
'    lErro = Preenche_Tela(objItemNF)
'    If lErro <> SUCESSO Then gError 75628

    giItemNF = CInt(Item.Text)

    iAlterado = 0

    Exit Sub

Erro_Item_Click:

    Select Case gErr

        Case 75619, 75620, 75621, 75623, 75628, 75629

        Case 75622
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objItemNF.sProduto)

        Case 83109
            For iIndice = 0 To Item.ListCount - 1
                If CInt(Item.List(iIndice)) = giItemNF Then
                    Item.ListIndex = iIndice
                    Exit For
                End If
            Next

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166054)

    End Select

    Exit Sub

End Sub

Private Sub ComboAlmox_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objItemNF As ClassItemNF
Dim objProduto As New ClassProduto
Dim objItemNFAloc As ClassItemNFAlocacao
Dim dQuantidade As Double

On Error GoTo Erro_ComboAlmox_Click

    'testa se houva alguma alteração
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 83110

    'Verifica se tem um item selecionado
    If ComboAlmox.ListIndex = -1 Then Exit Sub

    If ComboAlmox.ItemData(ComboAlmox.ListIndex) = giAlmox Then Exit Sub

    'Limpa o Grid
    Call Grid_Limpa(objGridRastro)

'    'Lê Rastreamento de Lotes associados as alocações do ItemNF
'    lErro = RastroLoteSaldo_Le_ItemNF2(objItemNF, colEstoqueProduto)
'    If lErro <> SUCESSO Then gError 75623

'    'Lê Rastreamento de Lotes abertos
'    lErro = RastroLoteSaldo_Le_ItemNF(objItemNF, colEstoqueProduto)
'    If lErro <> SUCESSO Then gError 75623

    'Prenche a tela com os dados do estoque do produto
    lErro = Preenche_Tela()
    If lErro <> SUCESSO Then gError 83111

    giAlmox = ComboAlmox.ItemData(ComboAlmox.ListIndex)
    
    Set objItemNF = gcolItemNF.Item(giItemNF)
    
    For Each objItemNFAloc In objItemNF.colAlocacoes
        If objItemNFAloc.iAlmoxarifado = giAlmox Then
            dQuantidade = dQuantidade + objItemNFAloc.dQuantidade
        End If
    Next
    
    Quantidade.Caption = CStr(dQuantidade)
    
    iAlterado = 0

    Exit Sub

Erro_ComboAlmox_Click:

    Select Case gErr

        Case 75619, 75620, 75621, 75623, 75628

        Case 75622
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objItemNF.sProduto)

        Case 83110
            For iIndice = 0 To Item.ListCount - 1
                If ComboAlmox.ItemData(iIndice) = giAlmox Then
                    ComboAlmox.ListIndex = iIndice
                    Exit For
                End If
            Next

        Case 83111

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166055)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Grid_Rastreamento(objGridInt As AdmGrid, lNumAlmoxarifados As Long) As Long
'Inicializa o Grid de Alocação

Dim iIndice As Integer

    Set objGridRastro.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
'    objGridInt.colColuna.Add ("Almoxarifado")
'    objGridInt.colColuna.Add ("Qtd. Alocada Almox.")
    objGridInt.colColuna.Add ("Lote")
    objGridInt.colColuna.Add ("FilialOP do Lote")
    objGridInt.colColuna.Add ("Data do Lote")
'    objGridInt.colColuna.Add ("Qtd. Disp. Lote")
    objGridInt.colColuna.Add ("Qtd. Alocada Lote")

    'Controles que participam do Grid
'    objGridInt.colCampo.Add (Almoxarifado.Name)
'    objGridInt.colCampo.Add (QuantAlocadaAlmox.Name)
    objGridInt.colCampo.Add (Lote.Name)
    objGridInt.colCampo.Add (FilialOP.Name)
    objGridInt.colCampo.Add (LoteData.Name)
'    objGridInt.colCampo.Add (LoteQuantDisponivel.Name)
    objGridInt.colCampo.Add (QuantLote.Name)

    'Colunas da Grid
'    iGrid_Almoxarifado_Col = 1
'    iGrid_QuantAlocadaAlm_Col = 2
    iGrid_Lote_Col = 1
    iGrid_FilialOP_Col = 2
    iGrid_LoteData_Col = 3
'    iGrid_LoteQuantDisp_Col = 5
    iGrid_QuantLote_Col = 4
 
    'Grid do GridInterno
    objGridInt.objGrid = GridRastro

    'Largura da primeira coluna
    GridRastro.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    objGridInt.iLinhasVisiveis = 6
    
    'Posiciona os painéis totalizadores
    QuantTotal.Top = objGridInt.objGrid.Top + objGridInt.objGrid.Height
    QuantTotal.Left = objGridInt.objGrid.Left
    For iIndice = 0 To iGrid_QuantLote_Col - 1
        QuantTotal.Left = QuantTotal.Left + objGridInt.objGrid.ColWidth(iIndice) + objGridInt.objGrid.GridLineWidth + 20
    Next
    
    QuantTotal.Width = objGridInt.objGrid.ColWidth(iGrid_QuantLote_Col)
    
    LabelTotal.Top = QuantTotal.Top + (QuantTotal.Height - LabelTotal.Height) / 2
    LabelTotal.Left = QuantTotal.Left - LabelTotal.Width
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridRastro)

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
        If lErro <> SUCESSO And lErro <> 28030 Then gError 75641

        'Se o produto não está cadastrado, erro
        If lErro = 28030 Then gError 75642

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

        Case 75641

        Case 75642
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166056)

    End Select

    Exit Function

End Function

Private Function Preenche_Tela() As Long

Dim objRastroItemNF As ClassRastroItemNF
Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim objItemNF As ClassItemNF
Dim dQuantTotal As Double

On Error GoTo Erro_Preenche_Tela

    Set objItemNF = gcolItemNF(CInt(Item.Text))

    'Coloca no grid os rastros do item/almoxarifado
    For Each objRastroItemNF In objItemNF.colRastreamento

        'Pega somente os rastros do almoxarifado selecionado
        If objRastroItemNF.iAlmoxCodigo = ComboAlmox.ItemData(ComboAlmox.ListIndex) Then

'            GridRastro.TextMatrix(objGridRastro.iLinhasExistentes + 1, iGrid_Almoxarifado_Col) = objRastroItemNF.sAlmoxNomeRed
'            GridRastro.TextMatrix(objGridRastro.iLinhasExistentes + 1, iGrid_QuantAlocadaAlm_Col) = Formata_Estoque(objRastroItemNF.dAlmoxQtdAlocada)
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

'            GridRastro.TextMatrix(objGridRastro.iLinhasExistentes + 1, iGrid_LoteQuantDisp_Col) = Formata_Estoque(objRastroItemNF.dLoteQtdDisp)
            GridRastro.TextMatrix(objGridRastro.iLinhasExistentes + 1, iGrid_QuantLote_Col) = Formata_Estoque(objRastroItemNF.dLoteQdtAlocada)

            'Incrementa o número de linhas existentes no Grid
            objGridRastro.iLinhasExistentes = objGridRastro.iLinhasExistentes + 1

        End If

    Next

    QuantTotal.Caption = Formata_Estoque(dQuantTotal)

    Preenche_Tela = SUCESSO

    Exit Function

Erro_Preenche_Tela:

    Select Case gErr

        Case 75696

        Case 75697
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166057)

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
                If lErro <> SUCESSO Then gError 83112

            'FilialOP
            Case iGrid_FilialOP_Col
                lErro = Saida_Celula_FilialOP(objGridInt)
                If lErro <> SUCESSO Then gError 83113
            
            'Quantidade Alocada do Lote
            Case iGrid_QuantLote_Col
                lErro = Saida_Celula_QuantLote(objGridInt)
                If lErro <> SUCESSO Then gError 75630

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 75631

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 75630, 83112, 83113

        Case 75631
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166058)

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
        If lErro <> SUCESSO Then gError 83114

        'Se a quantidade alocada do lote for maior que a quantidade disponível, erro
        If StrParaDbl(QuantLote.Text) > StrParaDbl(Quantidade.Caption) Then gError 75634

    End If

    'totaliza as quantidades dos lotes e mostra no campo QuantTotal
    QuantTotal.Caption = Format(GridQuantLote_Soma() + StrParaDbl(QuantLote.Text), FORMATO_ESTOQUE)

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 83115

    Saida_Celula_QuantLote = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantLote:

    Saida_Celula_QuantLote = gErr

    Select Case gErr

        Case 83114, 83115
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 75634
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTLOTE_MAIOR_QUANTALM", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166059)

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
        If Len(Trim(Produto.Caption)) = 0 Then gError 83116
        
        If Item.ListIndex = -1 Then gError 83117
        
        'Formata o Produto para o BD
        lErro = CF("Produto_Formata",Produto.Caption, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 83118
            
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        objProduto.sCodigo = sProdutoFormatado
        lErro = CF("Produto_Le",objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 83119
            
        'Se o produto não está cadastrado, erro
        If lErro = 28030 Then gError 83120
                
        'Se o Produto foi preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
            objRastroLote.dtDataEntrada = DATA_NULA
            
            'Se o produto possuir rastro por lote
            If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
                
                For iLinha = 1 To objGridInt.iLinhasExistentes
                    If iLinha <> objGridInt.objGrid.Row Then
                        If objGridInt.objGrid.TextMatrix(iLinha, iGrid_Lote_Col) = Lote.Text Then gError 83121
                    End If
                Next
                
                objRastroLote.sCodigo = Lote.Text
                objRastroLote.sProduto = sProdutoFormatado
                
                'Lê o Rastreamento do Lote vinculado ao produto
                lErro = CF("RastreamentoLote_Le",objRastroLote)
                If lErro <> SUCESSO And lErro <> 75710 Then gError 83122
                
                'Se não encontrou --> Erro
                If lErro = 75710 Then gError 83123
                
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
                                gError 83124
                            End If
                            Exit For
                        End If
                    End If
                Next
                
                If objRastroLote.iFilialOP <> 0 Then
                
                    'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                    lErro = CF("RastreamentoLote_Le",objRastroLote)
                    If lErro <> SUCESSO And lErro <> 75710 Then gError 83125
                
                    'Se não encontrou --> Erro
                    If lErro = 75710 Then gError 83126
                
                End If
                
            End If
        
        End If
    
        'Preenche campos do lote
        lErro = Lote_Saida_Celula(objRastroLote)
        If lErro <> SUCESSO Then gError 83127
    
    Else
    
        GridRastro.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteData_Col) = ""
        
    End If
                                    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 83128

    Saida_Celula_Lote = SUCESSO

    Exit Function

Erro_Saida_Celula_Lote:

    Saida_Celula_Lote = gErr

    Select Case gErr
        
        Case 83116, 83118, 83119, 83122, 83125, 83127, 83128
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83117
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMNF_NAO_SELECIONADO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83120
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83121
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_JA_UTILIZADO_GRID", gErr, Lote.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 83123
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("RastreamentoLote", objRastroLote)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
        
        Case 83124
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_FILIALOP_JA_UTILIZADO_GRID", gErr, Lote.Text, objRastroLote.iFilialOP)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83126
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OP_RASTREAMENTO_INEXISTENTE", gErr, objRastroLote.sCodigo, objRastroLote.iFilialOP, objRastroLote.sProduto)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166060)

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
        
        If Item.ListIndex = -1 Then gError 83129
        
        'Valida a Filial
        lErro = TP_FilialEmpresa_Le(FilialOP.Text, objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 71971 And lErro <> 71972 Then gError 83130

        'Se não for encontrado --> Erro
        If lErro = 71971 Then gError 83131
        If lErro = 71972 Then gError 83132
        
        If Len(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Lote_Col)) <> 0 Then
        
            For iLinha = 1 To objGridInt.iLinhasExistentes
                If iLinha <> objGridInt.objGrid.Row Then
                    If objGridInt.objGrid.TextMatrix(iLinha, iGrid_Lote_Col) = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Lote_Col) And Codigo_Extrai(objGridInt.objGrid.TextMatrix(iLinha, iGrid_FilialOP_Col)) = objFilialEmpresa.iCodFilial Then gError 83133
                End If
            Next
        
            'Formata o produto
            lErro = CF("Produto_Formata",Produto.Caption, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 83134
        
            objRastroLote.sCodigo = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Lote_Col)
            objRastroLote.sProduto = sProdutoFormatado
            objRastroLote.iFilialOP = objFilialEmpresa.iCodFilial
                
            'Lê o Rastreamento do Lote vinculado ao produto
            lErro = CF("RastreamentoLote_Le",objRastroLote)
            If lErro <> SUCESSO And lErro <> 75710 Then gError 83135
                
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
    If lErro <> SUCESSO Then gError 83137

    Saida_Celula_FilialOP = SUCESSO

    Exit Function

Erro_Saida_Celula_FilialOP:

    Saida_Celula_FilialOP = gErr

    Select Case gErr

        Case 83129
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMNF_NAO_SELECIONADO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 83130, 83134, 83135, 83137
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83131, 83132
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, FilialOP.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 83133
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_FILIALOP_JA_UTILIZADO_GRID", gErr, objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Lote_Col), objFilialEmpresa.iCodFilial)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166061)

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

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Se nenhum item foi selecionado ==> erro
    If giItemNF = -1 Then gError 83133
    
    If giAlmox = -1 Then gError 83134

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 75635

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 75635

        Case 83133
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMNF_NAO_SELECIONADO", gErr)
        
        Case 83134
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOX_NAO_SELECIONADO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 166062)

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
        If lErro <> SUCESSO And lErro <> 27378 Then gError 83104

        'Se não encontrou a FilialEmpresa
        If lErro = 27378 Then gError 83105

        GridRastro.TextMatrix(GridRastro.Row, iGrid_FilialOP_Col) = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome
    
    End If
    
    If objGridRastro.objGrid.Row - objGridRastro.objGrid.FixedRows = objGridRastro.iLinhasExistentes Then
        objGridRastro.iLinhasExistentes = objGridRastro.iLinhasExistentes + 1
    End If
    
    Lote_Saida_Celula = SUCESSO
    
    Exit Function
        
Erro_Lote_Saida_Celula:

    Lote_Saida_Celula = gErr
    
    Select Case gErr
        
        Case 83104
        
        Case 83105
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 166063)
    
    End Select
    
    Exit Function
    
End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objItemNF As ClassItemNF
Dim iIndice As Integer
Dim iLinha As Integer
Dim dQuantLoteTotal As Double
Dim sAlmoxarifado As String
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Gravar_Registro

    If giItemNF <> -1 And giAlmox <> -1 Then

        GL_objMDIForm.MousePointer = vbHourglass
                
        'Formata o Produto para o BD
        lErro = CF("Produto_Formata",Produto.Caption, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 83135
            
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le",objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 83136
            
        'Se o produto não está cadastrado, erro
        If lErro = 28030 Then gError 83137
            
        'Para cada linha do Grid
        For iIndice = 1 To objGridRastro.iLinhasExistentes
                                    
            'Se o lote não foi preenchido, erro
            If Len(Trim(GridRastro.TextMatrix(iIndice, iGrid_Lote_Col))) = 0 Then gError 83138
            
            'Se o produto possuir rastro por lote
            If objProduto.iRastro = PRODUTO_RASTRO_OP And Len(Trim(GridRastro.TextMatrix(iIndice, iGrid_FilialOP_Col))) = 0 Then gError 83139
    
            'Se a quantidade não foi preenchida,erro
            If Len(Trim(GridRastro.TextMatrix(iIndice, iGrid_QuantLote_Col))) = 0 Then gError 83140
            
            'Se a quantidade está zerada, erro
            If StrParaDbl(GridRastro.TextMatrix(iIndice, iGrid_QuantLote_Col)) = 0 Then gError 83141
            
            dQuantLoteTotal = dQuantLoteTotal + StrParaDbl(GridRastro.TextMatrix(iIndice, iGrid_QuantLote_Col))
            
        Next
    
        'Se a quantidade alocada do Lote foi maior que a quantidade alocada no almoxarifado
        If dQuantLoteTotal > StrParaDbl(Quantidade.Caption) Then gError 83142
    
        Set objItemNF = gcolItemNF.Item(giItemNF)
    
        'Move dados da tela para a memória
        lErro = Move_Tela_Memoria(objItemNF)
        If lErro <> SUCESSO Then gError 83143
    
        giRetornoTela = vbOK
    
        GL_objMDIForm.MousePointer = vbDefault

    End If

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 83135, 83136, 83143

        Case 83137
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 83138
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRID_LOTE_NAO_PREENCHIDO", gErr, iIndice)
        
        Case 83139
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_OP_NAO_PREENCHIDA", gErr, iIndice)
        
        Case 83140
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRID_QUANTLOTE_NAO_PREENCHIDA", gErr, iIndice)
        
        Case 83141
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRID_QUANTLOTE_ZERADA", gErr, iIndice)

        Case 83142
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTTOTAL_LOTE_MAIOR_ALMOXARIFADO", gErr, dQuantLoteTotal, StrParaDbl(Quantidade.Caption))

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166064)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(objItemNF As ClassItemNF) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objRastroItemNF As ClassRastroItemNF
Dim iLinha As Integer

On Error GoTo Erro_Move_Tela_Memoria

    'Retira os rastreamentos antigos deste item/almoxarifado
    For iIndice = objItemNF.colRastreamento.Count To 1 Step -1
        Set objRastroItemNF = objItemNF.colRastreamento.Item(iIndice)
        If objRastroItemNF.iAlmoxCodigo = giAlmox Then objItemNF.colRastreamento.Remove (iIndice)
    Next
    
    
    'Coloca os novos rastreamentos para o almoxarifado em questão
    For iLinha = 1 To objGridRastro.iLinhasExistentes

        'Cria novo item de rastreamento
        Set objRastroItemNF = New ClassRastroItemNF
        objRastroItemNF.dAlmoxQtdAlocada = StrParaDbl(Quantidade.Caption)
        objRastroItemNF.dLoteQdtAlocada = StrParaDbl(GridRastro.TextMatrix(iLinha, iGrid_QuantLote_Col))
        objRastroItemNF.dtLoteData = StrParaDate(GridRastro.TextMatrix(iLinha, iGrid_LoteData_Col))
        objRastroItemNF.iAlmoxCodigo = giAlmox
        objRastroItemNF.sLote = GridRastro.TextMatrix(iLinha, iGrid_Lote_Col)
        
        'Lê almoxarifado
        objAlmoxarifado.iCodigo = giAlmox
        
        lErro = CF("Almoxarifado_Le",objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25060 Then gError 83144
        
        'Se não encontrou o almoxarifado, erro
        If lErro = 25060 Then gError 83145
        
        objRastroItemNF.sAlmoxNomeRed = objAlmoxarifado.sNomeReduzido
        objRastroItemNF.iLoteFilialOP = Codigo_Extrai(GridRastro.TextMatrix(iLinha, iGrid_FilialOP_Col))
                    
        'Adiciona na coleção
        objItemNF.colRastreamento.Add objRastroItemNF

    Next

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 83144

        Case 83145
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", gErr, objAlmoxarifado.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166065)

    End Select

    Exit Function

End Function

Private Sub Limpa_Tela_Rastreamento()

Dim iIndice As Integer

    giItemNF = -1
    giAlmox = -1
    Item.ListIndex = -1
    ComboAlmox.ListIndex = -1

    'Limpa o Grid de Rastreamento
    Call Grid_Limpa(objGridRastro)
    
    Produto.Caption = ""
    Descricao.Caption = ""
    UnidadeMedida.Caption = ""
    Quantidade.Caption = ""
    QuantTotal.Caption = ""

End Sub

'Private Sub Procura_Almoxarifado(objEstoqueProduto As ClassEstoqueProduto, colReserva As colReservaItem)
''Percorre uma colecao de reservas e verifica se o almoxarifado passado está na coleção.
'
'Dim iIndice As Integer
'
'    For iIndice = 1 To colReserva.Count
'
'        If objEstoqueProduto.iAlmoxarifado = colReserva(iIndice).iAlmoxarifado Then
'            objEstoqueProduto.dQuantDispNossa = objEstoqueProduto.dQuantDispNossa + colReserva(iIndice).dQuantidade
'            Exit For
'        End If
'
'    Next
'
'    Exit Sub
'
'End Sub

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166066)

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

    Parent.HelpContextID = IDH_RASTROPRODNFFAT
    Set Form_Load_Ocx = Me
    Caption = "Rastreamento de Produto"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RastroProdNFFAT"

End Function

Public Sub Show()
'    Parent.Show
'    Parent.SetFocus
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

Private Sub UnidadeMedida_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(UnidadeMedida, Source, X, Y)
End Sub

Private Sub UnidadeMedida_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(UnidadeMedida, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Descricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Descricao, Source, X, Y)
End Sub

Private Sub Descricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Descricao, Button, Shift, X, Y)
End Sub

Private Sub Produto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Produto, Source, X, Y)
End Sub

Private Sub Produto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Produto, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub LabelItem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelItem, Source, X, Y)
End Sub

Private Sub LabelItem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelItem, Button, Shift, X, Y)
End Sub

Private Sub LabelAlmoxarifado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelAlmoxarifado, Source, X, Y)
End Sub

Private Sub LabelAlmoxarifado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelAlmoxarifado, Button, Shift, X, Y)
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

Private Sub Quantidade_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Quantidade, Source, X, Y)
End Sub

Private Sub Quantidade_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Quantidade, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub



'***************** Funções de Leitura e Gravação *******************************************

''Subir para RotinasMAT - ClassSELECT
'Function RastroLoteSaldo_Le_ItemNF(objItemNF As ClassItemNF, colEstoqueProduto As colEstoqueProduto) As Long
''Lê dados de RastreamentoLoteSaldo e RastreamentoLote que estão vinculados ao Produto do ItemNF
''e almoxarifados passados em ColEstoqueProduto
'
'Dim lErro As Long
'Dim alComando(0 To 1) As Long
'Dim iIndice As Integer
'Dim iIndice2 As Integer
'Dim tRastroLoteSaldo As typeRastreamentoLoteSaldo
'Dim dtDataLote As Date
'Dim objRastroItemNF As ClassRastroItemNF
'Dim dQuantMovto As Double
'Dim objAlocacao As ClassItemNFAlocacao
'Dim objItemMovEstoque As New ClassItemMovEstoque
'
'On Error GoTo Erro_RastroLoteSaldo_Le_ItemNF
'
'    'Abre Comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 75624
'    Next
'
'    'Para cada Almoxarifado vinculado ao produto
'    For iIndice = 1 To colEstoqueProduto.Count
'
'        tRastroLoteSaldo.sLote = String(STRING_OPCODIGO, 0)
'
'        'Procura por Rastreamentos vinculado ao Produto e almoxarifado passados
'        lErro = Comando_Executar(alComando(0), "SELECT DISTINCT RastreamentoLote.Lote, RastreamentoLoteSaldo.QuantDispNossa, RastreamentoLoteSaldo.QuantConsig3, DataEntrada, FilialOP, RastreamentoLote.NumIntDoc FROM RastreamentoLoteSaldo, RastreamentoLote WHERE RastreamentoLoteSaldo.NumIntDocLote = RastreamentoLote.NumIntDoc AND RastreamentoLote.Status <> ? AND RastreamentoLoteSaldo.Produto = ? AND RastreamentoLoteSaldo.Almoxarifado = ? ORDER BY DataEntrada", tRastroLoteSaldo.sLote, tRastroLoteSaldo.dQuantDispNossa, tRastroLoteSaldo.dQuantConsig3, dtDataLote, tRastroLoteSaldo.iFilialOP, tRastroLoteSaldo.lNumIntDocLote, STATUS_BAIXADO, objItemNF.sProduto, colEstoqueProduto(iIndice).iAlmoxarifado)
'        If lErro <> AD_SQL_SUCESSO Then gError 75625
'
'        lErro = Comando_BuscarPrimeiro(alComando(0))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 75626
'
'        'Equanto encontrar
'        Do While lErro = AD_SQL_SUCESSO
'
'            'Verifica se o rastreamento já estava na coleção
'            For iIndice2 = 1 To objItemNF.colRastreamento.Count
'                If objItemNF.colRastreamento(iIndice2).iAlmoxCodigo = colEstoqueProduto(iIndice).iAlmoxarifado And objItemNF.colRastreamento(iIndice2).sLote = tRastroLoteSaldo.sLote Then
'                    Exit For
'                End If
'            Next
'
'            'Se não encontrou o rastreamento
'            If iIndice2 > objItemNF.colRastreamento.Count Then
'
'                Set objRastroItemNF = New ClassRastroItemNF
'
'                'Guarda os dados lidos no BD na coleção de Rastreamento do ItemNF
'                objRastroItemNF.sLote = tRastroLoteSaldo.sLote
'                objRastroItemNF.iAlmoxCodigo = colEstoqueProduto(iIndice).iAlmoxarifado
'                objRastroItemNF.dtLoteData = dtDataLote
'                objRastroItemNF.iLoteFilialOP = tRastroLoteSaldo.iFilialOP
'                objRastroItemNF.dLoteQtdDisp = tRastroLoteSaldo.dQuantDispNossa + tRastroLoteSaldo.dQuantConsig3 + dQuantMovto
'                objRastroItemNF.sAlmoxNomeRed = colEstoqueProduto(iIndice).sAlmoxarifadoNomeReduzido
'                objRastroItemNF.dLoteQdtAlocada = 0
'
'                'Procura pela Quantidade alocada no almoxarifado
'                For Each objAlocacao In objItemNF.colAlocacoes
'                    If objAlocacao.iAlmoxarifado = objRastroItemNF.iAlmoxCodigo Then
'                        objRastroItemNF.dAlmoxQtdAlocada = objAlocacao.dQuantidade
'                    End If
'                Next
'
'                'Adiciona rastreamento na coleção
'                objItemNF.colRastreamento.Add objRastroItemNF
'
'            End If
'
'            'Busca próximo Rastreamento
'            lErro = Comando_BuscarProximo(alComando(0))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 75627
'
'        Loop
'
'    Next
'
'    'Fecha comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    RastroLoteSaldo_Le_ItemNF = SUCESSO
'
'    Exit Function
'
'Erro_RastroLoteSaldo_Le_ItemNF:
'
'    RastroLoteSaldo_Le_ItemNF = gErr
'
'    Select Case gErr
'
'        Case 75624
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 75625, 75626, 75627
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOLOTE", gErr)
'
'        Case 75688
'
'        Case 75802, 75803
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_RASTREAMENTOMOVTO", gErr)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166067)
'
'    End Select
'
'    'Fecha comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function
'
''Copiada de outras telas
'Function MovEstoque_Le_ItemNF(objItemMovEstoque As ClassItemMovEstoque) As Long
''Lê o NumIntDoc e Código do MovimentoEstoque a partir do NumIntDoc do ItemNF passado
'
'Dim lErro As Long
'Dim lComando As Long
'Dim lNumIntDoc As Long
'Dim lCodigo As Long
'Dim dtData As Date
'
'On Error GoTo Erro_MovEstoque_Le_ItemNF
'
'    'Abre comandos
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 75793
'
'    'Lê NumIntDoc de MovimentoEstoque
'    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Codigo, Data FROM MovimentoEstoque WHERE NumIntDocOrigem = ? AND TipoNumIntDocOrigem = ? AND FilialEmpresa = ?", lNumIntDoc, lCodigo, dtData, objItemMovEstoque.lNumIntDocOrigem, objItemMovEstoque.iTipoNumIntDocOrigem, objItemMovEstoque.iFilialEmpresa)
'    If lErro <> AD_SQL_SUCESSO Then gError 75794
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 75795
'
'    'Se não encontrou o movimento estoque, erro
'    If lErro = AD_SQL_SEM_DADOS Then gError 75796
'
'    objItemMovEstoque.lNumIntDoc = lNumIntDoc
'    objItemMovEstoque.lCodigo = lCodigo
'    objItemMovEstoque.dtData = dtData
'
'    'Fecha comandos
'    Call Comando_Fechar(lComando)
'
'    MovEstoque_Le_ItemNF = SUCESSO
'
'    Exit Function
'
'Erro_MovEstoque_Le_ItemNF:
'
'    MovEstoque_Le_ItemNF = gErr
'
'    Select Case gErr
'
'        Case 75793
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 75794, 75795
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)
'
'        Case 75796
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166068)
'
'    End Select
'
'    'Fecha comandos
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function
'
''Subir para RotinasMAT - ClassSELECT
'Function RastroLoteSaldo_Le_ItemNF2(objItemNF As ClassItemNF, colEstoqueProduto As colEstoqueProduto) As Long
''Lê dados de RastreamentoLoteSaldo e RastreamentoLote que estão vinculados ao Produto do ItemNF
''e almoxarifados passados em ColEstoqueProduto
''????
'
'Dim lErro As Long
'Dim alComando(0 To 1) As Long
'Dim iIndice As Integer
'Dim iIndice2 As Integer
'Dim tRastroLoteSaldo As typeRastreamentoLoteSaldo
'Dim dtDataLote As Date, iAlmoxarifado As Integer
'Dim objRastroItemNF As ClassRastroItemNF
'Dim dQuantMovto As Double
'Dim objAlocacao As ClassItemNFAlocacao
'Dim objItemMovEstoque As New ClassItemMovEstoque, sAlmoxNomeRed As String
'
'On Error GoTo Erro_RastroLoteSaldo_Le_ItemNF2
'
'    'Abre Comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 75624
'    Next
'
'    tRastroLoteSaldo.sLote = String(STRING_OPCODIGO, 0)
'    sAlmoxNomeRed = String(STRING_ALMOXARIFADO_NOMEREDUZIDO, 0)
'
'    'Procura por Rastreamentos vinculado ao Produto e itemnf
'    lErro = Comando_Executar(alComando(0), "SELECT Almoxarifado.NomeReduzido, MovimentoEstoque.Almoxarifado, RastreamentoMovto.Quantidade, RastreamentoLote.Lote, RastreamentoLoteSaldo.QuantDispNossa, RastreamentoLoteSaldo.QuantConsig3, DataEntrada, FilialOP FROM RastreamentoMovto, MovimentoEstoque, RastreamentoLote, RastreamentoLoteSaldo, Almoxarifado WHERE RastreamentoLoteSaldo.NumIntDocLote = RastreamentoLote.NumIntDoc AND RastreamentoLoteSaldo.Produto = ? AND RastreamentoMovto.NumIntDocLote = RastreamentoLote.NumIntDoc AND RastreamentoMovto.TipoDocOrigem = ? AND RastreamentoMovto.NumIntDocOrigem = MovimentoEstoque.NumIntDoc AND MovimentoEstoque.FilialEmpresa = ? AND MovimentoEstoque.TipoNumIntDocOrigem = ? AND MovimentoEstoque.NumIntDocOrigem = ? AND MovimentoEstoque.Almoxarifado = Almoxarifado.Codigo", _
'        sAlmoxNomeRed, iAlmoxarifado, dQuantMovto, tRastroLoteSaldo.sLote, tRastroLoteSaldo.dQuantDispNossa, tRastroLoteSaldo.dQuantConsig3, dtDataLote, tRastroLoteSaldo.iFilialOP, objItemNF.sProduto, TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE, giFilialEmpresa, TIPO_ORIGEM_ITEMNF, objItemNF.lNumIntDoc)
'    If lErro <> AD_SQL_SUCESSO Then gError 75625
'
'    lErro = Comando_BuscarPrimeiro(alComando(0))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 75626
'
'    'Equanto encontrar
'    Do While lErro = AD_SQL_SUCESSO
'
'        'Verifica se o rastreamento já estava na coleção
'        For iIndice2 = 1 To objItemNF.colRastreamento.Count
'            If objItemNF.colRastreamento(iIndice2).iAlmoxCodigo = iAlmoxarifado And objItemNF.colRastreamento(iIndice2).sLote = tRastroLoteSaldo.sLote Then
'                Exit For
'            End If
'        Next
'
'        'Se não encontrou o rastreamento
'        If iIndice2 > objItemNF.colRastreamento.Count Then
'
'            Set objRastroItemNF = New ClassRastroItemNF
'
'            'Guarda os dados lidos no BD na coleção de Rastreamento do ItemNF
'            objRastroItemNF.sLote = tRastroLoteSaldo.sLote
'            objRastroItemNF.iAlmoxCodigo = iAlmoxarifado
'            objRastroItemNF.dtLoteData = dtDataLote
'            objRastroItemNF.iLoteFilialOP = tRastroLoteSaldo.iFilialOP
'            objRastroItemNF.dLoteQtdDisp = tRastroLoteSaldo.dQuantDispNossa + tRastroLoteSaldo.dQuantConsig3 + dQuantMovto
'            objRastroItemNF.sAlmoxNomeRed = sAlmoxNomeRed
'            objRastroItemNF.dLoteQdtAlocada = dQuantMovto
'
'            'Procura pela Quantidade alocada no almoxarifado
'            For Each objAlocacao In objItemNF.colAlocacoes
'                If objAlocacao.iAlmoxarifado = objRastroItemNF.iAlmoxCodigo Then
'                    objRastroItemNF.dAlmoxQtdAlocada = objAlocacao.dQuantidade
'                End If
'            Next
'
'            'Adiciona rastreamento na coleção
'            objItemNF.colRastreamento.Add objRastroItemNF
'
'        End If
'
'        'Busca próximo Rastreamento
'        lErro = Comando_BuscarProximo(alComando(0))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 75627
'
'    Loop
'
'    'Fecha comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    RastroLoteSaldo_Le_ItemNF2 = SUCESSO
'
'    Exit Function
'
'Erro_RastroLoteSaldo_Le_ItemNF2:
'
'    RastroLoteSaldo_Le_ItemNF2 = gErr
'
'    Select Case gErr
'
'        Case 75624
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 75625, 75626, 75627
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOLOTE", gErr)
'
'        Case 75688
'
'        Case 75802, 75803
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_RASTREAMENTOMOVTO", gErr)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166069)
'
'    End Select
'
'    'Fecha comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function
'
