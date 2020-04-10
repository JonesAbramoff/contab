VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl AlocacaoProdutoOcx 
   ClientHeight    =   5400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7935
   LockControls    =   -1  'True
   ScaleHeight     =   5400
   ScaleWidth      =   7935
   Begin VB.ComboBox Produto 
      Height          =   315
      Left            =   1860
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   870
      Width           =   1665
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6240
      ScaleHeight     =   495
      ScaleWidth      =   1560
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   1620
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1065
         Picture         =   "AlocacaoProdutoOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   570
         Picture         =   "AlocacaoProdutoOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   75
         Picture         =   "AlocacaoProdutoOcx.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.ComboBox Item 
      Height          =   315
      Left            =   1875
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   345
      Width           =   750
   End
   Begin VB.Frame Frame7 
      Caption         =   "Reserva do Produto"
      Height          =   2685
      Left            =   135
      TabIndex        =   12
      Top             =   2280
      Width           =   7725
      Begin VB.TextBox Responsavel 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   5355
         MaxLength       =   50
         TabIndex        =   5
         Top             =   330
         Width           =   1875
      End
      Begin MSMask.MaskEdBox DataValidade 
         Height          =   225
         Left            =   4245
         TabIndex        =   4
         Top             =   330
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Almoxarifado 
         Height          =   225
         Left            =   645
         TabIndex        =   1
         Top             =   375
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox QuantDisponivel 
         Height          =   225
         Left            =   1980
         TabIndex        =   2
         Top             =   345
         Width           =   1110
         _ExtentX        =   1958
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
      Begin MSMask.MaskEdBox QuantReservada 
         Height          =   225
         Left            =   3120
         TabIndex        =   3
         Top             =   330
         Width           =   1110
         _ExtentX        =   1958
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
         Format          =   "FORMATO_ESTOQUE"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridReserva 
         Height          =   1860
         Left            =   75
         TabIndex        =   6
         Top             =   270
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   3281
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin VB.Label TotalReservado 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3360
         TabIndex        =   14
         Top             =   2265
         Width           =   1440
      End
      Begin VB.Label QuantTotalReserva 
         AutoSize        =   -1  'True
         Caption         =   "Quant. Reservada:"
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
         Left            =   1560
         TabIndex        =   13
         Top             =   2295
         Width           =   1620
      End
   End
   Begin VB.CommandButton BotaoSubstituir 
      Caption         =   "Substituição do Produto"
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
      Left            =   5460
      TabIndex        =   7
      Top             =   5010
      Width           =   2430
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
      Left            =   1350
      TabIndex        =   21
      Top             =   390
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
      Left            =   1065
      TabIndex        =   20
      Top             =   855
      Width           =   735
   End
   Begin VB.Label Descricao 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3675
      TabIndex        =   19
      Top             =   870
      Width           =   3435
   End
   Begin VB.Label QuantReservar 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1845
      TabIndex        =   18
      Top             =   1755
      Width           =   1440
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Quant. a Reservar:"
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
      TabIndex        =   17
      Top             =   1800
      Width           =   1635
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
      Left            =   975
      TabIndex        =   16
      Top             =   1350
      Width           =   780
   End
   Begin VB.Label UnidadeMedida 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1845
      TabIndex        =   15
      Top             =   1320
      Width           =   1440
   End
End
Attribute VB_Name = "AlocacaoProdutoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim gcolItemPedido As ColItemPedido
Dim iProdutoAtual As Integer
Dim bGravandoProduto As Boolean

Dim giListIndexAnterior As Integer 'Guarda o índice do item anterior
Dim dTotalReservado As Double
Dim gcolEstoque As colEstoqueProduto

Dim objGrid2 As AdmGrid
Dim iGrid_Sequencial_Col As Integer
Dim iGrid_Almoxarifado_Col As Integer
Dim iGrid_QuantDisponivel_Col As Integer
Dim iGrid_DataValidade_Col As Integer
Dim iGrid_QuantReservada_Col As Integer
Dim iGrid_Responsavel_Col As Integer

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    'Se não há nenhum ítem selecionado --> sai da rotina
    If Item.ListIndex = -1 Then Exit Sub
    
    'Verificar se foi informada alguma quantidade para reserva
    If Len(Trim(TotalReservado.Caption)) > 0 Then
        If Not CDbl(TotalReservado.Caption) > 0 Then Error 23761
    Else
        Error 23762
    End If

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 23763
    
    If Not bGravandoProduto Then Call Limpa_Tela_Alocacao
    
    iAlterado = 0
    
    Exit Sub

Erro_BotaoGravar_Click:
    
    Select Case Err

        Case 23761, 23762
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TOTAL_RESERVADO_SEM_PREENCHIMENTO", Err)

        Case 23763

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 142741)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 60770

    Call Limpa_Tela_Alocacao

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 60770

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142742)

    End Select

    Exit Sub

End Sub

Private Sub BotaoSubstituir_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objSubstProduto As New ClassSubstProduto
Dim objProduto As New ClassProduto
Dim iNumEstoques As Integer

On Error GoTo Erro_BotaoSubstituir_Click

    If Item.ListIndex = -1 Then Exit Sub
    
    If gcolItemPedido(CInt(Item.Text)).dQuantFaturada > 0 Then gError 23776

    objSubstProduto.sCodProduto = gcolItemPedido(CInt(Item.Text)).sProduto
    Set objSubstProduto.ColItemPedido = gcolItemPedido
    
    giRetornoTela = vbCancel
    Call Chama_Tela_Modal("SubstProduto", objSubstProduto)
    If giRetornoTela = vbOK Then

        objProduto.sCodigo = objSubstProduto.sCodProdutoSubstituto
        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 23758

        'Se não achou o Produto
        If lErro = 28030 Then gError 23807
        
        'Substitui o produto na coleção
        iIndice = Item.ListIndex + 1
        gcolItemPedido(iIndice).sProduto = objProduto.sCodigo
        gcolItemPedido(iIndice).sUnidadeMed = objProduto.sSiglaUMVenda
        gcolItemPedido(iIndice).sDescricao = objProduto.sDescricao
        gcolItemPedido(iIndice).sUMEstoque = objProduto.sSiglaUMEstoque
        gcolItemPedido(iIndice).iClasseUM = objProduto.iClasseUM
        Set gcolItemPedido(iIndice).colReserva = New colReserva
        
        Set gcolEstoque = New colEstoqueProduto
        
        lErro = CF("EstoquesProduto_Le", objSubstProduto.sCodProdutoSubstituto, gcolEstoque)
        If lErro <> SUCESSO And lErro <> 30100 Then gError 23759
        
        'Nenhum dos almoxarifados tem quantidade para este Produto
        If lErro = 30100 Then gError 64430

        'Definir valor de saldo
        For iIndice = 1 To gcolEstoque.Count
            gcolEstoque(iIndice).dSaldo = gcolEstoque(iIndice).dQuantDisponivel
        Next
                
        Call GridReserva_Limpa

        'Retira da coleção os que tem saldo zero
        For iIndice = 1 To gcolEstoque.Count
            If gcolEstoque.Item(iIndice).dSaldo = 0 Then gcolEstoque.Remove iIndice
        Next
                
        lErro = Preenche_Tela2(objProduto, gcolEstoque)
        If lErro <> SUCESSO Then gError 23760

        iAlterado = REGISTRO_ALTERADO

    End If

    Exit Sub

Erro_BotaoSubstituir_Click:

    Select Case gErr

        Case 23758, 23759, 23760

        Case 23776
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_FATURADA_MAIORZERO", gErr, gcolItemPedido(CInt(Item.Text)).dQuantFaturada)

        Case 23807
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, gcolItemPedido(CInt(Item.Text)).sProduto)
        
        Case 64430
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NAO_EXISTE_ESTOQUE", gErr, gcolItemPedido(CInt(Item.Text)).sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142743)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim lAlmoxarifados As Long

On Error GoTo Erro_Form_Load

    QuantReservada.Format = FORMATO_ESTOQUE

    lErro = CF("Almoxarifados_Le_Quantidade", lAlmoxarifados)
    If lErro <> SUCESSO And lErro <> 23798 Then Error 23737

    'Verifica se há algum almoxarifado cadastrado no BD
    If lAlmoxarifados = 0 Then Error 23738

    'Inicializa Grid Reserva
    Set objGrid2 = New AdmGrid
    
    lErro = Inicializa_Grid_Reserva(objGrid2, CInt(lAlmoxarifados))
    If lErro <> SUCESSO Then Error 23739

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 23737, 23739

        Case 23738
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NENHUM_ALMOXARIFADO_BD", Err)

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142744)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Posiciona_TotalReservado()
 
Dim iIndice As Integer

    'Posiciona o Total Reservado e o label correspondente
    TotalReservado.Top = GridReserva.Top + GridReserva.Height
    TotalReservado.Left = GridReserva.Left
    For iIndice = 0 To iGrid_QuantReservada_Col - 1
        TotalReservado.Left = TotalReservado.Left + GridReserva.ColWidth(iIndice) + GridReserva.GridLineWidth + 20
    Next

    TotalReservado.Width = GridReserva.ColWidth(iGrid_QuantReservada_Col)
    QuantTotalReserva.Top = TotalReservado.Top + (TotalReservado.Height - QuantTotalReserva.Height) / 2
    QuantTotalReserva.Left = TotalReservado.Left - QuantTotalReserva.Width

End Function

Private Function Inicializa_Grid_Reserva(objGridInt As AdmGrid, iNumLinhas As Integer) As Long
'Inicializa o Grid

Dim iIndice As Integer

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("Disponível")
    objGridInt.colColuna.Add ("Reservada")
    objGridInt.colColuna.Add ("Data Validade")
    objGridInt.colColuna.Add ("Responsável")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Almoxarifado.Name)
    objGridInt.colCampo.Add (QuantDisponivel.Name)
    objGridInt.colCampo.Add (QuantReservada.Name)
    objGridInt.colCampo.Add (DataValidade.Name)
    objGridInt.colCampo.Add (Responsavel.Name)

    'Colunas do Grid
    iGrid_Sequencial_Col = 0
    iGrid_Almoxarifado_Col = 1
    iGrid_QuantDisponivel_Col = 2
    iGrid_QuantReservada_Col = 3
    iGrid_DataValidade_Col = 4
    iGrid_Responsavel_Col = 5

    'Grid do GridInterno
    objGridInt.objGrid = GridReserva

    'Todas as linhas do grid
    If iNumLinhas + 1 > 7 Then
        objGridInt.objGrid.Rows = iNumLinhas + 1
    Else
        objGridInt.objGrid.Rows = 8
    End If

    objGridInt.iLinhasVisiveis = 7
    
    'Largura da primeira coluna
    GridReserva.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Call Posiciona_TotalReservado

    Inicializa_Grid_Reserva = SUCESSO

    Exit Function

End Function

Function Trata_Parametros(Optional colItens As ColItemPedido) As Long

Dim lErro As Long
Dim objItemPedido As ClassItemPedido
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    Set gcolItemPedido = colItens

    'Verifica se houve passagem de parametro
    If Not (gcolItemPedido Is Nothing) Then

        'Coloca itens pedidos na Combo Item
        For iIndice = 1 To gcolItemPedido.Count
            Item.AddItem iIndice
        Next

    End If

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142745)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Set objGrid2 = Nothing
    
    Set gcolItemPedido = Nothing
    Set gcolEstoque = Nothing

End Sub

Private Sub GridReserva_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid2, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid2, iAlterado)
    End If

End Sub

Private Sub GridReserva_EnterCell()

    Call Grid_Entrada_Celula(objGrid2, iAlterado)

End Sub

Private Sub GridReserva_GotFocus()

    Call Grid_Recebe_Foco(objGrid2)

End Sub

Private Sub GridReserva_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGrid2)

End Sub

Private Sub GridReserva_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid2, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid2, iAlterado)
    End If

End Sub

Private Sub GridReserva_LeaveCell()

    Call Saida_Celula(objGrid2)

End Sub

Private Sub GridReserva_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGrid2)

End Sub

Private Sub GridReserva_RowColChange()

    Call Grid_RowColChange(objGrid2)

End Sub

Private Sub GridReserva_Scroll()

    Call Grid_Scroll(objGrid2)

End Sub

Private Sub Item_Click()

Dim lErro As Long
Dim vbMsg As VbMsgBoxResult
Dim sProdutoMascarado As String
Dim objItemRomaneio As ClassItemRomaneioGrade

On Error GoTo Erro_Item_Click

    If Item.ListIndex = -1 Then
        'Limpa a COmbo de produtos
        Produto.Clear
        'Sai
        Exit Sub
    End If
    
    'Verificar se o item anterior sofreu alguma alteração
    If iAlterado = REGISTRO_ALTERADO Then
        'Verificar se o usuário deseja salvar as alterações
        vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ITEM_ANTERIOR_ALTERADO")
                
        If vbMsg = vbYes Then
            bGravandoProduto = True
            Call BotaoGravar_Click
            bGravandoProduto = False
        End If

    End If

    'Limpa a COmbo de produtos
    Produto.Clear

    If StrParaInt(Item.Text) = 0 Then Exit Sub
    
    'Se o produto for de Grade
    If gcolItemPedido(CInt(Item.Text)).iPossuiGrade = MARCADO Then
        
        BotaoSubstituir.Enabled = False

        For Each objItemRomaneio In gcolItemPedido(CInt(Item.Text)).colItensRomaneioGrade
        
            'Preenche dados do Item Pedido
            lErro = Mascara_MascararProduto(objItemRomaneio.sProduto, sProdutoMascarado)
            If lErro <> SUCESSO Then Error 23800
            
            Produto.AddItem sProdutoMascarado
        
        Next

    Else
    
        'Preenche dados do Item Pedido
        lErro = Mascara_MascararProduto(gcolItemPedido(CInt(Item.Text)).sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then Error 23800
        
        Produto.AddItem sProdutoMascarado
    
    End If

    Produto.ListIndex = 0
    iProdutoAtual = 0
    
    iAlterado = 0

    Exit Sub

Erro_Item_Click:

    Select Case gErr

        Case 23800

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142746)

    End Select

    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iUltimaLinha As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        If objGridInt.objGrid = GridReserva Then

            Select Case objGridInt.objGrid.Col

                Case iGrid_QuantReservada_Col
                    lErro = Saida_Celula_QuantReservada(objGridInt)
                    If lErro <> SUCESSO Then Error 23745

                Case iGrid_Responsavel_Col
                    lErro = Saida_Celula_Responsavel(objGridInt)
                    If lErro <> SUCESSO Then Error 23746

                Case iGrid_DataValidade_Col
                    lErro = Saida_Celula_DataValidade(objGridInt)
                    If lErro <> SUCESSO Then Error 23747

            End Select

        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 23748

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 23745, 23746, 23747

        Case 23748
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Responsavel(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid Responsável que está deixando de ser a corrente

Dim lErro As Long
Dim dColunaSoma As Double

On Error GoTo Erro_Saida_Celula_Responsavel

    Set objGridInt.objControle = Responsavel

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 23757

    Saida_Celula_Responsavel = SUCESSO

    Exit Function

Erro_Saida_Celula_Responsavel:

    Saida_Celula_Responsavel = Err

    Select Case Err

        Case 23757
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142747)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantReservada(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Quantidade Reservada que está deixando de ser a corrente

Dim lErro As Long
Dim dColunaSoma As Double
Dim dTotalReservado As Double
Dim dTolerancia As Double

On Error GoTo Erro_Saida_Celula_QuantReservada

    Set objGridInt.objControle = QuantReservada

    'Verifica se valor está preenchido
    If Len(Trim(QuantReservada.Text)) > 0 Then
    
        'Critica se valor é não negativo
        lErro = Valor_NaoNegativo_Critica(QuantReservada.Text)
        If lErro <> SUCESSO Then Error 23752

        'Quantidade reservada não pode ser maior que a quantidade disponível
        If StrParaDbl(QuantReservada.Text) > StrParaDbl(GridReserva.TextMatrix(GridReserva.Row, iGrid_QuantDisponivel_Col)) Then Error 23753

    End If

    Call SomaReservado(dTotalReservado)
    
'    lErro = Critica_QuantReservada(dTotalReservado, StrParaDbl(QuantReservar.Caption))
'    If lErro <> SUCESSO Then Error 51423
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 23756

    Saida_Celula_QuantReservada = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantReservada:

    Saida_Celula_QuantReservada = Err

    Select Case Err

        Case 23752, 51423
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 23753
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTRESERVADA_MAIOR", Err, CDbl(QuantReservada.Text))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 23755
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TOTAL_RESERVADO_MAIOR", Err, CDbl(TotalReservado.Caption))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 23756

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142748)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataValidade(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Data Validade que está deixando de ser a corrente

Dim lErro As Long
Dim dtDataValidade As Date

On Error GoTo Erro_Saida_Celula_DataValidade

    Set objGridInt.objControle = DataValidade

    'Verifica se Data de Validade esta preenchida
    If Len(Trim(DataValidade.ClipText)) > 0 Then

        'Critica a data
        lErro = Data_Critica(DataValidade.Text)
        If lErro <> SUCESSO Then Error 23749

         dtDataValidade = CDate(DataValidade.Text)

        'Se data de Validade é menor que a Data Corrente
        If dtDataValidade < Date Then Error 23750

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 23751

    Saida_Celula_DataValidade = SUCESSO

    Exit Function

Erro_Saida_Celula_DataValidade:

    Saida_Celula_DataValidade = Err

    Select Case Err

        Case 23749, 23751
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 23750
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVALIDADE_MENOR", Err, DataValidade.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142749)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objGenerico As New AdmGenerico

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = Critica_QuantReservada(StrParaDbl(TotalReservado.Caption), StrParaDbl(QuantReservar.Caption))
    If lErro <> SUCESSO Then Error 60771
    
    'Verificar se o Total Reservado é igual a Quantidade a reservar-->
    '--> se não chama tela AlocacaoProdutoSaida
    If TotalReservado.Caption <> QuantReservar.Caption Then
        '????
        If StrParaDbl(TotalReservado) < StrParaDbl(QuantReservar) Then
            Call Chama_Tela_Modal("AlocacaoProdutoSaida", objGenerico)
        Else
            giRetornoTela = vbOK
        End If
        
        If giRetornoTela = vbOK Then
        
            lErro = Atualiza_Dados_Reserva(objGenerico)
            If lErro <> SUCESSO Then Error 23938
            
            iAlterado = 0
            
        End If
        
        Gravar_Registro = SUCESSO
        
        GL_objMDIForm.MousePointer = vbDefault
        
        Exit Function
        
    End If
    
    lErro = Move_Tela_Memoria(gcolItemPedido)
    If lErro <> SUCESSO Then Error 23765
    
    iAlterado = 0

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 23765, 60771

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 142750)

    End Select

    Exit Function

End Function

Private Sub Produto_Click()

Dim iIndice As Integer
Dim dQuantReservadaPedido As Double
Dim colReservaItemBD As New colReservaItem
Dim objEstoque As ClassEstoqueProduto
Dim objReservaItemBD As ClassReservaItem
Dim vbMsg As VbMsgBoxResult
Dim iNumEstoques As Integer
Dim objItemPV As New ClassItemPedido
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim lErro As Long
Dim objReservaItem As ClassReservaItem
Dim objProduto As New ClassProduto

On Error GoTo Erro_Produto_Click

    bGravandoProduto = False
    
    'Verificar se o item anterior sofreu alguma alteração
    If iProdutoAtual <> -1 And iAlterado = REGISTRO_ALTERADO Then
        'Verificar se o usuário deseja salvar as alterações
        vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ITEM_ANTERIOR_ALTERADO")
                
        If vbMsg = vbYes Then
            bGravandoProduto = True
            Call BotaoGravar_Click
            bGravandoProduto = False
        End If

    End If
    
    iProdutoAtual = Produto.ListIndex
    
    Set gcolEstoque = New colEstoqueProduto

    If gcolItemPedido(CInt(Item.Text)).iPossuiGrade = MARCADO Then
    
        objProduto.sCodigo = gcolItemPedido(CInt(Item.Text)).sProduto
    
        'Lê o Produto que está sendo Passado
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 23741

        objItemPV.lCodPedido = gcolItemPedido(CInt(Item.Text)).lCodPedido
        objItemPV.iFilialEmpresa = gcolItemPedido(CInt(Item.Text)).iFilialEmpresa
        
        Set objItemRomaneio = gcolItemPedido(CInt(Item.Text)).colItensRomaneioGrade(iProdutoAtual + 1)
        
        objItemPV.sProduto = objItemRomaneio.sProduto
        objItemPV.dQuantidade = objItemRomaneio.dQuantidade
        objItemPV.dQuantReservada = objItemRomaneio.dQuantReservada
        objItemPV.dQuantCancelada = objItemRomaneio.dQuantCancelada
        objItemPV.dQuantFaturada = objItemRomaneio.dQuantFaturada
        objItemPV.iItem = CInt(Item.Text)
        objItemPV.lNumIntDoc = objItemRomaneio.lNumIntDoc
        objItemPV.sUMEstoque = objItemRomaneio.sUMEstoque
        objItemPV.sProdutoDescricao = objItemRomaneio.sDescricao
        If objProduto.iKitVendaComp = MARCADO Then
            objItemPV.sUnidadeMed = objItemRomaneio.sUMEstoque
            objItemPV.iClasseUM = 0
        Else
            objItemPV.sUnidadeMed = gcolItemPedido(CInt(Item.Text)).sUnidadeMed
            objItemPV.iClasseUM = gcolItemPedido(CInt(Item.Text)).iClasseUM
        End If
        objItemPV.iFilialEmpresa = gcolItemPedido(CInt(Item.Text)).iFilialEmpresa
        objItemPV.lCodPedido = gcolItemPedido(CInt(Item.Text)).lCodPedido
        objItemPV.iPossuiGrade = MARCADO
        
        For Each objReservaItem In objItemRomaneio.colLocalizacao
                    
            objItemPV.colReserva.Add objItemPV.iFilialEmpresa, objItemPV.lCodPedido, objItemPV.sProduto, objReservaItem.iAlmoxarifado, 0, objItemPV.lCodPedido, objItemRomaneio.lNumIntDoc, objReservaItem.dQuantidade, DATA_NULA, objReservaItem.dtDataValidade, "", objReservaItem.sResponsavel, 0, objReservaItem.sAlmoxarifado
                    
        Next
        

    Else

        Set objItemPV = gcolItemPedido(CInt(Item.Text))

    End If

    'Lê nas tabelas de EstoqueProduto e Almoxarifado as posições de Estoque do Item
    lErro = CF("EstoquesProduto_Le", objItemPV.sProduto, gcolEstoque)
    If lErro <> SUCESSO And lErro <> 30100 Then gError 23741

    'Nenhum dos almoxarifados tem quantidade para este Produto
    If lErro = 30100 Then gError 64431

    'Lê nas tabelas de Reserva e Almoxarifado as reservas do item
    lErro = CF("ReservasItem_Le", objItemPV, colReservaItemBD)
    If lErro <> 30099 And lErro <> SUCESSO Then gError 23742

    'Verifica se existe reserva correspondente no BD
    For Each objEstoque In gcolEstoque
        dQuantReservadaPedido = 0
        For Each objReservaItemBD In colReservaItemBD
            If objEstoque.iAlmoxarifado = objReservaItemBD.iAlmoxarifado Then
                dQuantReservadaPedido = objReservaItemBD.dQuantidade
                Exit For
            End If
        Next
        objEstoque.dSaldo = objEstoque.dQuantDisponivel + dQuantReservadaPedido
    Next


    'Se não existir disponibilidade do produto --> Erro
    If gcolEstoque.Count = 0 Then gError 23743

    'Retira da coleção os que tem saldo zero
    For iIndice = gcolEstoque.Count To 1 Step -1
        If gcolEstoque.Item(iIndice).dSaldo = 0 Then gcolEstoque.Remove iIndice
    Next

    Call GridReserva_Limpa

    lErro = Preenche_Tela(objItemPV)
    If lErro <> SUCESSO Then gError 23744

    giListIndexAnterior = Item.ListIndex

    iAlterado = 0

    Exit Sub

Erro_Produto_Click:

    Select Case gErr

        Case 23741, 23742

        Case 23743
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_DISPONIVEL", gErr, gcolItemPedido(CInt(Item.Text)).sProduto)

        Case 23744

        Case 64431
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NAO_EXISTE_ESTOQUE", gErr, objItemPV.sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142751)

    End Select

    Exit Sub

End Sub

Private Sub Produto_GotFocus()
    iProdutoAtual = Produto.ListIndex
End Sub

Private Sub QuantReservada_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantReservada_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid2)

End Sub

Private Sub QuantReservada_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid2)

End Sub

Private Sub QuantReservada_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid2.objControle = QuantReservada
    lErro = Grid_Campo_Libera_Foco(objGrid2)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Responsavel_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Responsavel_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid2)

End Sub

Private Sub Responsavel_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid2)

End Sub

Private Sub Responsavel_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid2.objControle = Responsavel
    lErro = Grid_Campo_Libera_Foco(objGrid2)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DataValidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataValidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid2)

End Sub

Private Sub DataValidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid2)

End Sub

Private Sub DataValidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid2.objControle = DataValidade
    lErro = Grid_Campo_Libera_Foco(objGrid2)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Function GridReserva_Limpa()

    'Limpa o Grid de Reservas
    Call Grid_Limpa(objGrid2)

    'Limpa label TotalReservado
    TotalReservado.Caption = Formata_Estoque(0)

End Function

Private Function Preenche_Tela(objItemPedido As ClassItemPedido) As Long
'Mostra os dados do Item Pedido na tela

Dim lErro As Long
Dim iLinha As Integer
Dim iIndice As Integer
Dim dFator As Double
Dim objEstoqueProduto As ClassEstoqueProduto

On Error GoTo Erro_Preenche_Tela
    
    Descricao.Caption = objItemPedido.sProdutoDescricao
    UnidadeMedida.Caption = objItemPedido.sUMEstoque

    If objItemPedido.iClasseUM <> 0 Then
        'Calcular fator de conversão de UM de Venda para Estoque
        lErro = CF("UM_Conversao", objItemPedido.iClasseUM, objItemPedido.sUnidadeMed, objItemPedido.sUMEstoque, dFator)
        If lErro <> SUCESSO Then Error 23769
    Else
        dFator = 1
    End If

    'Preenche QuantReservar
    QuantReservar.Caption = Formata_Estoque((objItemPedido.dQuantidade - objItemPedido.dQuantCancelada - objItemPedido.dQuantFaturada) * dFator)

    'Linhas Existentes
    objGrid2.iLinhasExistentes = gcolEstoque.Count

    'Preenche GridReserva
    For Each objEstoqueProduto In gcolEstoque

        iLinha = iLinha + 1

        GridReserva.TextMatrix(iLinha, iGrid_Almoxarifado_Col) = objEstoqueProduto.sAlmoxarifadoNomeReduzido
        GridReserva.TextMatrix(iLinha, iGrid_QuantDisponivel_Col) = Formata_Estoque(objEstoqueProduto.dSaldo)
        
        For iIndice = 1 To objItemPedido.colReserva.Count
            If objItemPedido.colReserva(iIndice).iAlmoxarifado = objEstoqueProduto.iAlmoxarifado Then
                GridReserva.TextMatrix(iLinha, iGrid_QuantReservada_Col) = Formata_Estoque(objItemPedido.colReserva(iIndice).dQuantidade)
                GridReserva.TextMatrix(iLinha, iGrid_Responsavel_Col) = objItemPedido.colReserva(iIndice).sResponsavel
                
                If objItemPedido.colReserva(iIndice).dtDataValidade <> DATA_NULA Then
                    GridReserva.TextMatrix(iLinha, iGrid_DataValidade_Col) = CStr(objItemPedido.colReserva(iIndice).dtDataValidade)
                    
                End If
                
                Exit For

            End If
        Next

    Next

    'Soma o total reservado
    dTotalReservado = 0
    For iIndice = 1 To objGrid2.iLinhasExistentes

        If Len(Trim(GridReserva.TextMatrix(iIndice, iGrid_QuantReservada_Col))) > 0 Then
            dTotalReservado = dTotalReservado + CDbl(GridReserva.TextMatrix(iIndice, iGrid_QuantReservada_Col))
        End If

    Next

    TotalReservado.Caption = Formata_Estoque(dTotalReservado)

    Preenche_Tela = SUCESSO

    Exit Function

Erro_Preenche_Tela:

    Preenche_Tela = Err

    Select Case Err

        Case 23769, 23800

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142752)

    End Select

    Exit Function

End Function

Private Function SomaReservado(dTotalReservado As Double)
'Soma a coluna de quantidade reservada e coloca o total em TotalReservado

Dim iIndice As Integer

    dTotalReservado = 0
    For iIndice = 1 To objGrid2.iLinhasExistentes

        If Len(Trim(GridReserva.TextMatrix(iIndice, iGrid_QuantReservada_Col))) > 0 Then
            dTotalReservado = dTotalReservado + CDbl(GridReserva.TextMatrix(iIndice, iGrid_QuantReservada_Col))
        End If

    Next
    
    'Somar a célula ativa
    If Len(Trim(QuantReservada.Text)) > 0 Then dTotalReservado = dTotalReservado + CDbl(QuantReservada.Text)
    If Len(Trim(GridReserva.TextMatrix(GridReserva.Row, iGrid_QuantReservada_Col))) > 0 And GridReserva.Row <> 0 Then
        dTotalReservado = dTotalReservado - CDbl(GridReserva.TextMatrix(GridReserva.Row, iGrid_QuantReservada_Col))
    End If

    'Armazena dTotalReservado no label TotalReservado
    'com o formato correto
    TotalReservado.Caption = Formata_Estoque(dTotalReservado)

End Function

Private Function Preenche_Tela2(objProduto As ClassProduto, gcolEstoque As colEstoqueProduto) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim sProdutoMascarado As String
Dim objEstoqueProduto As ClassEstoqueProduto

On Error GoTo Erro_Preenche_Tela2
    
    lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then Error 23801
    
    Descricao.Caption = objProduto.sDescricao
    UnidadeMedida.Caption = objProduto.sSiglaUMEstoque

    'Linhas Existentes
    objGrid2.iLinhasExistentes = gcolEstoque.Count

    For Each objEstoqueProduto In gcolEstoque

        iLinha = iLinha + 1

        GridReserva.TextMatrix(iLinha, iGrid_Almoxarifado_Col) = objEstoqueProduto.sAlmoxarifadoNomeReduzido
        GridReserva.TextMatrix(iLinha, iGrid_QuantDisponivel_Col) = Formata_Estoque(objEstoqueProduto.dSaldo)

    Next

    iAlterado = REGISTRO_ALTERADO

    Preenche_Tela2 = SUCESSO

    Exit Function

Erro_Preenche_Tela2:

    Preenche_Tela2 = Err

    Select Case Err

        Case 23801
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 142753)

    End Select

    Exit Function

End Function

Public Function Move_Tela_Memoria(gcolItemPedido As ColItemPedido) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim dFator As Double
Dim iItem As Integer
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objReservaItem As ClassReservaItem

On Error GoTo Erro_Move_Tela_Memoria
    
    iItem = giListIndexAnterior + 1
    
    If gcolItemPedido(iItem).iPossuiGrade = MARCADO Then
    
        Set objItemRomaneio = gcolItemPedido(iItem).colItensRomaneioGrade(iProdutoAtual + 1)
    
        Set objItemRomaneio.colLocalizacao = New Collection
        
        objItemRomaneio.dQuantReservada = 0
    
        'Preenche a coleção com as linhas do grid -->
        For iLinha = 1 To objGrid2.iLinhasExistentes
    
            '--> Se a quantidade reservada for positiva
            If Len(Trim(GridReserva.TextMatrix(iLinha, iGrid_QuantReservada_Col))) > 0 Then
                If CDbl(GridReserva.TextMatrix(iLinha, iGrid_QuantReservada_Col)) > 0 Then
                    
                    Set objReservaItem = New ClassReservaItem
                 
                    objReservaItem.dQuantidade = CDbl(GridReserva.TextMatrix(iLinha, iGrid_QuantReservada_Col))
                    objReservaItem.dtDataValidade = StrParaDate(GridReserva.TextMatrix(iLinha, iGrid_DataValidade_Col))
                    objReservaItem.iAlmoxarifado = gcolEstoque(iLinha).iAlmoxarifado
                    objReservaItem.sAlmoxarifado = GridReserva.TextMatrix(iLinha, iGrid_Almoxarifado_Col)
                    objReservaItem.sResponsavel = GridReserva.TextMatrix(iLinha, iGrid_Responsavel_Col)
                    
                    objItemRomaneio.dQuantReservada = objItemRomaneio.dQuantReservada + objReservaItem.dQuantidade
                    
                    objItemRomaneio.colLocalizacao.Add objReservaItem
                End If
            End If
    
        Next
    
        'Calcular o fator de conversão de UM de Estoque para Venda
        lErro = CF("UM_Conversao", gcolItemPedido(iItem).iClasseUM, gcolItemPedido(iItem).sUMEstoque, gcolItemPedido(iItem).sUnidadeMed, dFator)
        If lErro <> SUCESSO Then Error 23777
        
        objItemRomaneio.dQuantReservada = objItemRomaneio.dQuantReservada * dFator
        
        'Preenche a Quantidade Reservada com o Total Reservado
        gcolItemPedido(iItem).dQuantReservada = 0
        For Each objItemRomaneio In gcolItemPedido(iItem).colItensRomaneioGrade
            gcolItemPedido(iItem).dQuantReservada = gcolItemPedido(iItem).dQuantReservada + objItemRomaneio.dQuantReservada
        Next
    
    Else
        'Remove todos os elementos da coleção
        Set gcolItemPedido(iItem).colReserva = New colReserva
    
        'Preenche a coleção com as linhas do grid -->
        For iLinha = 1 To objGrid2.iLinhasExistentes
    
            '--> Se a quantidade reservada for positiva
            If Len(Trim(GridReserva.TextMatrix(iLinha, iGrid_QuantReservada_Col))) > 0 Then
                If CDbl(GridReserva.TextMatrix(iLinha, iGrid_QuantReservada_Col)) > 0 Then
                    If Len(Trim(GridReserva.TextMatrix(iLinha, iGrid_DataValidade_Col))) > 0 Then
                        gcolItemPedido(iItem).colReserva.Add 0, 0, "", gcolEstoque(iLinha).iAlmoxarifado, 0, 0, 0, CDbl(GridReserva.TextMatrix(iLinha, iGrid_QuantReservada_Col)), DATA_NULA, CDate(GridReserva.TextMatrix(iLinha, iGrid_DataValidade_Col)), "", GridReserva.TextMatrix(iLinha, iGrid_Responsavel_Col), 0, GridReserva.TextMatrix(iLinha, iGrid_Almoxarifado_Col)
                    Else
                        gcolItemPedido(iItem).colReserva.Add 0, 0, "", gcolEstoque(iLinha).iAlmoxarifado, 0, 0, 0, CDbl(GridReserva.TextMatrix(iLinha, iGrid_QuantReservada_Col)), DATA_NULA, DATA_NULA, "", GridReserva.TextMatrix(iLinha, iGrid_Responsavel_Col), 0, GridReserva.TextMatrix(iLinha, iGrid_Almoxarifado_Col)
                    End If
                End If
            End If
    
        Next
    
        'Calcular o fator de conversão de UM de Estoque para Venda
        lErro = CF("UM_Conversao", gcolItemPedido(iItem).iClasseUM, gcolItemPedido(iItem).sUMEstoque, gcolItemPedido(iItem).sUnidadeMed, dFator)
        If lErro <> SUCESSO Then Error 23777
    
        'Preenche a Quantidade Reservada com o Total Reservado
        gcolItemPedido(iItem).dQuantReservada = StrParaDbl(TotalReservado.Caption) * dFator
    
    End If

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

    Case 23777

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 142754)

    End Select

    Exit Function

End Function

Private Sub Limpa_Tela_Alocacao()

Dim iIndice As Integer


    Item.ListIndex = -1
    
    Call Grid_Limpa(objGrid2)
    
    Produto.Clear
    Descricao.Caption = ""
    UnidadeMedida.Caption = ""
    QuantReservar.Caption = ""
    TotalReservado.Caption = ""
    
'
'    'Limpa Colunas do GridReserva (QuantReservada , Responsável e Data de Validade)
'    For iIndice = 1 To objGrid2.iLinhasExistentes
'
'        GridReserva.TextMatrix(iIndice, iGrid_DataValidade_Col) = ""
'        GridReserva.TextMatrix(iIndice, iGrid_QuantReservada_Col) = ""
'        GridReserva.TextMatrix(iIndice, iGrid_Responsavel_Col) = ""
'
'    Next

    'Limpa label TotalReservado

End Sub

Function Atualiza_Dados_Reserva(objGenerico As AdmGenerico) As Long

Dim lErro As Long
Dim dFator As Double
Dim iLinha As Integer
Dim dQuantCancelada As Double

On Error GoTo Erro_Atualiza_Dados_Reserva

    Select Case objGenerico.vVariavel
    
        Case NENHUMA_SELECAO
        
        Case SELECAO_OK
            lErro = Move_Tela_Memoria(gcolItemPedido)
            If lErro <> SUCESSO Then Error 23939
            
        Case CANCELA_ACIMA_DA_RESERVADA
            lErro = Move_Tela_Memoria(gcolItemPedido)
            If lErro <> SUCESSO Then Error 25189
            
            If gcolItemPedido(giListIndexAnterior + 1).iPossuiGrade = MARCADO Then
                
                Dim objItemRomaneio As ClassItemRomaneioGrade
                
                Set objItemRomaneio = gcolItemPedido(giListIndexAnterior + 1).colItensRomaneioGrade(iProdutoAtual + 1)
                
                objItemRomaneio.dQuantCancelada = objItemRomaneio.dQuantCancelada + StrParaDbl(QuantReservar.Caption) - StrParaDbl(TotalReservado.Caption)
            
            Else
                '???
                lErro = CF("UM_Conversao", gcolItemPedido(giListIndexAnterior + 1).iClasseUM, gcolItemPedido(giListIndexAnterior + 1).sUMEstoque, gcolItemPedido(giListIndexAnterior + 1).sUnidadeMed, dFator)
                If lErro <> SUCESSO Then Error 23940
                            
                dQuantCancelada = (gcolItemPedido(giListIndexAnterior + 1).dQuantidade - gcolItemPedido(giListIndexAnterior + 1).dQuantFaturada) - gcolItemPedido(giListIndexAnterior + 1).dQuantAFaturar
                
                gcolItemPedido(giListIndexAnterior + 1).dQuantCancelada = dQuantCancelada + (CDbl(QuantReservar.Caption) - CDbl(TotalReservado.Caption)) * dFator
            End If
                   
        Case NAO_RESERVAR_PRODUTO
        
            'Limpa as colunas de QuantReservada e Responsavel no GridReserva
            For iLinha = 1 To objGrid2.iLinhasExistentes
                GridReserva.TextMatrix(iLinha, iGrid_QuantReservada_Col) = ""
                GridReserva.TextMatrix(iLinha, iGrid_Responsavel_Col) = ""
            Next

            lErro = Move_Tela_Memoria(gcolItemPedido)
            If lErro <> SUCESSO Then Error 23941
        
    End Select
    
    Atualiza_Dados_Reserva = SUCESSO
    
    Exit Function
    
Erro_Atualiza_Dados_Reserva:

    Atualiza_Dados_Reserva = Err
    
    Select Case Err
    
        Case 23939, 25189, 23940, 23941
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142755)

    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_ALOCACAO_PRODUTO
    Set Form_Load_Ocx = Me
    Caption = "Reserva de Produto"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "AlocacaoProduto"
    
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

Function Calcula_Tolerancia(iClassUM As Integer, sUMVenda As String, sUMEstoque As String, dTolerancia As Double) As Long
'???
Dim lErro As Long
Dim dFator As Double
Dim iNumCasasDec As Integer

On Error GoTo Erro_Calcula_Tolerancia

    'Calcula o número de casas decimais do Formato de Estoque
    iNumCasasDec = Len(Mid(FORMATO_ESTOQUE, (InStr(FORMATO_ESTOQUE, ".")) + 1))

    'Calcula o Fator de conversão entre a UM de Venda e de Estoque
    lErro = CF("UM_Conversao", iClassUM, sUMVenda, sUMEstoque, dFator)
    If lErro <> SUCESSO Then Error 51422
    
    'Se as UMs forem iguais
    If dFator = 1 Then
        dTolerancia = 0
    
    'Se a unidade de venda for maior que a de estoque
    ElseIf dFator > 1 Then
        dTolerancia = (10 ^ -iNumCasasDec) * dFator
    
    'Se a unidade de venda for menor que a de estoque
    Else
        dTolerancia = (10 ^ -iNumCasasDec)
    
    End If
    
    Calcula_Tolerancia = SUCESSO
    
    Exit Function
    
Erro_Calcula_Tolerancia:

    Calcula_Tolerancia = Err
    
    Select Case Err
    
        Case 51422
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142756)
            
    End Select
    
    Exit Function

End Function

Function Critica_QuantReservada(dTotalReservado As Double, dQuantReservar As Double) As Long
'???
Dim lErro As Long
Dim dTolerancia As Double

On Error GoTo Erro_Critica_QuantReservada

    'Se o total reservado for maior que o a reservar
    If Formata_Estoque(dTotalReservado) > Formata_Estoque(StrParaDbl(QuantReservar.Caption)) Then
        'Calcula a tolerância de acordo com as UMs de Venda e de Estoque
        lErro = Calcula_Tolerancia(gcolItemPedido(StrParaInt(Item.Text)).iClasseUM, gcolItemPedido(StrParaInt(Item.Text)).sUnidadeMed, UnidadeMedida.Caption, dTolerancia)
        If lErro <> SUCESSO Then Error 51424
        'Verifica se a quantidade resrvada ultrapassa a quantidade a reservar com a tolerância
        If StrParaDbl(Formata_Estoque(dTotalReservado)) > StrParaDbl(Formata_Estoque(StrParaDbl(QuantReservar.Caption) + dTolerancia)) Then Error 23754
    
    End If
    
    Critica_QuantReservada = SUCESSO
    
    Exit Function
    
Erro_Critica_QuantReservada:

    Critica_QuantReservada = Err
    
    Select Case Err
    
        Case 51424
        
        Case 23754
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TOTAL_RESERVADO_MAIOR", Err, CDbl(TotalReservado.Caption))
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142757)
    
    End Select
    
    Exit Function
    
End Function

Private Sub TotalReservado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalReservado, Source, X, Y)
End Sub

Private Sub TotalReservado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalReservado, Button, Shift, X, Y)
End Sub

Private Sub QuantTotalReserva_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantTotalReserva, Source, X, Y)
End Sub

Private Sub QuantTotalReserva_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantTotalReserva, Button, Shift, X, Y)
End Sub

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


Private Sub Descricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Descricao, Source, X, Y)
End Sub

Private Sub Descricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Descricao, Button, Shift, X, Y)
End Sub

Private Sub QuantReservar_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantReservar, Source, X, Y)
End Sub

Private Sub QuantReservar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantReservar, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub UnidadeMedida_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(UnidadeMedida, Source, X, Y)
End Sub

Private Sub UnidadeMedida_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(UnidadeMedida, Button, Shift, X, Y)
End Sub


