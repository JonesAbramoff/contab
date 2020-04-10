VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl AlocacaoProduto1Ocx 
   ClientHeight    =   4965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8175
   LockControls    =   -1  'True
   ScaleHeight     =   4965
   ScaleWidth      =   8175
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6405
      ScaleHeight     =   495
      ScaleWidth      =   1560
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   135
      Width           =   1620
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1065
         Picture         =   "AlocacaoProduto1Ocx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   570
         Picture         =   "AlocacaoProduto1Ocx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   75
         Picture         =   "AlocacaoProduto1Ocx.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Reserva do Produto"
      Height          =   2760
      Left            =   195
      TabIndex        =   10
      Top             =   2040
      Width           =   7770
      Begin MSMask.MaskEdBox QuantReservada 
         Height          =   225
         Left            =   3345
         TabIndex        =   2
         Top             =   375
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
      Begin VB.TextBox Responsavel 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4800
         MaxLength       =   50
         TabIndex        =   4
         Top             =   660
         Width           =   1965
      End
      Begin MSMask.MaskEdBox DataValidade 
         Height          =   225
         Left            =   4470
         TabIndex        =   3
         Top             =   345
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
         Left            =   780
         TabIndex        =   0
         Top             =   390
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
         Left            =   2220
         TabIndex        =   1
         Top             =   315
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
      Begin MSFlexGridLib.MSFlexGrid GridReserva 
         Height          =   1860
         Left            =   210
         TabIndex        =   5
         Top             =   300
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   3281
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
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
         TabIndex        =   12
         Top             =   2295
         Width           =   1620
      End
      Begin VB.Label TotalReservado 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3375
         TabIndex        =   11
         Top             =   2235
         Width           =   1440
      End
   End
   Begin VB.Label Descricao 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1830
      TabIndex        =   20
      Top             =   720
      Width           =   3075
   End
   Begin VB.Label Produto 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1830
      TabIndex        =   19
      Top             =   300
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
      Left            =   1020
      TabIndex        =   18
      Top             =   330
      Width           =   735
   End
   Begin VB.Label QuantReservar 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1830
      TabIndex        =   17
      Top             =   1575
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
      TabIndex        =   16
      Top             =   1605
      Width           =   1635
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "U.M.:"
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
      Left            =   1275
      TabIndex        =   15
      Top             =   1185
      Width           =   480
   End
   Begin VB.Label UnidadeMedida 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1830
      TabIndex        =   14
      Top             =   1155
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descrição:"
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
      Left            =   825
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   13
      Top             =   780
      Width           =   930
   End
End
Attribute VB_Name = "AlocacaoProduto1Ocx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Declaração de variáveis globais
Dim iAlterado As Integer
Dim gobjItemPedido As ClassItemPedido
Dim dTotalReservado As Double
Dim gcolEstoque As colEstoqueProduto

Dim objGrid2 As AdmGrid
Dim iGrid_Sequencial_Col As Integer
Dim iGrid_Almoxarifado_Col As Integer
Dim iGrid_QuantDisponivel_Col As Integer
Dim iGrid_QuantReservada_Col As Integer
Dim iGrid_DataValidade_Col As Integer
Dim iGrid_Responsavel_Col As Integer

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 23708

    Unload Me

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 23708

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142729)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela_Alocacao

    iAlterado = REGISTRO_ALTERADO

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

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    QuantReservada.Format = FORMATO_ESTOQUE

    giRetornoTela = vbCancel

    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142730)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objItemPedido As ClassItemPedido) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iNumEstoques As Integer
Dim sProduto As String
Dim dQuantReservadaPedido As Double
Dim colReservaBD As colReservaItem
Dim objEstoqueProduto As ClassEstoqueProduto
Dim objReservaItem As ClassReservaItem

On Error GoTo Erro_Trata_Parametros

    Set gobjItemPedido = objItemPedido
    Set colReservaBD = New colReservaItem
    Set gcolEstoque = New colEstoqueProduto
    
    dTotalReservado = 0

    'Verifica se houve passagem de parametro
    If Not (objItemPedido Is Nothing) Then

        lErro = CF("EstoquesProduto_Le",objItemPedido.sProduto, gcolEstoque)
        If lErro <> SUCESSO And lErro <> 30100 Then gError 23698
        
        'Nenhum dos almoxarifados tem quantidade para este Produto
        If lErro = 30100 Then gError 64426
        
        lErro = CF("ReservasItem_Le",objItemPedido, colReservaBD)
        If lErro <> SUCESSO And lErro <> 30099 Then gError 23699

        'Verifica se existe reserva correspondente no BD
        For Each objEstoqueProduto In gcolEstoque
            dQuantReservadaPedido = 0
            For Each objReservaItem In colReservaBD
                If objEstoqueProduto.iAlmoxarifado = objReservaItem.iAlmoxarifado Then
                    dQuantReservadaPedido = objReservaItem.dQuantidade
                    Exit For
                End If
            Next
            objEstoqueProduto.dSaldo = objEstoqueProduto.dQuantDisponivel + dQuantReservadaPedido
        Next

        'Se não existir disponibilidade do produto --> Erro
        If gcolEstoque.Count = 0 Then gError 23729

        iIndice = 1

        'Filtrar saldo zero
        Do While iIndice <= gcolEstoque.Count
            If gcolEstoque.Item(iIndice).dSaldo = 0 Then
                gcolEstoque.Remove iIndice
                iIndice = iIndice - 1
            End If
            iIndice = iIndice + 1
        Loop
        
        Set objGrid2 = New AdmGrid
        
        'Inicializa Grid Reserva
        lErro = Inicializa_Grid_Reserva(objGrid2, gcolEstoque.Count + 1)
        If lErro <> SUCESSO Then gError 23700

        'Preenche a Tela
        lErro = Preenche_Tela(objItemPedido, gcolEstoque)
        If lErro <> SUCESSO Then gError 23711

    End If
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 23643, 23644, 23700, 23711

        Case 23729
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_DISPONIVEL", gErr, objItemPedido.sProduto)
            Unload Me
        
        Case 64426
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NAO_EXISTE_ESTOQUE", gErr, objItemPedido.sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142731)

    End Select

    Exit Function

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

    objGridInt.iLinhasVisiveis = 7
    If iNumLinhas > 7 Then
        objGridInt.objGrid.Rows = iNumLinhas + 1
    Else
        objGridInt.objGrid.Rows = 8
    End If

    'Linhas Existentes
    objGridInt.iLinhasExistentes = iNumLinhas - 1

    'Largura da primeira coluna
    GridReserva.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    'Posiciona o Total Reservado e o label correspondente
    TotalReservado.Top = GridReserva.Top + GridReserva.Height
    TotalReservado.Left = GridReserva.Left
    For iIndice = 0 To iGrid_QuantReservada_Col - 1
        TotalReservado.Left = TotalReservado.Left + GridReserva.ColWidth(iIndice) + GridReserva.GridLineWidth + 20
    Next

    TotalReservado.Width = GridReserva.ColWidth(iGrid_QuantReservada_Col)
    QuantTotalReserva.Top = TotalReservado.Top + (TotalReservado.Height - QuantTotalReserva.Height) / 2
    QuantTotalReserva.Left = TotalReservado.Left - QuantTotalReserva.Width - 20

    Inicializa_Grid_Reserva = SUCESSO

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
    If iAlterado = REGISTRO_ALTERADO Then giRetornoTela = vbCancel

End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Set objGrid2 = Nothing
    
    Set gobjItemPedido = Nothing
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

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iUltimaLinha As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        If objGridInt.objGrid = GridReserva Then

            Select Case objGridInt.objGrid.Col

                Case iGrid_QuantReservada_Col
                    lErro = Saida_Celula_QuantReservada(objGridInt)
                    If lErro <> SUCESSO Then Error 23701

                Case iGrid_Responsavel_Col
                    lErro = Saida_Celula_Responsavel(objGridInt)
                    If lErro <> SUCESSO Then Error 23702

                Case iGrid_DataValidade_Col
                    lErro = Saida_Celula_DataValidade(objGridInt)
                    If lErro <> SUCESSO Then Error 23733

            End Select

        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 23703

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 23701, 23702, 23733

        Case 23703
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Responsavel(objGridInt As AdmGrid) As Long
'Faz a crítica da celula Responsavel do grid

Dim lErro As Long
Dim dColunaSoma As Double

On Error GoTo Erro_Saida_Celula_Responsavel

    Set objGridInt.objControle = Responsavel

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 23704

    Saida_Celula_Responsavel = SUCESSO

    Exit Function

Erro_Saida_Celula_Responsavel:

    Saida_Celula_Responsavel = Err

    Select Case Err

        Case 23704
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142732)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantReservada(objGridInt As AdmGrid) As Long
'Faz a crítica da celula Valor do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dColunaSoma As Double
Dim dTotalReservado As Double

On Error GoTo Erro_Saida_Celula_QuantReservada

    Set objGridInt.objControle = QuantReservada

    'Verifica se valor está preenchido
    If Len(Trim(QuantReservada.Text)) > 0 Then
    
        'Critica se valor é não negativo
        lErro = Valor_NaoNegativo_Critica(QuantReservada.Text)
        If lErro <> SUCESSO Then Error 23705

        'Quantidade reservada não pode ser maior que a quantidade disponível
        If CDbl(QuantReservada.Text) > CDbl(GridReserva.TextMatrix(GridReserva.Row, iGrid_QuantDisponivel_Col)) Then Error 23706

    End If

    Call SomaReservado(dTotalReservado)
    lErro = Critica_QuantReservada(dTotalReservado, StrParaDbl(QuantReservar.Caption))
    If lErro <> SUCESSO Then Error 23707
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 23712

    Saida_Celula_QuantReservada = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantReservada:

    Saida_Celula_QuantReservada = Err

    Select Case Err

        Case 23705
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 23706
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTRESERVADA_MAIOR", Err, CDbl(QuantReservada.Text))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 23707
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 23712

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142733)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objGenerico As New AdmGenerico

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verificar se foi informada alguma quantidade para reserva
    If Len(Trim(TotalReservado.Caption)) > 0 Then
        If Not CDbl(TotalReservado.Caption) > 0 Then Error 23730
    Else
        Error 23731
    End If

    objGenerico.vVariavel = 0

    'Verificar se o Total Reservado é igual a Quantidade a reservar, se não for --> erro
    If TotalReservado.Caption <> QuantReservar.Caption Then
        Call Chama_Tela_Modal("AlocacaoProdutoSaida1", objGenerico)
        If giRetornoTela = vbOK Then
            lErro = Atualiza_Dados_Reserva(objGenerico)
            If lErro <> SUCESSO Then Error 23935
            
            iAlterado = 0
            Gravar_Registro = SUCESSO
        Else
            Gravar_Registro = 23709
        End If

        GL_objMDIForm.MousePointer = vbDefault
        Exit Function

    End If
    
    lErro = Move_Tela_Memoria(gobjItemPedido)
    If lErro <> SUCESSO Then Error 23710

    giRetornoTela = vbOK
    iAlterado = 0

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 23710, 23935

        Case 23730, 23731
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TOTAL_RESERVADO_SEM_PREENCHIMENTO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142734)

    End Select

    Exit Function

End Function

Public Function Move_Tela_Memoria(objItemPedido As ClassItemPedido) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim dFator As Double
Dim dtValidade As Date

On Error GoTo Erro_Move_Tela_Memoria

    Set objItemPedido.colReserva = New colReserva

    'Preenche colReservaItem com as linhas do GridReserva
    For iIndice = 1 To objGrid2.iLinhasExistentes
    
        'Se a quantidade reservada for positiva
        If Len(Trim(GridReserva.TextMatrix(iIndice, iGrid_QuantReservada_Col))) > 0 Then
            If CDbl(GridReserva.TextMatrix(iIndice, iGrid_QuantReservada_Col)) > 0 Then
                
                If Len(Trim(GridReserva.TextMatrix(iIndice, iGrid_DataValidade_Col))) > 0 Then
                    dtValidade = CDate(GridReserva.TextMatrix(iIndice, iGrid_DataValidade_Col))
                Else
                    dtValidade = DATA_NULA
                End If
                
                objItemPedido.colReserva.Add 0, 0, "", gcolEstoque(iIndice).iAlmoxarifado, 0, 0, 0, CDbl(GridReserva.TextMatrix(iIndice, iGrid_QuantReservada_Col)), DATA_NULA, dtValidade, "", GridReserva.TextMatrix(iIndice, iGrid_Responsavel_Col), 0, gcolEstoque(iIndice).sAlmoxarifadoNomeReduzido
            
            End If
        End If

    Next

    'Calcula o Fator de conversão de UM Estoque para Venda
    lErro = CF("UM_Conversao",objItemPedido.iClasseUM, objItemPedido.sUMEstoque, objItemPedido.sUnidadeMed, dFator)
    If lErro <> SUCESSO Then Error 23779

    gobjItemPedido.dQuantReservada = StrParaDbl(TotalReservado.Caption) * dFator

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case 23779

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142735)

    End Select

    Exit Function

End Function

Private Sub Limpa_Tela_Alocacao()

Dim iIndice As Integer

    'Limpa Colunas do GridReserva (QuantReservada, Data de Validade e Responsável)
    For iIndice = 1 To objGrid2.iLinhasExistentes
        GridReserva.TextMatrix(iIndice, iGrid_QuantReservada_Col) = ""
        GridReserva.TextMatrix(iIndice, iGrid_Responsavel_Col) = ""
        GridReserva.TextMatrix(iIndice, iGrid_DataValidade_Col) = ""
    Next
    
    'Limpa label TotalReservado
    TotalReservado.Caption = Formata_Estoque(0)

End Sub


Private Function Preenche_Tela(objItemPedido As ClassItemPedido, gcolEstoque As colEstoqueProduto) As Long
'Traz os dados do Item Pedido para a Tela

Dim lErro As Long
Dim iLinha As Integer
Dim dFator As Double
Dim sProdutoMascarado As String
Dim objEstoqueProduto As ClassEstoqueProduto

On Error GoTo Erro_Preenche_Tela

    'Preenche dados do Item Pedido
    
    lErro = Mascara_MascararProduto(objItemPedido.sProduto, sProdutoMascarado)
    If lErro <> SUCESSO Then Error 23802
    
    Produto.Caption = sProdutoMascarado
    Descricao.Caption = objItemPedido.sProdutoDescricao
    UnidadeMedida.Caption = objItemPedido.sUMEstoque

    'Calcular o fator de conversão de UM de Venda para Estoque
    lErro = CF("UM_Conversao",objItemPedido.iClasseUM, objItemPedido.sUnidadeMed, objItemPedido.sUMEstoque, dFator)
    If lErro <> SUCESSO Then Error 23778

    'Preenche QuantReservar
    QuantReservar.Caption = Formata_Estoque((objItemPedido.dQuantidade - objItemPedido.dQuantCancelada - objItemPedido.dQuantFaturada) * dFator)

    'Preenche GridReserva
    For Each objEstoqueProduto In gcolEstoque

        iLinha = iLinha + 1

        GridReserva.TextMatrix(iLinha, iGrid_Almoxarifado_Col) = objEstoqueProduto.sAlmoxarifadoNomeReduzido
        GridReserva.TextMatrix(iLinha, iGrid_QuantDisponivel_Col) = Formata_Estoque(objEstoqueProduto.dSaldo)

    Next

    Preenche_Tela = SUCESSO

    Exit Function

Erro_Preenche_Tela:

    Preenche_Tela = Err

    Select Case Err

        Case 23778, 23802

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142736)

    End Select

    Exit Function

End Function


Private Function SomaReservado(dTotalReservado As Double)
'Soma a coluna de quantidade reservada e coloca o total em TotalReservado

Dim iIndice As Integer

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



Private Function Saida_Celula_DataValidade(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dtDataValidade As Date

On Error GoTo Erro_Saida_Celula_DataValidade

    Set objGridInt.objControle = DataValidade

    'Verifica se Data de Validade esta preenchida
    If Len(Trim(DataValidade.ClipText)) > 0 Then

        'Critica a data
        lErro = Data_Critica(DataValidade.Text)
        If lErro <> SUCESSO Then Error 23734

         dtDataValidade = CDate(DataValidade.Text)

        'Se data de Validade é menor que a Data Corrente
        If dtDataValidade < Date Then Error 23735

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 23736

    Saida_Celula_DataValidade = SUCESSO

    Exit Function

Erro_Saida_Celula_DataValidade:

    Saida_Celula_DataValidade = Err

    Select Case Err

        Case 23734, 23736
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 23735
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVALIDADE_MENOR", Err, DataValidade.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142737)

    End Select

    Exit Function

End Function

Function Atualiza_Dados_Reserva(objGenerico As AdmGenerico) As Long

Dim lErro As Long
Dim dFator As Double
Dim dQuantCancelada As Double

On Error GoTo Erro_Atualiza_Dados_Reserva

    Select Case objGenerico.vVariavel
    
        Case NENHUMA_SELECAO
        
        Case SELECAO_OK
            lErro = Move_Tela_Memoria(gobjItemPedido)
            If lErro <> SUCESSO Then Error 23935
            
        Case CANCELA_ACIMA_DA_RESERVADA
            lErro = Move_Tela_Memoria(gobjItemPedido)
            If lErro <> SUCESSO Then Error 23936
            
            lErro = CF("UM_Conversao",gobjItemPedido.iClasseUM, gobjItemPedido.sUMEstoque, gobjItemPedido.sUnidadeMed, dFator)
            If lErro <> SUCESSO Then Error 23937
            
            dQuantCancelada = Formata_Estoque((gobjItemPedido.dQuantidade - gobjItemPedido.dQuantFaturada) / dFator) - StrParaDbl(QuantReservar.Caption)
            
            gobjItemPedido.dQuantCancelada = Formata_Estoque((dQuantCancelada + StrParaDbl(QuantReservar.Caption) - CDbl(TotalReservado.Caption)) * dFator)
        
    End Select
    
    Atualiza_Dados_Reserva = SUCESSO
    
    Exit Function
    
Erro_Atualiza_Dados_Reserva:

    Atualiza_Dados_Reserva = Err
    
    Select Case Err
    
        Case 23935, 23936, 23937
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142738)

    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object
    
    Parent.HelpContextID = IDH_ALOCACAO_PRODUTO1
    Set Form_Load_Ocx = Me
    Caption = "Reserva de Produto"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "AlocacaoProduto1"
    
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
    lErro = CF("UM_Conversao",iClassUM, sUMVenda, sUMEstoque, dFator)
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142739)
            
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
        lErro = Calcula_Tolerancia(gobjItemPedido.iClasseUM, gobjItemPedido.sUnidadeMed, UnidadeMedida.Caption, dTolerancia)
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142740)
    
    End Select
    
    Exit Function
    
End Function


Private Sub QuantTotalReserva_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantTotalReserva, Source, X, Y)
End Sub

Private Sub QuantTotalReserva_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantTotalReserva, Button, Shift, X, Y)
End Sub

Private Sub TotalReservado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalReservado, Source, X, Y)
End Sub

Private Sub TotalReservado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalReservado, Button, Shift, X, Y)
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

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

