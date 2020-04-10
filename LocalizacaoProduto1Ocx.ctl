VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl LocalizacaoProduto1Ocx 
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6750
   LockControls    =   -1  'True
   ScaleHeight     =   5685
   ScaleWidth      =   6750
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4920
      ScaleHeight     =   495
      ScaleWidth      =   1560
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   1620
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1065
         Picture         =   "LocalizacaoProduto1Ocx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   570
         Picture         =   "LocalizacaoProduto1Ocx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   75
         Picture         =   "LocalizacaoProduto1Ocx.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
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
      Left            =   4110
      TabIndex        =   4
      Top             =   5220
      Width           =   2430
   End
   Begin VB.Frame Frame7 
      Caption         =   "Localização do Produto"
      Height          =   2760
      Left            =   120
      TabIndex        =   9
      Top             =   2265
      Width           =   6420
      Begin MSMask.MaskEdBox Almoxarifado 
         Height          =   225
         Left            =   1065
         TabIndex        =   0
         Top             =   510
         Width           =   1590
         _ExtentX        =   2805
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
         Left            =   2835
         TabIndex        =   1
         Top             =   555
         Width           =   1440
         _ExtentX        =   2540
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
      Begin MSMask.MaskEdBox QuantAlocada 
         Height          =   225
         Left            =   4485
         TabIndex        =   2
         Top             =   510
         Width           =   1440
         _ExtentX        =   2540
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
      Begin MSFlexGridLib.MSFlexGrid GridAlocacao 
         Height          =   1860
         Left            =   390
         TabIndex        =   3
         Top             =   390
         Width           =   5625
         _ExtentX        =   9922
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
         Caption         =   "Quant. Alocada:"
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
         Left            =   1845
         TabIndex        =   11
         Top             =   2295
         Width           =   1395
      End
      Begin VB.Label TotalAlocado 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3330
         TabIndex        =   10
         Top             =   2265
         Width           =   1440
      End
   End
   Begin VB.Label UnidadeMedida 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1830
      TabIndex        =   20
      Top             =   1305
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
      Left            =   960
      TabIndex        =   19
      Top             =   1335
      Width           =   780
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Quant. a Alocar:"
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
      Left            =   330
      TabIndex        =   18
      Top             =   1785
      Width           =   1410
   End
   Begin VB.Label QuantAlocar 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1830
      TabIndex        =   17
      Top             =   1740
      Width           =   1440
   End
   Begin VB.Label Descricao 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3420
      TabIndex        =   16
      Top             =   840
      Width           =   3075
   End
   Begin VB.Label Produto 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1860
      TabIndex        =   15
      Top             =   840
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
      Left            =   1050
      TabIndex        =   14
      Top             =   855
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
      Left            =   1350
      TabIndex        =   13
      Top             =   375
      Width           =   435
   End
   Begin VB.Label Item 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1860
      TabIndex        =   12
      Top             =   375
      Width           =   660
   End
End
Attribute VB_Name = "LocalizacaoProduto1Ocx"
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
Dim gobjItemNF As ClassItemNF
Dim gcolOutrosProdutos As Collection
Dim giCodigo As Integer
'Dim dTotalReservada as Double

Dim objGridAlocacao As AdmGrid
Dim iGrid_Almoxarifado_Col As Integer
Dim iGrid_QuantDisp_Col As Integer
Dim iGrid_QuantAloc_Col As Integer

Private Sub Almoxarifado_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub BotaoFechar_Click()
    
    Unload Me

End Sub
Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 39395

    'Limpa a Tela
    Call Limpa_Tela_Alocacao

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 39395

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162389)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

    'Inicializa o formato de estoque
    QuantDisponivel.Format = FORMATO_ESTOQUE
    QuantAlocada.Format = FORMATO_ESTOQUE
    
    giRetornoTela = vbCancel
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
    
End Sub

Public Sub Form_Unload(Cancel As Integer)

    'Liberar as variaveis globais
    Set gobjItemNF = Nothing
    Set gcolOutrosProdutos = Nothing

    Set objGridAlocacao = Nothing

End Sub

Private Sub QuantAlocada_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub
Private Sub QuantAlocada_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridAlocacao)

End Sub

Private Sub QuantAlocada_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAlocacao)

End Sub

Private Sub QuantAlocada_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAlocacao.objControle = QuantAlocada
    lErro = Grid_Campo_Libera_Foco(objGridAlocacao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantDisponivel_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Function Trata_Parametros(objItemNF As ClassItemNF, ByVal colOutrosProdutos, ByVal dQuantAlocar As Double, ByVal iCodigo As Integer) As Long

Dim lErro As Long
Dim objEstoqueProduto As ClassEstoqueProduto
Dim iIndice As Integer
Dim colEstoqueProduto As New colEstoqueProduto
Dim objItemPV As New ClassItemPedido
Dim colReserva As New colReservaItem

On Error GoTo Erro_Trata_Parametros

    giCodigo = iCodigo
    
    'Atribui os dados do obj para o objGlobal da tela
    Set gobjItemNF = objItemNF
    'Atribui a coleção passada por prâmetro para a coleção da tela
    Set gcolOutrosProdutos = colOutrosProdutos
    'Lê os Estoque do produto nessa filialempresa
    lErro = CF("EstoquesProduto_Le_Filial", objItemNF.sProduto, colEstoqueProduto)
    If lErro <> SUCESSO Then Error 39349
    
    'Se o itemNF foi gerado por um itemPV
    If objItemNF.lNumIntItemPedVenda > 0 Then
        objItemPV.lNumIntDoc = objItemNF.lNumIntItemPedVenda
        objItemPV.sProduto = objItemNF.sProduto
        'Lê as reservas do item do pedido
        lErro = CF("ReservasItemPV_Le_NumIntOrigem", objItemPV, colReserva)
        If lErro <> SUCESSO And lErro <> 51601 Then Error 51596

    End If
    
    iIndice = 0
    
    'Exclui almoxarifados sem estoque disponível
    For Each objEstoqueProduto In colEstoqueProduto
        iIndice = iIndice + 1
        'Verifica se há uma reserva para esse item nesse almoxarifado
        'e inclui a qtd reservada como disponivel p\ esse pedido
        Call Procura_Almoxarifado(objEstoqueProduto, colReserva)
        'SE não tiver quantidade disponível e nem reserva nesse almoxarifado
        
        Select Case iCodigo
        
            Case DOCINFO_NFISPC, DOCINFO_NFFISPC
        
                If objEstoqueProduto.dQuantConsig <= 0 And gobjMAT.iAceitaEstoqueNegativo = DESMARCADO Then
                    'Retira o estoque produto da coleção
                    colEstoqueProduto.Remove (iIndice)
                    iIndice = iIndice - 1
                End If
        
            Case DOCINFO_NFISBF, DOCINFO_NFISFBF, DOCINFO_NFISRMB3PV 'Inserido por Leo em 15/01/02
            
                If objEstoqueProduto.dQuantBenef3 <= 0 And gobjMAT.iAceitaEstoqueNegativo = DESMARCADO Then
                    'Retira o estoque produto da coleção
                    colEstoqueProduto.Remove (iIndice)
                    iIndice = iIndice - 1
                End If
            
            Case Else
        
                If objEstoqueProduto.dQuantDisponivel <= 0 And gobjMAT.iAceitaEstoqueNegativo = DESMARCADO Then
                    'Retira o estoque produto da coleção
                    colEstoqueProduto.Remove (iIndice)
                    iIndice = iIndice - 1
                End If
        
        End Select
        
    Next
    
    'Se não tiver nenhum estoque -->Erro
    'If colEstoqueProduto.Count = 0 Then Error 39351 'Não pode dar erro quando não tem saldo pois tem que abrir a tela e deixar disponível a opção de substituir o produto
    
    Set objGridAlocacao = New AdmGrid
        
    lErro = Inicializa_Grid_Alocacao(objGridAlocacao, colEstoqueProduto, iCodigo)
    If lErro <> SUCESSO Then Error 39352
    'Preenche a tela com os dados do estoque desse produto
    lErro = Preenche_Tela(objItemNF, colEstoqueProduto, dQuantAlocar, iCodigo)
    If lErro <> SUCESSO Then Error 39353
    
    If objItemNF.iPossuiGrade Then BotaoSubstituir.Visible = False
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = Err
    
    Select Case Err
    
        Case 39349, 39352, 39353, 51596
        
        Case 39351
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NAO_EXISTE_ESTOQUE", Err, objItemNF.sProduto)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162390)
            
    End Select
    
    iAlterado = 0
        
    Exit Function

End Function

Private Function Inicializa_Grid_Alocacao(objGridInt As AdmGrid, colEstoqueProduto As colEstoqueProduto, iCodigo As Integer) As Long
'Inicializa o Grid de Alocação

Dim iIndice As Integer

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Almoxarifado")
    
    If iCodigo = DOCINFO_NFISPC Or iCodigo = DOCINFO_NFFISPC Then
        objGridInt.colColuna.Add ("Quant. Consignada")
    
    ElseIf iCodigo = DOCINFO_NFISRMB3PV Then 'Inserido por Leo em 15/01/02
        objGridInt.colColuna.Add ("Quant. Benef. 3º")
    
    Else
        objGridInt.colColuna.Add ("Quant. Disponivel")
    
    End If
    
    objGridInt.colColuna.Add ("Quant. Alocada")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Almoxarifado.Name)
    objGridInt.colCampo.Add (QuantDisponivel.Name)
    objGridInt.colCampo.Add (QuantAlocada.Name)

    'Colunas da Grid
    iGrid_Almoxarifado_Col = 1
    iGrid_QuantDisp_Col = 2
    iGrid_QuantAloc_Col = 3

    'Grid do GridInterno
    objGridInt.objGrid = GridAlocacao

    objGridInt.iLinhasVisiveis = 7
    If colEstoqueProduto.Count > 7 Then
        objGridInt.objGrid.Rows = colEstoqueProduto.Count + 1
    Else
        objGridInt.objGrid.Rows = 8
    End If

    'Largura da primeira coluna
    GridAlocacao.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    'Posiciona o totalizador e seu label
    TotalAlocado.top = GridAlocacao.top + GridAlocacao.Height
    TotalAlocado.left = GridAlocacao.left
    For iIndice = 0 To iGrid_QuantAloc_Col - 1
        TotalAlocado.left = TotalAlocado.left + GridAlocacao.ColWidth(iIndice) + GridAlocacao.GridLineWidth + 20
    Next

    TotalAlocado.Width = GridAlocacao.ColWidth(iGrid_QuantAloc_Col)

    QuantTotalReserva.top = TotalAlocado.top + (TotalAlocado.Height - QuantTotalReserva.Height) / 2
    QuantTotalReserva.left = TotalAlocado.left - QuantTotalReserva.Width

    Inicializa_Grid_Alocacao = SUCESSO

    Exit Function

End Function

Private Function Preenche_Tela(objItemNF As ClassItemNF, colEstoqueProduto As colEstoqueProduto, dQuantAlocar As Double, iCodigo As Integer) As Long

Dim lErro As Long
Dim sProduto As String
Dim objEstoqueProduto As ClassEstoqueProduto
Dim iIndice As Integer

On Error GoTo Erro_Preenche_Tela

    'Coloca o item na tela
    Item.Caption = objItemNF.iItem
    'Mascara o produto
    lErro = Mascara_MascararProduto(objItemNF.sProduto, sProduto)
    If lErro <> SUCESSO Then Error 39354
    'Coloca os dados do produto na tela
    Produto.Caption = sProduto
    Descricao.Caption = objItemNF.sDescricaoItem
    UnidadeMedida.Caption = objItemNF.sUMEstoque
    QuantAlocar.Caption = Formata_Estoque(dQuantAlocar)
    'Preenche o grid com os estoques do produto
    For Each objEstoqueProduto In colEstoqueProduto
        iIndice = iIndice + 1
        GridAlocacao.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = objEstoqueProduto.sAlmoxarifadoNomeReduzido
        
        Select Case giCodigo
        
            Case DOCINFO_NFISPC, DOCINFO_NFFISPC
            
                GridAlocacao.TextMatrix(iIndice, iGrid_QuantDisp_Col) = Formata_Estoque(objEstoqueProduto.dQuantConsig)
                
            Case DOCINFO_NFISBF, DOCINFO_NFISFBF, DOCINFO_NFISRMB3PV 'Inluido a constante DOCINFO_NFISRMB3PV por Leo em 15/01/02
                
                GridAlocacao.TextMatrix(iIndice, iGrid_QuantDisp_Col) = Formata_Estoque(objEstoqueProduto.dQuantBenef3)
                
            Case Else
                GridAlocacao.TextMatrix(iIndice, iGrid_QuantDisp_Col) = Formata_Estoque(objEstoqueProduto.dQuantDisponivel)
        
        End Select

    Next
    'Inicializa o número de linhas existentes
    objGridAlocacao.iLinhasExistentes = iIndice
    
    Preenche_Tela = SUCESSO
    
    Exit Function
    
Erro_Preenche_Tela:

    Preenche_Tela = Err
    
    Select Case Err
    
        Case 39354
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162391)
            
    End Select
    
    Exit Function
    
End Function
Private Sub GridAlocacao_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridAlocacao, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAlocacao, iAlterado)
    End If

End Sub

Private Sub GridAlocacao_EnterCell()

    Call Grid_Entrada_Celula(objGridAlocacao, iAlterado)

End Sub

Private Sub GridAlocacao_GotFocus()

    Call Grid_Recebe_Foco(objGridAlocacao)

End Sub

Private Sub GridAlocacao_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridAlocacao, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAlocacao, iAlterado)
    End If

End Sub

Private Sub GridAlocacao_LeaveCell()

    Call Saida_Celula(objGridAlocacao)

End Sub

Private Sub GridAlocacao_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridAlocacao)
    
End Sub

Private Sub GridAlocacao_RowColChange()

    Call Grid_RowColChange(objGridAlocacao)

End Sub

Private Sub GridAlocacao_Scroll()

    Call Grid_Scroll(objGridAlocacao)

End Sub


Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'Verifica qual a coluna do Grid em questão
        Select Case objGridInt.objGrid.Col
            'Quantidade Alocada
            Case iGrid_QuantAloc_Col
                lErro = Saida_Celula_QuantAlocada(objGridInt)
                If lErro <> SUCESSO Then Error 39355
        
        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 39356

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 39355

        Case 39356
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162392)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantAlocada(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantDisponivel As Double
Dim dQuantAlocada As Double
Dim dTotalAlocado As Double
Dim dQuantAlocar As Double

On Error GoTo Erro_Saida_Celula_QuantAlocada

    Set objGridInt.objControle = QuantAlocada
    
    'Verifica se a quant alocada está preenchida
    If Len(Trim(QuantAlocada.ClipText)) > 0 Then
        'Valida a qtd informada
        lErro = Valor_NaoNegativo_Critica(QuantAlocada.Text)
        If lErro <> SUCESSO Then Error 39357
        'Recolhe a quant alocada e a quant disponível da tela
        dQuantAlocada = StrParaDbl(QuantAlocada.Text)
        dQuantDisponivel = StrParaDbl(GridAlocacao.TextMatrix(GridAlocacao.Row, iGrid_QuantDisp_Col))
        'Verifica se a quant disponível é inferioe a quant alocada
        If gobjMAT.iAceitaEstoqueNegativo = DESMARCADO And dQuantDisponivel < dQuantAlocada Then Error 39358
        
        GridAlocacao.TextMatrix(GridAlocacao.Row, iGrid_QuantAloc_Col) = Formata_Estoque(dQuantAlocada)
        'Obtem o total alocado no grid
        Call SubTotal_Calcula(objGridAlocacao, iGrid_QuantAloc_Col, dTotalAlocado)
        
        dQuantAlocar = StrParaDbl(QuantAlocar.Caption)
        'Se o total alocado for superior a qtd alocar --> erro.
        If dTotalAlocado > dQuantAlocar Then Error 39359
        'Atuliza o total alocado da tela
        TotalAlocado.Caption = Formata_Estoque(dTotalAlocado)
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 39360
    
    Saida_Celula_QuantAlocada = SUCESSO

    Exit Function
    
Erro_Saida_Celula_QuantAlocada:

    Saida_Celula_QuantAlocada = Err
    
    Select Case Err
    
        Case 39357, 39360
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 39358
            lErro = Rotina_Erro(vbOKOnly + vbSystemModal, "ERRO_QUANT_ALOCADA_MAIOR_DISPONIVEL", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case 39359
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TOTAL_ALOCACAO_SUPERIOR_ALOCAR", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162393)
            
    End Select
 
    Exit Function
 
End Function

Private Sub BotaoSubstituir_Click()

Dim objSubstProdutoNF As New ClassSubstProdutoNF

    objSubstProdutoNF.sProduto = gobjItemNF.sProduto
    
    Set objSubstProdutoNF.colOutrosProdutosNF = gcolOutrosProdutos
    
    'Chama a tela de substituição de produtos
    Call Chama_Tela_Modal("SubstProdutoNF", objSubstProdutoNF, giCodigo)
    'Verifica se houve substituição
    If giRetornoTela = vbOK Then
        'Substitui o produto
        gobjItemNF.sProduto = objSubstProdutoNF.sProdutoSubstituto
        iAlterado = 0
        'Fecha a tela
        Unload Me
    End If

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 39361
    
    Unload Me

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 39361

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 162394)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_Alocacao()

Dim iIndice As Integer
    'Limpa o total alocado
    TotalAlocado.Caption = ""
    
    'Limpa as qts alocadas no grid
    For iIndice = 1 To objGridAlocacao.iLinhasExistentes
        GridAlocacao.TextMatrix(iIndice, iGrid_QuantAloc_Col) = ""
    Next
    
    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim dQuantAlocar As Double
Dim dTotalAlocado As Double
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    dTotalAlocado = StrParaDbl(TotalAlocado.Caption)
    dQuantAlocar = StrParaDbl(QuantAlocar.Caption)
    'Se o total alocado for diferente do total
    If dTotalAlocado <> dQuantAlocar Then
        'Avisa que o total alocado é diferente do total
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALOCADO_MENOR_ALOCAR", dTotalAlocado, dQuantAlocar)
        If vbMsgRes = vbNo Then Error 39362
        
    End If
    
    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(gobjItemNF)
    If lErro <> SUCESSO Then Error 39363
    'Retorna OK
    giRetornoTela = vbOK

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 39362, 39363
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162395)
            
    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(objItemNF As ClassItemNF) As Long

Dim dQuantAlocada As Double
Dim iIndice As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria
    'Limpa as alocaçoes do item
    Set objItemNF.colAlocacoes = New ColAlocacoesItemNF
    'Recolhe as alocacoes do grid de alocacoes
    For iIndice = 1 To objGridAlocacao.iLinhasExistentes
    
        dQuantAlocada = StrParaDbl(GridAlocacao.TextMatrix(iIndice, iGrid_QuantAloc_Col))
        If dQuantAlocada > 0 Then
            objAlmoxarifado.sNomeReduzido = GridAlocacao.TextMatrix(iIndice, iGrid_Almoxarifado_Col)
                    
            lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25060 Then Error 39364
            If lErro <> SUCESSO Then Error 39365
            'Adiciona alocacao na coleção de alocacoes
            objItemNF.colAlocacoes.Add objAlmoxarifado.iCodigo, objAlmoxarifado.sNomeReduzido, dQuantAlocada
        End If
    Next
            
    Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err
    
    Select Case Err
        
        Case 39364
        
        Case 39365
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", Err, objAlmoxarifado.sNomeReduzido)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162396)
    End Select
    
    Exit Function

End Function

Private Sub SubTotal_Calcula(objGridInt As AdmGrid, iGrid_Coluna As Integer, dSubtotal As Double)
'Faz a soma da Coluna no Grid passado e devolve em dValorTotal

Dim iIndice As Integer

    dSubtotal = 0

    For iIndice = 1 To objGridInt.iLinhasExistentes
        'Acumula em dSubTotal
        dSubtotal = dSubtotal + StrParaDbl(IIf(Len(Trim(objGridInt.objGrid.TextMatrix(iIndice, iGrid_Coluna))) > 0, objGridInt.objGrid.TextMatrix(iIndice, iGrid_Coluna), 0))

    Next

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_LOCALIZACAO_PRODUTO1
    Set Form_Load_Ocx = Me
    Caption = "Localização de Produto"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "LocalizacaoProduto1"
    
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
Private Sub Procura_Almoxarifado(objEstoqueProduto As ClassEstoqueProduto, colReserva As colReservaItem)
'Percorre uma colecao de reservas e verifica se o almoxarifado passado está na coleção.

Dim iIndice As Integer

    For iIndice = 1 To colReserva.Count
    
        If objEstoqueProduto.iAlmoxarifado = colReserva(iIndice).iAlmoxarifado Then
            objEstoqueProduto.dQuantDispNossa = objEstoqueProduto.dQuantDispNossa + colReserva(iIndice).dQuantidade
            Exit For
        End If

    Next
    
    Exit Sub
    
End Sub


Private Sub QuantTotalReserva_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantTotalReserva, Source, X, Y)
End Sub

Private Sub QuantTotalReserva_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantTotalReserva, Button, Shift, X, Y)
End Sub

Private Sub TotalAlocado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalAlocado, Source, X, Y)
End Sub

Private Sub TotalAlocado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalAlocado, Button, Shift, X, Y)
End Sub

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

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub QuantAlocar_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantAlocar, Source, X, Y)
End Sub

Private Sub QuantAlocar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantAlocar, Button, Shift, X, Y)
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

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Item_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Item, Source, X, Y)
End Sub

Private Sub Item_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Item, Button, Shift, X, Y)
End Sub

